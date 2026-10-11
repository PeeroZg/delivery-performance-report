<?php

declare(strict_types=1);

require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/*
|--------------------------------------------------------------------------
| Configuration
|--------------------------------------------------------------------------
*/

$databaseConfig = require __DIR__ . '/config/database.php';

$outputDirectory = __DIR__ . '/output';

/*
|--------------------------------------------------------------------------
| Database connection
|--------------------------------------------------------------------------
*/

try {
	$pdo = new PDO(
		sprintf(
			'pgsql:host=%s;port=%s;dbname=%s',
			$databaseConfig['host'],
			$databaseConfig['port'],
			$databaseConfig['database']
		),
		$databaseConfig['username'],
		$databaseConfig['password'],
		[
			PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
			PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
		]
	);
} catch (PDOException $exception) {
	fwrite(
		STDERR,
		sprintf(
			"Database connection failed: %s\n",
			$exception->getMessage()
		)
	);

	exit(1);
}

/*
|--------------------------------------------------------------------------
| Report date
|--------------------------------------------------------------------------
|
| The report covers parcels picked up on the given date. If no date is
| passed as a CLI argument, we default to "yesterday" (the previous
| calendar day relative to today).
*/

$reportDate = $argv[1] ?? date('Y-m-d', strtotime('-1 day'));

$dateCheck = DateTime::createFromFormat('Y-m-d', $reportDate);

if (!$dateCheck || $dateCheck->format('Y-m-d') !== $reportDate) {
	fwrite(
		STDERR,
		sprintf(
			"Invalid report date \"%s\". Expected format: YYYY-MM-DD\n",
			$reportDate
		)
	);

	exit(1);
}

/*
|--------------------------------------------------------------------------
| Fetch report data
|--------------------------------------------------------------------------
|
| A single query retrieves, for every parcel picked up on the report date:
|   - parcel + customer details
|   - the final/current status (latest event by event_time)
|   - the timestamp of that latest event
|   - the number of delivery attempts (OUT_FOR_DELIVERY events)
|
| Doing this in one query (instead of looping over parcels in PHP and
| issuing per-parcel queries) avoids N+1 database round-trips.
*/

$sql = <<<SQL
	WITH report_parcels AS (
		SELECT
			p.id,
			p.parcel_number,
			c.customer_number,
			c.name AS customer_name,
			p.receiver_name,
			p.receiver_zip,
			p.receiver_city,
			p.cod_amount,
			p.pickup_date
		FROM parcels p
		INNER JOIN customers c ON c.id = p.customer_id
		WHERE p.pickup_date = :report_date
	),
	ranked_events AS (
		SELECT
			e.parcel_id,
			e.event_code,
			e.event_time,
			ROW_NUMBER() OVER (
				PARTITION BY e.parcel_id
				ORDER BY e.event_time DESC, e.id DESC
			) AS rn
		FROM parcel_events e
		INNER JOIN report_parcels rp ON rp.id = e.parcel_id
	),
	final_status AS (
		SELECT
			parcel_id,
			event_code AS final_status,
			event_time AS last_event_time
		FROM ranked_events
		WHERE rn = 1
	),
	delivery_attempts AS (
		SELECT
			e.parcel_id,
			COUNT(*) AS attempts
		FROM parcel_events e
		INNER JOIN report_parcels rp ON rp.id = e.parcel_id
		WHERE e.event_code = 'OUT_FOR_DELIVERY'
		GROUP BY e.parcel_id
	)
	SELECT
		rp.parcel_number,
		rp.customer_number,
		rp.customer_name,
		rp.receiver_name,
		rp.receiver_zip,
		rp.receiver_city,
		rp.cod_amount,
		rp.pickup_date,
		COALESCE(da.attempts, 0) AS delivery_attempts,
		COALESCE(fs.final_status, 'NO_EVENTS') AS final_status,
		fs.last_event_time
	FROM report_parcels rp
	LEFT JOIN final_status fs ON fs.parcel_id = rp.id
	LEFT JOIN delivery_attempts da ON da.parcel_id = rp.id
	ORDER BY rp.parcel_number ASC
	SQL;

try {
	$statement = $pdo->prepare($sql);
	$statement->execute(['report_date' => $reportDate]);
	$parcels = $statement->fetchAll();
} catch (PDOException $exception) {
	fwrite(
		STDERR,
		sprintf(
			"SQL execution failed: %s\n",
			$exception->getMessage()
		)
	);

	exit(1);
}

/*
|--------------------------------------------------------------------------
| Calculate summary statistics
|--------------------------------------------------------------------------
*/

$totalParcels = count($parcels);
$delivered = 0;
$returned = 0;
$deliveredFirstAttempt = 0;
$deliveredSecondAttempt = 0;
$codParcels = 0;
$codAmount = 0.0;

foreach ($parcels as $parcel) {
	if ($parcel['final_status'] === 'DELIVERED') {
		$delivered++;

		if ((int) $parcel['delivery_attempts'] === 1) {
			$deliveredFirstAttempt++;
		} elseif ((int) $parcel['delivery_attempts'] === 2) {
			$deliveredSecondAttempt++;
		}
	} elseif ($parcel['final_status'] === 'RETURNED') {
		$returned++;
	}

	if ((float) $parcel['cod_amount'] > 0) {
		$codParcels++;
		$codAmount += (float) $parcel['cod_amount'];
	}
}

// "Not delivered" = everything that is neither successfully delivered
// nor returned (still in transit, failed attempts, no events, etc.).
$notDelivered = $totalParcels - $delivered - $returned;

$successRate = $totalParcels > 0
	? round(($delivered / $totalParcels) * 100, 2)
	: 0.0;

/*
|--------------------------------------------------------------------------
| Build the Excel report
|--------------------------------------------------------------------------
*/

try {
	$spreadsheet = new Spreadsheet();

	/*
	|----------------------------------------------------------------
	| Summary sheet
	|----------------------------------------------------------------
	*/

	$summarySheet = $spreadsheet->getActiveSheet();
	$summarySheet->setTitle('Summary');

	$summarySheet->setCellValue('A1', 'Delivery Performance Report');
	$summarySheet->setCellValue('A2', 'Report date (pickup):');
	$summarySheet->setCellValue('B2', $reportDate);
	$summarySheet->getStyle('A1')->getFont()->setBold(true)->setSize(14);
	$summarySheet->getStyle('A2')->getFont()->setBold(true);

	$summaryRows = [
		['Metric', 'Value'],
		['Total parcels', $totalParcels],
		['Delivered', $delivered],
		['Not delivered', $notDelivered],
		['Returned', $returned],
		['Delivery success rate', $successRate / 100],
		['Delivered on first attempt', $deliveredFirstAttempt],
		['Delivered on second attempt', $deliveredSecondAttempt],
		['COD parcels', $codParcels],
		['COD amount', $codAmount],
	];

	$startRow = 4;

	foreach ($summaryRows as $index => $row) {
		$rowNumber = $startRow + $index;

		$summarySheet->setCellValue("A{$rowNumber}", $row[0]);
		$summarySheet->setCellValue("B{$rowNumber}", $row[1]);
	}

	// Header row styling
	$summarySheet->getStyle("A{$startRow}:B{$startRow}")
		->getFont()->setBold(true);

	// Delivery success rate as a percentage (row: startRow + 5)
	$successRateRow = $startRow + 5;
	$summarySheet->getStyle("B{$successRateRow}")
		->getNumberFormat()
		->setFormatCode('0.00%');

	// COD amount as currency (row: startRow + 9)
	$codAmountRow = $startRow + 9;
	$summarySheet->getStyle("B{$codAmountRow}")
		->getNumberFormat()
		->setFormatCode('#,##0.00');

	$summarySheet->getColumnDimension('A')->setWidth(30);
	$summarySheet->getColumnDimension('B')->setWidth(20);

	/*
	|----------------------------------------------------------------
	| Parcels sheet
	|----------------------------------------------------------------
	*/

	$parcelsSheet = $spreadsheet->createSheet();
	$parcelsSheet->setTitle('Parcels');

	$parcelHeaders = [
		'Parcel number',
		'Customer number',
		'Customer',
		'Receiver name',
		'Receiver ZIP',
		'Receiver city',
		'COD amount',
		'Pickup date',
		'Delivery attempts',
		'Final status',
		'Last event time',
	];

	// Column letters for the 11 Parcels-sheet columns (A..K).
	$parcelColumns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'];

	foreach ($parcelHeaders as $columnIndex => $header) {
		$parcelsSheet->setCellValue(
			$parcelColumns[$columnIndex] . '1',
			$header
		);
	}

	$parcelsSheet->getStyle('A1:K1')->getFont()->setBold(true);

	// Parcel/customer numbers are text (VARCHAR) in the DB, but they look
	// like numbers (e.g. "1755000000123"). Force them to be stored as
	// explicit strings so Excel doesn't reinterpret them as numbers and
	// display them in scientific notation (e.g. "1E+13").
	$parcelsSheet->getStyle('A:B')->getNumberFormat()->setFormatCode(
		\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_TEXT
	);

	$rowNumber = 2;

	foreach ($parcels as $parcel) {
		$parcelsSheet->setCellValueExplicit(
			"A{$rowNumber}",
			$parcel['parcel_number'],
			\PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING
		);
		$parcelsSheet->setCellValueExplicit(
			"B{$rowNumber}",
			$parcel['customer_number'],
			\PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING
		);
		$parcelsSheet->setCellValue("C{$rowNumber}", $parcel['customer_name']);
		$parcelsSheet->setCellValue("D{$rowNumber}", $parcel['receiver_name']);
		$parcelsSheet->setCellValue("E{$rowNumber}", $parcel['receiver_zip']);
		$parcelsSheet->setCellValue("F{$rowNumber}", $parcel['receiver_city']);
		$parcelsSheet->setCellValue("G{$rowNumber}", (float) $parcel['cod_amount']);
		$parcelsSheet->setCellValue("H{$rowNumber}", $parcel['pickup_date']);
		$parcelsSheet->setCellValue("I{$rowNumber}", (int) $parcel['delivery_attempts']);
		$parcelsSheet->setCellValue("J{$rowNumber}", $parcel['final_status']);
		$parcelsSheet->setCellValue("K{$rowNumber}", $parcel['last_event_time'] ?? '');

		$rowNumber++;
	}

	// Number/date formatting for the data rows
	$lastRow = $rowNumber - 1;

	if ($lastRow >= 2) {
		$parcelsSheet->getStyle("G2:G{$lastRow}")
			->getNumberFormat()->setFormatCode('#,##0.00');

		$parcelsSheet->getStyle("H2:H{$lastRow}")
			->getNumberFormat()->setFormatCode('YYYY-MM-DD');

		$parcelsSheet->getStyle("K2:K{$lastRow}")
			->getNumberFormat()->setFormatCode('YYYY-MM-DD HH:MM:SS');
	}

	// Auto-size columns based on content
	foreach (range('A', 'K') as $columnLetter) {
		$parcelsSheet->getColumnDimension($columnLetter)
			->setAutoSize(true);
	}

	$spreadsheet->setActiveSheetIndex(0);
} catch (Throwable $exception) {
	fwrite(
		STDERR,
		sprintf(
			"Failed to build Excel report: %s\n",
			$exception->getMessage()
		)
	);

	exit(1);
}

/*
|--------------------------------------------------------------------------
| Save the report
|--------------------------------------------------------------------------
*/

if (!is_dir($outputDirectory) && !mkdir($outputDirectory, 0755, true) && !is_dir($outputDirectory)) {
	fwrite(
		STDERR,
		sprintf(
			"Output directory \"%s\" does not exist and could not be created.\n",
			$outputDirectory
		)
	);

	exit(1);
}

if (!is_writable($outputDirectory)) {
	fwrite(
		STDERR,
		sprintf(
			"Output directory \"%s\" is not writable.\n",
			$outputDirectory
		)
	);

	exit(1);
}

$outputPath = sprintf('%s/delivery_report_%s.xlsx', $outputDirectory, $reportDate);

try {
	$writer = new Xlsx($spreadsheet);
	$writer->save($outputPath);
} catch (Throwable $exception) {
	fwrite(
		STDERR,
		sprintf(
			"Failed to save Excel report: %s\n",
			$exception->getMessage()
		)
	);

	exit(1);
}

fwrite(
	STDOUT,
	sprintf(
		"Report generated successfully: %s (%d parcels)\n",
		$outputPath,
		$totalParcels
	)
);

exit(0);