# Delivery Performance Report

PHP command-line reporting application developed as a technical recruitment assignment. It reads parcel and tracking data from PostgreSQL and generates a daily Excel report.

## Tech stack

- PHP 8.2+, PDO / PostgreSQL
- SQL (CTEs and window functions)
- Composer and PhpSpreadsheet

## Features

- Report for parcels collected on a chosen date (defaults to the previous day)
- Latest parcel status determined by the tracking event timestamp
- Delivery-attempt counts and first-/second-attempt delivery statistics
- Delivery success rate, returns and cash-on-delivery (COD) totals
- Excel workbook with `Summary` and `Parcels` worksheets
- One main SQL query to avoid per-parcel database queries

## Run locally

Requires PHP 8.2+, `pdo_pgsql`, Composer and PostgreSQL.

1. Create a PostgreSQL database named `delivery_report`.
2. Run `database/schema.sql`, then `database/seed.sql`.
3. Set local database environment variables `DB_HOST`, `DB_PORT`, `DB_DATABASE`, `DB_USERNAME`, `DB_PASSWORD`, or use the defaults shown in `config/database.php` for local development only.
4. Install dependencies and generate a report:

   ```bash
   composer install
   php generate_report.php 2026-08-10
   ```

   To report on the previous day instead:

   ```bash
   php generate_report.php
   ```

The generated file is saved in `output/delivery_report_YYYY-MM-DD.xlsx`.

## Project structure

- `generate_report.php` — SQL retrieval, statistics and Excel export
- `config/database.php` — environment-based database configuration
- `database/` — schema and sample data
- `output/` — generated Excel files (not committed)

## Notes

This is a local technical assignment project using sample data. It is not a production reporting integration. Never commit real database passwords or generated reports containing personal information.
