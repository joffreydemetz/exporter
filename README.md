# JDZ Exporter

Spreadsheet export (xlsx / csv / ods / pdf) built on [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet).

Feed associative rows, get a rendered multi-worksheet file — the first row's keys become a bold header row, cell rendering is driven by an optional per-cell `type` (string / bool / formula / text / phone / id / number / euro / date).

```php
use JDZ\Exporter\Exporter;

(new Exporter('My App'))
    ->setTitle('Users')
    ->setRows($rows, 'Users')
    ->render()
    ->save('/tmp/users.xlsx');
```

Writers: `xlsx` (default), `csv` (single sheet, `;` delimited), `ods`, `pdf` (TCPDF).

## Install

```
composer require jdz/exporter
```

## Tests

```
composer test
```
