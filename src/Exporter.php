<?php

/**
 * (c) Joffrey Demetz <joffrey.demetz@gmail.com>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace JDZ\Exporter;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Multi-worksheet spreadsheet exporter (xlsx / csv / ods / pdf).
 *
 * @author Joffrey Demetz <joffrey.demetz@gmail.com>
 */
class Exporter
{
  protected Spreadsheet $spreadsheet;

  protected array $worksheets = [];

  protected string $writer = 'xlsx';
  protected bool $csv = false;
  protected bool $ods = false;

  public function __construct(string $creator = 'JDZ')
  {
    $this->spreadsheet = new Spreadsheet();

    $this->spreadsheet->getProperties()
      ->setCreator($creator)
      ->setLastModifiedBy($creator)
      ->setSubject('Export')
      ->setTitle('Export');
  }

  public function setCreator(string $creator): static
  {
    $this->spreadsheet->getProperties()
      ->setCreator($creator)
      ->setLastModifiedBy($creator);

    return $this;
  }

  public function setTitle(string $title): static
  {
    $this->spreadsheet->getProperties()
      ->setTitle($title);

    return $this;
  }

  public function setSubject(string $subject): static
  {
    $this->spreadsheet->getProperties()
      ->setSubject($subject);

    return $this;
  }

  public function setDescription(string $description): static
  {
    $this->spreadsheet->getProperties()
      ->setDescription($description);

    return $this;
  }

  public function withCsv(bool $csv = true): static
  {
    $this->csv = $csv;
    $this->ods = false;
    $this->writer = 'csv';
    return $this;
  }

  public function withOds(bool $ods = true): static
  {
    $this->ods = $ods;
    $this->csv = false;
    $this->writer = 'ods';
    return $this;
  }

  public function setWriter(string $writer): static
  {
    $this->writer = $writer;
    $this->ods = 'ods' === $writer;
    $this->csv = 'csv' === $writer;
    return $this;
  }

  public function getWriter(): string
  {
    return $this->writer;
  }

  /**
   * Fill a worksheet from associative rows — first row's keys become the header.
   */
  public function setRows(array $rows, string $worksheet = '', bool $canBeEmpty = false): static
  {
    if (empty($rows)) {
      if (true === $canBeEmpty) {
        return $this;
      }

      throw new \Exception('Nothing to export to ' . $worksheet);
    }

    if ('' === $worksheet) {
      $worksheet = 'Sheet';
    }

    $fields = array_keys($rows[0]);
    $this->addWorksheetRow((new Header())->setData($fields), $worksheet);

    foreach ($rows as $row) {
      $this->addWorksheetRow((new Line())->setData(array_values($row)), $worksheet);
    }

    return $this;
  }

  public function render(): static
  {
    $sheetI = 0;

    foreach ($this->worksheets as $name => $rows) {
      if (0 === $sheetI) {
        $worksheet = $this->spreadsheet->getSheet(0);
      } else {
        if (true === $this->csv) {
          // csv holds a single sheet — ignore the others
          break;
        }
        $worksheet = $this->spreadsheet->createSheet($sheetI);
      }

      $worksheet->setTitle($name);

      $rowI = 1;
      foreach ($rows as $row) {
        $row->setRowNumber($rowI);
        $row->toSpreadsheet($worksheet);
        $rowI++;
      }

      $sheetI++;
    }

    return $this;
  }

  public function save(string $path): void
  {
    $this->spreadsheet->setActiveSheetIndex(0);

    if ('csv' === $this->writer) {
      $writer = IOFactory::createWriter($this->spreadsheet, 'Csv')
        ->setDelimiter(';')
        ->setEnclosure('"')
        ->setSheetIndex(0);
    } elseif ('ods' === $this->writer) {
      $writer = IOFactory::createWriter($this->spreadsheet, 'Ods');
    } elseif ('pdf' === $this->writer) {
      $writer = IOFactory::createWriter($this->spreadsheet, 'Tcpdf');
    } else {
      $writer = IOFactory::createWriter($this->spreadsheet, 'Xlsx');
    }

    $writer->save($path);
  }

  public function output(string $filename = 'export.xlsx'): static
  {
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="' . $filename . '"');
    header('Cache-Control: max-age=0');
    header('Cache-Control: max-age=1');
    header('Expires: Mon, 26 Jul 1997 05:00:00 GMT');
    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
    header('Cache-Control: cache, must-revalidate');
    header('Pragma: public');

    $this->save('php://output');

    return $this;
  }

  public function getSpreadsheet(): Spreadsheet
  {
    return $this->spreadsheet;
  }

  protected function addWorksheetRow(Row $row, ?string $worksheet = null): static
  {
    if (!$worksheet) {
      $worksheet = 'Worksheet';
    }

    if (!isset($this->worksheets[$worksheet])) {
      $this->worksheets[$worksheet] = [];
    }

    $this->worksheets[$worksheet][] = $row;
    return $this;
  }
}
