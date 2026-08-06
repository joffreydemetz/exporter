<?php

/**
 * (c) Joffrey Demetz <joffrey.demetz@gmail.com>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace JDZ\Exporter;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Spreadsheet row — cells are assigned to columns in insertion order.
 *
 * @author Joffrey Demetz <joffrey.demetz@gmail.com>
 */
abstract class Row
{
  /** @var \stdClass[] keyed by column letter */
  protected array $cells = [];

  protected int $rowNumber = 0;

  abstract public function toSpreadsheet(Worksheet $worksheet): void;

  public function setData(array $data): static
  {
    foreach ($data as $value) {
      $this->addCell($value);
    }
    return $this;
  }

  public function setRowNumber(int $rowNumber): static
  {
    $this->rowNumber = $rowNumber;
    return $this;
  }

  public function addCell(mixed $value): static
  {
    $x = $this->getNextColumn();
    $this->cells[$x] = $this->checkCellValue($value);
    return $this;
  }

  protected function checkCellValue(mixed $value): \stdClass
  {
    if (!is_object($value)) {
      $value = (object)[
        'value' => $value,
      ];
    }

    if (!isset($value->type)) {
      $value->type = 'string';
    }

    if ($value->value && 'string' === $value->type) {
      $value->value = trim($value->value);
    }

    return $value;
  }

  protected function getNextColumn(): string
  {
    return Coordinate::stringFromColumnIndex(count($this->cells) + 1);
  }
}
