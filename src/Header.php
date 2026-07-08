<?php

/**
 * (c) Joffrey Demetz <joffrey.demetz@gmail.com>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace JDZ\Exporter;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Spreadsheet header row — bold, centered.
 *
 * @author Joffrey Demetz <joffrey.demetz@gmail.com>
 */
class Header extends Row
{
  public function toSpreadsheet(Worksheet $worksheet): void
  {
    foreach ($this->cells as $column => $cell) {
      $worksheet->setCellValueExplicit($column . $this->rowNumber, $cell->value, DataType::TYPE_STRING);

      $worksheet->getStyle($column . $this->rowNumber)->getAlignment()
        ->applyFromArray([
          'horizontal' => Alignment::HORIZONTAL_CENTER,
          'vertical' => Alignment::VERTICAL_TOP,
        ]);

      $worksheet->getStyle($column . $this->rowNumber)->getFont()->setBold(true);
    }
  }
}
