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
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Spreadsheet data row — cell rendering driven by the cell `type`
 * (string / bool / formula / text / phone / id / number / euro / date).
 *
 * @author Joffrey Demetz <joffrey.demetz@gmail.com>
 */
class Line extends Row
{
  public function toSpreadsheet(Worksheet $worksheet): void
  {
    foreach ($this->cells as $column => $cell) {
      $style = [];
      $style['horizontal'] = Alignment::HORIZONTAL_LEFT;
      $style['vertical'] = Alignment::VERTICAL_TOP;
      $style['wrapText'] = false;

      $dataType = DataType::TYPE_STRING;
      $text = $cell->value;

      if ('bool' === $cell->type) {
        $text = true === $cell->value ? 'X' : '';
      }

      if ('formula' === $cell->type || 'text' === $cell->type || 'phone' === $cell->type) {
        $text = '="' . str_replace('"', '""', (string)$cell->value) . '"';
        $dataType = DataType::TYPE_FORMULA;
      }

      if ('id' === $cell->type || 'number' === $cell->type) {
        $style['horizontal'] = Alignment::HORIZONTAL_RIGHT;
        $dataType = DataType::TYPE_NUMERIC;
      }

      if ('euro' === $cell->type) {
        $dataType = DataType::TYPE_NUMERIC;
      }

      if ('date' === $cell->type || 'bool' === $cell->type) {
        $style['horizontal'] = Alignment::HORIZONTAL_CENTER;
      }

      $worksheet->setCellValueExplicit($column . $this->rowNumber, $text, $dataType);

      $worksheet->getStyle($column . $this->rowNumber)->getAlignment()
        ->applyFromArray($style);

      if ('euro' === $cell->type) {
        $worksheet->getStyle($column . $this->rowNumber)->getNumberFormat()
          ->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR_INTEGER);
      }
    }
  }
}
