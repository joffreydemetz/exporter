<?php

/**
 * (c) Joffrey Demetz <joffrey.demetz@gmail.com>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace JDZ\Exporter\Tests;

use JDZ\Exporter\Exporter;
use JDZ\Exporter\Header;
use JDZ\Exporter\Line;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PHPUnit\Framework\TestCase;

/**
 * @covers \JDZ\Exporter\Exporter
 * @covers \JDZ\Exporter\Row
 * @covers \JDZ\Exporter\Header
 * @covers \JDZ\Exporter\Line
 */
class ExporterTest extends TestCase
{
    private function rows(): array
    {
        return [
            ['id' => 1, 'title' => 'One', 'price' => 10],
            ['id' => 2, 'title' => 'Two', 'price' => 20],
        ];
    }

    public function testSetRowsBuildsHeaderAndLines(): void
    {
        $exporter = (new Exporter())->setRows($this->rows(), 'Items')->render();

        $sheet = $exporter->getSpreadsheet()->getSheet(0);

        $this->assertEquals('Items', $sheet->getTitle());
        // header from first row's keys
        $this->assertEquals('id', $sheet->getCell('A1')->getValue());
        $this->assertEquals('title', $sheet->getCell('B1')->getValue());
        $this->assertTrue($sheet->getStyle('A1')->getFont()->getBold());
        // data rows
        $this->assertEquals('One', $sheet->getCell('B2')->getValue());
        $this->assertEquals('Two', $sheet->getCell('B3')->getValue());
    }

    public function testSetRowsEmptyThrowsUnlessAllowed(): void
    {
        $exporter = new Exporter();

        $this->assertSame($exporter, $exporter->setRows([], 'Empty', true));

        $this->expectExceptionMessage('Nothing to export to Empty');
        $exporter->setRows([], 'Empty');
    }

    public function testWriterSelection(): void
    {
        $this->assertEquals('csv', (new Exporter())->withCsv()->getWriter());
        $this->assertEquals('ods', (new Exporter())->withOds()->getWriter());
        $this->assertEquals('xlsx', (new Exporter())->setWriter('xlsx')->getWriter());
    }

    public function testSaveWritesFilePerWriter(): void
    {
        foreach (['xlsx', 'csv', 'ods'] as $writer) {
            $path = tempnam(sys_get_temp_dir(), 'jdzexp') . '.' . $writer;

            (new Exporter())
                ->setWriter($writer)
                ->setRows($this->rows(), 'Items')
                ->render()
                ->save($path);

            $this->assertFileExists($path);
            $this->assertGreaterThan(0, filesize($path), $writer . ' output is empty');
            unlink($path);
        }
    }

    public function testCreatorIsConfigurable(): void
    {
        $exporter = new Exporter('Custom Creator');
        $this->assertEquals('Custom Creator', $exporter->getSpreadsheet()->getProperties()->getCreator());

        $exporter->setCreator('Other');
        $this->assertEquals('Other', $exporter->getSpreadsheet()->getProperties()->getCreator());
        $this->assertEquals('Other', $exporter->getSpreadsheet()->getProperties()->getLastModifiedBy());
    }

    public function testLineCellTyping(): void
    {
        $exporter = new Exporter();
        $sheet = $exporter->getSpreadsheet()->getSheet(0);

        $line = (new Line())->setData([
            (object)['type' => 'id', 'value' => 12],
            (object)['type' => 'bool', 'value' => true],
            (object)['type' => 'phone', 'value' => '0660000000'],
            'plain string ',
        ]);
        $line->setRowNumber(1);
        $line->toSpreadsheet($sheet);

        $this->assertEquals(12, $sheet->getCell('A1')->getValue());
        $this->assertEquals('X', $sheet->getCell('B1')->getValue());
        // phone renders as formula-wrapped text to preserve leading zeros
        $this->assertEquals('="0660000000"', $sheet->getCell('C1')->getValue());
        // plain string values are trimmed
        $this->assertEquals('plain string', $sheet->getCell('D1')->getValue());
    }

    public function testQuotesAreEscapedInFormulaCells(): void
    {
        $exporter = new Exporter();
        $sheet = $exporter->getSpreadsheet()->getSheet(0);

        $line = (new Line())->setData([
            (object)['type' => 'text', 'value' => 'Ecole dite "Nationale"'],
        ]);
        $line->setRowNumber(1);
        $line->toSpreadsheet($sheet);

        $this->assertEquals('="Ecole dite ""Nationale"""', $sheet->getCell('A1')->getValue());
    }

    public function testMoreThan26Columns(): void
    {
        $exporter = new Exporter();
        $sheet = $exporter->getSpreadsheet()->getSheet(0);

        $header = (new Header())->setData(array_map(
            static fn(int $i): string => 'col' . $i,
            range(1, 30)
        ));
        $header->setRowNumber(1);
        $header->toSpreadsheet($sheet);

        $this->assertEquals('col26', $sheet->getCell('Z1')->getValue());
        $this->assertEquals('col27', $sheet->getCell('AA1')->getValue());
        $this->assertEquals('col30', $sheet->getCell('AD1')->getValue());
    }

    public function testContentTypeFollowsWriter(): void
    {
        $this->assertEquals('text/csv', (new Exporter())->withCsv()->getContentType());
        $this->assertEquals('application/vnd.oasis.opendocument.spreadsheet', (new Exporter())->withOds()->getContentType());
        $this->assertEquals('application/pdf', (new Exporter())->setWriter('pdf')->getContentType());
        $this->assertEquals(
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            (new Exporter())->getContentType()
        );
    }
}
