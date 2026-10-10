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
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\DataProvider;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PHPUnit\Framework\TestCase;

#[CoversClass(Exporter::class)]
#[CoversClass(\JDZ\Exporter\Row::class)]
#[CoversClass(Header::class)]
#[CoversClass(Line::class)]
class ExporterTest extends TestCase
{
    /** @var string[] */
    private array $tmpFiles = [];

    protected function tearDown(): void
    {
        foreach ($this->tmpFiles as $path) {
            if (is_file($path)) {
                unlink($path);
            }
        }
        $this->tmpFiles = [];
    }

    private function tmpPath(string $extension): string
    {
        $path = sys_get_temp_dir() . '/jdz-exporter-' . uniqid('', true) . '.' . $extension;
        $this->tmpFiles[] = $path;

        return $path;
    }

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

    public static function writerToggles(): array
    {
        return [
            'default' => [fn(Exporter $e) => $e, 'xlsx'],
            'withCsv()' => [fn(Exporter $e) => $e->withCsv(), 'csv'],
            'withOds()' => [fn(Exporter $e) => $e->withOds(), 'ods'],
            'setWriter()' => [fn(Exporter $e) => $e->withCsv()->setWriter('xlsx'), 'xlsx'],
            'withCsv(false) by default' => [fn(Exporter $e) => $e->withCsv(false), 'xlsx'],
            'withCsv(false) after withCsv()' => [fn(Exporter $e) => $e->withCsv()->withCsv(false), 'xlsx'],
            'withOds(false) keeps csv' => [fn(Exporter $e) => $e->withCsv()->withOds(false), 'csv'],
            'withCsv(false) keeps ods' => [fn(Exporter $e) => $e->withOds()->withCsv(false), 'ods'],
        ];
    }

    #[DataProvider('writerToggles')]
    public function testWriterToggles(\Closure $toggle, string $writer): void
    {
        $this->assertSame($writer, $toggle(new Exporter())->getWriter());
    }

    public function testSaveWritesFilePerWriter(): void
    {
        foreach (['xlsx', 'csv', 'ods'] as $writer) {
            $path = $this->tmpPath($writer);

            (new Exporter())
                ->setWriter($writer)
                ->setRows($this->rows(), 'Items')
                ->render()
                ->save($path);

            $this->assertFileExists($path);
            $this->assertGreaterThan(0, filesize($path), $writer . ' output is empty');
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


    public function testEverySheetIsRenderedInOrder(): void
    {
        $book = (new Exporter())
            ->setRows([['a' => 'first']], 'First')
            ->setRows([['b' => 'x'], ['b' => 'y']], 'Second')
            ->render()
            ->getSpreadsheet();

        $this->assertSame(['First', 'Second'], $book->getSheetNames());
        $this->assertSame([['a'], ['first']], $book->getSheet(0)->toArray());
        $this->assertSame([['b'], ['x'], ['y']], $book->getSheet(1)->toArray());
    }

    public function testCsvHoldsTheFirstSheetOnlySemicolonSeparatedAndQuoted(): void
    {
        $path = $this->tmpPath('csv');

        (new Exporter())
            ->setRows([['id' => 1, 'title' => 'One "quoted"'], ['id' => 2, 'title' => 'Two;semi']], 'Items')
            ->setRows([['x' => 'other']], 'Other')
            ->withCsv()
            ->render()
            ->save($path);

        $this->assertSame(
            implode(PHP_EOL, ['"id";"title"', '"1";"One ""quoted"""', '"2";"Two;semi"']) . PHP_EOL,
            file_get_contents($path)
        );
    }

    public static function cellStyles(): array
    {
        return [
            'id: right, numeric' => ['id', 12, Alignment::HORIZONTAL_RIGHT, DataType::TYPE_NUMERIC, NumberFormat::FORMAT_GENERAL],
            'number: right, numeric' => ['number', 3.5, Alignment::HORIZONTAL_RIGHT, DataType::TYPE_NUMERIC, NumberFormat::FORMAT_GENERAL],
            'euro: numeric, whole euros' => ['euro', 12, Alignment::HORIZONTAL_LEFT, DataType::TYPE_NUMERIC, NumberFormat::FORMAT_CURRENCY_EUR_INTEGER],
            'date: centred text' => ['date', '2024-03-05', Alignment::HORIZONTAL_CENTER, DataType::TYPE_STRING, NumberFormat::FORMAT_GENERAL],
            'bool: centred' => ['bool', false, Alignment::HORIZONTAL_CENTER, DataType::TYPE_STRING, NumberFormat::FORMAT_GENERAL],
            'string: left text' => ['string', 'abc', Alignment::HORIZONTAL_LEFT, DataType::TYPE_STRING, NumberFormat::FORMAT_GENERAL],
        ];
    }

    #[DataProvider('cellStyles')]
    public function testCellAlignmentTypeAndFormat(string $type, mixed $value, string $horizontal, string $dataType, string $format): void
    {
        $sheet = (new Exporter())->getSpreadsheet()->getSheet(0);

        $line = (new Line())->setData([(object)['type' => $type, 'value' => $value]]);
        $line->setRowNumber(1);
        $line->toSpreadsheet($sheet);

        $this->assertSame($horizontal, $sheet->getStyle('A1')->getAlignment()->getHorizontal());
        $this->assertSame($dataType, $sheet->getCell('A1')->getDataType());
        $this->assertSame($format, $sheet->getStyle('A1')->getNumberFormat()->getFormatCode());
    }
}
