<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\General\Sugars\Excepts;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetImportServiceable;
use MagpieLib\Excelled\Objects\ExcelImages;
use MagpieLib\Excelled\Strategies\ExcelImportOptions;
use MagpieLib\Excelled\Strategies\ExcelNames;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing as PhpOfficeDrawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Row as PhpOfficeRow;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet as PhpOfficeWorksheet;

/**
 * Default import service (sheet)
 * @internal
 */
class DefaultExcelSheetImportService implements ExcelSheetImportServiceable
{
    /**
     * @var PhpOfficeWorksheet Associated worksheet
     */
    protected readonly PhpOfficeWorksheet $worksheet;
    /**
     * @var array<string, array<PhpOfficeDrawing>> Map of associated drawings
     */
    protected array $drawingLists = [];


    /**
     * Constructor
     * @param PhpOfficeWorksheet $worksheet
     * @param ExcelImportOptions $options
     */
    public function __construct(PhpOfficeWorksheet $worksheet, ExcelImportOptions $options)
    {
        $this->worksheet = $worksheet;

        if ($options->isImportImages) {
            /** @var PhpOfficeDrawing $drawing */
            foreach ($this->worksheet->getDrawingCollection()->getIterator() as $drawing) {
                $coordinate = $drawing->getCoordinates();
                $drawingList = $this->drawingLists[$coordinate] ?? [];
                $drawingList[] = $drawing;
                $this->drawingLists[$coordinate] = $drawingList;
            }
        }
    }


    /**
     * @inheritDoc
     */
    public function getRows(int $startRowIndex, int $startColIndex, ?int $endColIndex = null, int $lockRowCount = 1) : iterable
    {
        $startRow = $startRowIndex + 1;
        $startCol = $startColIndex;
        $endCol = $endColIndex;

        $currentLockRowCount = $lockRowCount;
        foreach ($this->worksheet->getRowIterator($startRow) as $row) {
            if ($currentLockRowCount <= 0 && $row->isEmpty()) break;

            $rowArray = $this->flattenRow($row, $startCol, $endCol);
            $rowSize = count($rowArray);
            if ($endCol === null || $rowSize > $endCol) $endCol = $rowSize;

            yield $rowArray;
            if ($currentLockRowCount > 0) {
                --$currentLockRowCount;
            }
        }
    }


    /**
     * Flatten a row into array
     * @param PhpOfficeRow $row
     * @param int $startCol
     * @param int|null $endCol
     * @return array
     */
    private function flattenRow(PhpOfficeRow $row, int $startCol, ?int $endCol = null) : array
    {
        $endColName = $endCol !== null ? ExcelNames::columnNameOf($endCol) : null;

        $ret = [];
        foreach ($row->getCellIterator(ExcelNames::columnNameOf($startCol), $endColName) as $cell) {
            $cellCoordinate = Excepts::noThrow(fn () => $cell->getCoordinate(), '');
            $cellValue = $cell->getValue();

            if ($cellValue === null) {
                if (array_key_exists($cellCoordinate, $this->drawingLists)) {
                    $cellValue = static::adaptDrawingList($this->drawingLists[$cellCoordinate]);
                }
            }

            if ($endColName === null && $cellValue === null) break;
            $ret[] = $cellValue;
        }

        return $ret;
    }


    /**
     * Adapt list of drawings
     * @param iterable<PhpOfficeDrawing> $drawings
     * @return ExcelImages|null
     */
    protected static function adaptDrawingList(iterable $drawings) : ?ExcelImages
    {
        $ret = [];

        foreach ($drawings as $drawing) {
            $ret[] = new ExcelDrawingImage($drawing);
        }

        if (count($ret) <= 0) return null;

        return Excepts::noThrow(fn () => new ExcelImages($ret));
    }
}