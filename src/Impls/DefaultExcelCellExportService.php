<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Exceptions\UnsupportedException;
use Magpie\General\Concepts\Releasable;
use MagpieLib\Excelled\Concepts\Services\ExcelCellExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetExportServiceable;
use MagpieLib\Excelled\Strategies\ExcelNames;
use PhpOffice\PhpSpreadsheet\Style\Style as PhpOfficeStyle;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet as PhpOfficeWorksheet;

/**
 * Default export service (cell)
 * @internal
 */
class DefaultExcelCellExportService extends DefaultExcelGeneralExportService implements ExcelCellExportServiceable
{
    /**
     * @var ExcelSheetExportServiceable Parent service
     */
    protected ExcelSheetExportServiceable $parentService;
    /**
     * @var PhpOfficeWorksheet Associated worksheet
     */
    protected readonly PhpOfficeWorksheet $worksheet;
    /**
     * @var string Associated cell name
     */
    protected readonly string $cellName;
    /**
     * @var bool If specification is a range
     */
    protected readonly bool $isRange;


    /**
     * Constructor
     * @param ExcelSheetExportServiceable $parentService
     * @param PhpOfficeWorksheet $worksheet
     * @param int $row
     * @param int $col
     * @param int|null $row2
     * @param int|null $col2
     */
    public function __construct(ExcelSheetExportServiceable $parentService, PhpOfficeWorksheet $worksheet, int $row, int $col, ?int $row2, ?int $col2)
    {
        $this->parentService = $parentService;
        $this->worksheet = $worksheet;
        $this->cellName = static::formatCellName($row, $col, $row2, $col2);
        $this->isRange = $row2 !== null || $col2 !== null;
    }


    /**
     * @inheritDoc
     */
    public function addReleasable(Releasable $resource) : void
    {
        $this->parentService->addReleasable($resource);
    }


    /**
     * @inheritDoc
     */
    public function setValue(mixed $value) : void
    {
        if ($this->isRange) throw new UnsupportedException();

        $this->worksheet->setCellValue($this->cellName, $value);
    }


    /**
     * @inheritDoc
     */
    protected function getStyle() : PhpOfficeStyle
    {
        return $this->worksheet->getStyle($this->cellName);
    }


    /**
     * Format cell name
     * @param int $row
     * @param int $col
     * @param int|null $row2
     * @param int|null $col2
     * @return string
     */
    protected static function formatCellName(int $row, int $col, ?int $row2, ?int $col2) : string
    {
        if ($row2 !== null && $col2 !== null) {
            // Assume cell range
            return ExcelNames::rangeNameOf($row, $col, $row2, $col2);
        }

        if ($col2 !== null) {
            // Assume column range
            return ExcelNames::columnRangeNameOf($col, $col2);
        }

        if ($row2 !== null) {
            // Assume row range
            return ExcelNames::rowRangeNameOf($row, $row2);
        }

        // Assume single cell
        return ExcelNames::cellNameOf($row, $col);
    }
}