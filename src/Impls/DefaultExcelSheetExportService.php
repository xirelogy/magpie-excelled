<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\General\Concepts\Releasable;
use MagpieLib\Excelled\Concepts\Services\ExcelCellExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelColumnExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelRowExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetExportServiceable;
use MagpieLib\Excelled\Strategies\ExcelNames;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet as PhpOfficeWorksheet;

/**
 * Default export service (sheet)
 * @internal
 */
class DefaultExcelSheetExportService implements ExcelSheetExportServiceable
{
    /**
     * @var ExcelExportServiceable Parent service
     */
    protected ExcelExportServiceable $parentService;
    /**
     * @var PhpOfficeWorksheet Associated worksheet
     */
    protected PhpOfficeWorksheet $worksheet;


    /**
     * Constructor
     * @param ExcelExportServiceable $parentService
     * @param PhpOfficeWorksheet $worksheet
     */
    public function __construct(ExcelExportServiceable $parentService, PhpOfficeWorksheet $worksheet)
    {
        $this->parentService = $parentService;
        $this->worksheet = $worksheet;
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
    public function accessRow(int $row) : ExcelRowExportServiceable
    {
        return new DefaultExcelRowExportService($this, $this->worksheet, $row);
    }


    /**
     * @inheritDoc
     */
    public function accessColumn(int $col) : ExcelColumnExportServiceable
    {
        return new DefaultExcelColumnExportService($this, $this->worksheet, $col);
    }


    /**
     * @inheritDoc
     */
    public function accessCell(int $row, int $col, ?int $row2 = null, ?int $col2 = null) : ExcelCellExportServiceable
    {
        return new DefaultExcelCellExportService($this, $this->worksheet, $row, $col, $row2, $col2);
    }


    /**
     * @inheritDoc
     */
    public function freezePane(int $row, int $col) : void
    {
        $cellName = ExcelNames::cellNameOf($row, $col);
        OfficeExcepts::protect(fn () => $this->worksheet->freezePane($cellName));
    }


    /**
     * @inheritDoc
     */
    public function finalize() : void
    {
        // Reset the cursor
        $this->worksheet->setSelectedCell(ExcelNames::cellNameOf(0, 0));
    }
}