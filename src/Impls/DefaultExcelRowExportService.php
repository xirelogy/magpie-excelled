<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\General\Concepts\Releasable;
use MagpieLib\Excelled\Concepts\Services\ExcelRowExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetExportServiceable;
use MagpieLib\Excelled\Strategies\ExcelNames;
use PhpOffice\PhpSpreadsheet\Style\Style as PhpOfficeStyle;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet as PhpOfficeWorksheet;

/**
 * Default export service (row)
 * @internal
 */
class DefaultExcelRowExportService extends DefaultExcelGeneralExportService implements ExcelRowExportServiceable
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
     * @var int Associated row index
     */
    protected readonly int $rowIndex;


    /**
     * Constructor
     * @param ExcelSheetExportServiceable $parentService
     * @param PhpOfficeWorksheet $worksheet
     * @param int $row
     */
    public function __construct(ExcelSheetExportServiceable $parentService, PhpOfficeWorksheet $worksheet, int $row)
    {
        $this->parentService = $parentService;
        $this->worksheet = $worksheet;
        $this->rowIndex = $row;
    }


    public function addReleasable(Releasable $resource) : void
    {
        $this->parentService->addReleasable($resource);
    }


    /**
     * @inheritDoc
     */
    protected function getStyle() : PhpOfficeStyle
    {
        return $this->worksheet->getStyle(ExcelNames::rowRangeNameOf($this->rowIndex));
    }
}