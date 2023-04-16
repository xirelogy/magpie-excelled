<?php

namespace MagpieLib\Excelled\Impls;

use MagpieLib\Excelled\Concepts\Services\ExcelRowExportServiceable;
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
     * @var PhpOfficeWorksheet Associated worksheet
     */
    protected readonly PhpOfficeWorksheet $worksheet;
    /**
     * @var int Associated row index
     */
    protected readonly int $rowIndex;


    /**
     * Constructor
     * @param PhpOfficeWorksheet $worksheet
     * @param int $row
     */
    public function __construct(PhpOfficeWorksheet $worksheet, int $row)
    {
        $this->worksheet = $worksheet;
        $this->rowIndex = $row;
    }


    /**
     * @inheritDoc
     */
    protected function getStyle() : PhpOfficeStyle
    {
        return $this->worksheet->getStyle(ExcelNames::rowRangeNameOf($this->rowIndex));
    }
}