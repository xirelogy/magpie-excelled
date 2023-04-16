<?php

namespace MagpieLib\Excelled\Impls;

use MagpieLib\Excelled\Concepts\Services\ExcelColumnExportServiceable;
use MagpieLib\Excelled\Objects\Shims\ExcelColumnAutoSize;
use MagpieLib\Excelled\Objects\Shims\ExcelColumnDefaultSize;
use MagpieLib\Excelled\Strategies\ExcelNames;
use PhpOffice\PhpSpreadsheet\Style\Style as PhpOfficeStyle;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet as PhpOfficeWorksheet;

/**
 * Default export service (column)
 * @internal
 */
class DefaultExcelColumnExportService extends DefaultExcelGeneralExportService implements ExcelColumnExportServiceable
{
    /**
     * @var PhpOfficeWorksheet Associated worksheet
     */
    protected readonly PhpOfficeWorksheet $worksheet;
    /**
     * @var int Associated column index
     */
    protected readonly int $columnIndex;
    /**
     * @var string Associated column name
     */
    protected readonly string $columnName;


    /**
     * Constructor
     * @param PhpOfficeWorksheet $worksheet
     * @param int $col
     */
    public function __construct(PhpOfficeWorksheet $worksheet, int $col)
    {
        $this->worksheet = $worksheet;
        $this->columnIndex = $col;
        $this->columnName = ExcelNames::columnNameOf($col);
    }


    /**
     * @inheritDoc
     */
    public function setWidth(float|ExcelColumnAutoSize|ExcelColumnDefaultSize $spec) : void
    {
        $columnDimension = $this->worksheet->getColumnDimension($this->columnName);

        if ($spec instanceof ExcelColumnDefaultSize) {
            if ($columnDimension->getAutoSize()) $columnDimension->setAutoSize(false);
            $columnDimension->setWidth(-1);
        } if ($spec instanceof ExcelColumnAutoSize) {
            $columnDimension->setAutoSize(true);
        } else {
            if ($columnDimension->getAutoSize()) $columnDimension->setAutoSize(false);
            $columnDimension->setWidth($spec + 0.71);
        }
    }


    /**
     * @inheritDoc
     */
    protected function getStyle() : PhpOfficeStyle
    {
        return $this->worksheet->getStyle(ExcelNames::columnRangeNameOf($this->columnIndex));
    }
}