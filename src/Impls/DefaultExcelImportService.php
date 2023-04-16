<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\UnsupportedValueException;
use MagpieLib\Excelled\Concepts\Services\ExcelImportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetImportServiceable;
use MagpieLib\Excelled\Strategies\ExcelImportOptions;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PhpOfficeSpreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet as PhpOfficeWorksheet;

/**
 * Default import service
 * @internal
 */
class DefaultExcelImportService implements ExcelImportServiceable
{
    /**
     * @var PhpOfficeSpreadsheet Associated workbook
     */
    protected readonly PhpOfficeSpreadsheet $workbook;


    /**
     * Constructor
     * @param PhpOfficeSpreadsheet $workbook
     */
    public function __construct(PhpOfficeSpreadsheet $workbook)
    {
        $this->workbook = $workbook;
    }


    /**
     * @inheritDoc
     */
    public function getSheet(ExcelImportOptions $options) : ExcelSheetImportServiceable
    {
        $worksheet = $this->getWorksheet($options->sheetName);
        return new DefaultExcelSheetImportService($worksheet, $options);
    }


    /**
     * Get corresponding worksheet
     * @param string|null $sheetName
     * @return PhpOfficeWorksheet
     * @throws SafetyCommonException
     */
    private function getWorksheet(?string $sheetName) : PhpOfficeWorksheet
    {
        if ($sheetName === null) return $this->workbook->getActiveSheet();

        return $this->workbook->getSheetByName($sheetName) ?? throw new UnsupportedValueException($sheetName);
    }
}