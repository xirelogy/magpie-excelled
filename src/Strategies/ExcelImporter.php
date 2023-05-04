<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Exceptions\SafetyCommonException;
use Magpie\General\Concepts\TargetReadable;
use MagpieLib\Excelled\Concepts\Services\ExcelImportServiceable;
use MagpieLib\Excelled\Impls\DefaultExcelImportService;
use MagpieLib\Excelled\Impls\ExcelIO;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PhpOfficeSpreadsheet;

/**
 * Excel importer instance
 */
class ExcelImporter
{
    /**
     * @var PhpOfficeSpreadsheet Associated workbook
     */
    protected PhpOfficeSpreadsheet $workbook;


    /**
     * Constructor
     * @param TargetReadable $target
     * @throws SafetyCommonException
     */
    protected function __construct(TargetReadable $target)
    {
        $this->workbook = ExcelIO::readWorkbookFromTarget($target);
    }


    /**
     * Access to the service interface
     * @return ExcelImportServiceable
     * @internal
     */
    public function _getService() : ExcelImportServiceable
    {
        return new DefaultExcelImportService($this->workbook);
    }


    /**
     * Create an instance
     * @param TargetReadable $target
     * @return static
     * @throws SafetyCommonException
     */
    public static function create(TargetReadable $target) : static
    {
        return new static($target);
    }
}