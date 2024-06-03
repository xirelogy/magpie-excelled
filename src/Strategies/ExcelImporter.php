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
class ExcelImporter extends CommonImporter
{
    /**
     * @var PhpOfficeSpreadsheet Associated workbook
     */
    protected PhpOfficeSpreadsheet $workbook;


    /**
     * Constructor
     * @param TargetReadable $target
     * @param ExcelImporterOptions $options
     * @throws SafetyCommonException
     */
    protected function __construct(TargetReadable $target, ExcelImporterOptions $options)
    {
        parent::__construct();
        $this->workbook = ExcelIO::readWorkbookFromTarget($target, $options->isSaveMemory);
    }


    /**
     * @inheritDoc
     */
    protected function onGetService() : ExcelImportServiceable
    {
        return new DefaultExcelImportService($this->workbook);
    }


    /**
     * @inheritDoc
     */
    protected static function onCreate(TargetReadable $target, CommonImporterOptions $options) : static
    {
        $options = $options instanceof ExcelImporterOptions ? $options : ExcelImporterOptions::default();
        return new static($target, $options);
    }
}