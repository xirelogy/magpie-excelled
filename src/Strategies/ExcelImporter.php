<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\UnsupportedValueException;
use Magpie\General\Concepts\PathTargetReadable;
use Magpie\General\Concepts\TargetReadable;
use MagpieLib\Excelled\Concepts\Services\ExcelImportServiceable;
use MagpieLib\Excelled\Impls\DefaultExcelImportService;
use MagpieLib\Excelled\Impls\OfficeExcepts;
use PhpOffice\PhpSpreadsheet\IOFactory as PhpOfficeIOFactory;
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
        if (!$target instanceof PathTargetReadable) throw new UnsupportedValueException($target);

        $path = $target->getPath();

        OfficeExcepts::protect(function () use ($path) {
            $type = PhpOfficeIOFactory::identify($path);
            $reader = PhpOfficeIOFactory::createReader($type);

            $this->workbook = $reader->load($path);
        });
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