<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\General\Concepts\TargetWritable;
use MagpieLib\Excelled\Concepts\Services\ExcelExportServiceable;
use MagpieLib\Excelled\Impls\DefaultCsvExportService;

/**
 * CSV exporter instance
 */
class CsvExporter extends CommonExporter
{
    /**
     * @inheritDoc
     */
    protected function createService(TargetWritable $target) : ExcelExportServiceable
    {
        return new DefaultCsvExportService($this->formatAdapter, $target);
    }


    /**
     * Create an instance
     * @param ExcelSheetExportDefinition $def
     * @return static
     */
    public static function create(ExcelSheetExportDefinition $def) : static
    {
        return new static($def);
    }
}