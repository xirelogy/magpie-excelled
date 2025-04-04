<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\General\Concepts\TargetWritable;
use MagpieLib\Excelled\Concepts\Services\ExcelExportServiceable;
use MagpieLib\Excelled\Impls\DefaultExcelExportService;

/**
 * Excel exporter instance
 */
class ExcelExporter extends CommonExporter
{
    /**
     * @inheritDoc
     */
    protected function createService(TargetWritable $target) : ExcelExportServiceable
    {
        return new DefaultExcelExportService($this->formatAdapter, $target);
    }


    /**
     * Create an instance
     * @param ExcelExportDefinition $def
     * @return static
     */
    public static function create(ExcelExportDefinition $def) : static
    {
        return new static($def);
    }
}