<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Exceptions\PersistenceException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\StreamException;
use Magpie\General\Concepts\TargetWritable;
use MagpieLib\Excelled\Concepts\ExcelColumnAdaptable;
use MagpieLib\Excelled\Impls\DefaultExcelExportService;

/**
 * Excel exporter instance
 */
class ExcelExporter
{
    /**
     * @var ExcelExportDefinition Definition to export excel
     */
    protected readonly ExcelExportDefinition $def;
    /**
     * @var ExcelColumnAdaptable Adapter for column
     */
    protected ExcelColumnAdaptable $columnAdapter;


    /**
     * Constructor
     * @param ExcelExportDefinition $def
     */
    protected function __construct(ExcelExportDefinition $def)
    {
        $this->def = $def;
        $this->columnAdapter = new DefaultExcelColumnAdapter();
    }


    /**
     * Export to given target
     * @param TargetWritable $target
     * @return void
     * @throws SafetyCommonException
     * @throws PersistenceException
     * @throws StreamException
     */
    public function to(TargetWritable $target) : void
    {
        $service = new DefaultExcelExportService($this->columnAdapter, $target);
        $this->def->_run($service);
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