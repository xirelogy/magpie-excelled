<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Exceptions\PersistenceException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\StreamException;
use Magpie\General\Concepts\TargetWritable;
use MagpieLib\Excelled\Concepts\ExcelFormatterAdaptable;
use MagpieLib\Excelled\Concepts\Services\ExcelExportServiceable;

/**
 * Common exporter instance
 */
abstract class CommonExporter
{
    /**
     * @var ExcelExportDefinition Definition to export (Excel/CSV)
     */
    protected readonly ExcelExportDefinition $def;
    /**
     * @var ExcelFormatterAdaptable Adapter for formatter
     */
    protected ExcelFormatterAdaptable $formatAdapter;


    /**
     * Constructor
     * @param ExcelExportDefinition $def
     */
    protected function __construct(ExcelExportDefinition $def)
    {
        $this->def = $def;
        $this->formatAdapter = new DefaultExcelFormatterAdapter();
    }


    /**
     * Export to given target
     * @param TargetWritable $target
     * @param string|null $mimeType
     * @return void
     * @throws SafetyCommonException
     * @throws PersistenceException
     * @throws StreamException
     */
    public function to(TargetWritable $target, ?string &$mimeType = null) : void
    {
        $service = $this->createService($target);
        $this->def->_run($service, $mimeType);
    }


    /**
     * Create export service
     * @param TargetWritable $target
     * @return ExcelExportServiceable
     * @throws SafetyCommonException
     * @throws PersistenceException
     * @throws StreamException
     */
    protected abstract function createService(TargetWritable $target) : ExcelExportServiceable;
}