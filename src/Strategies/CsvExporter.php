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
     * @var CsvExporterOptions Associated options
     */
    protected readonly CsvExporterOptions $options;


    /**
     * Constructor
     * @param ExcelExportDefinition $def
     * @param CsvExporterOptions $options
     */
    protected function __construct(ExcelExportDefinition $def, CsvExporterOptions $options)
    {
        parent::__construct($def);

        $this->options = $options;
    }


    /**
     * @inheritDoc
     */
    protected function createService(TargetWritable $target) : ExcelExportServiceable
    {
        return new DefaultCsvExportService($this->formatAdapter, $target, $this->options);
    }


    /**
     * Create an instance
     * @param ExcelSheetExportDefinition $def
     * @param CsvExporterOptions|null $options
     * @return static
     */
    public static function create(ExcelSheetExportDefinition $def, ?CsvExporterOptions $options = null) : static
    {
        $options = $options ?? CsvExporterOptions::default();

        return new static($def, $options);
    }
}