<?php

namespace MagpieLib\Excelled\Strategies;

use MagpieLib\Excelled\Concepts\Services\ExcelExportServiceable;

class ExcelTemplatedExportDefinition extends ExcelExportDefinition
{
    /**
     * @var ExcelExportTemplatedSchema Associated schema
     */
    protected readonly ExcelExportTemplatedSchema $schema;


    /**
     * Constructor
     * @param ExcelExportTemplatedSchema $schema
     */
    protected function __construct(ExcelExportTemplatedSchema $schema)
    {
        $this->schema = $schema;
    }


    /**
     * @inheritDoc
     */
    protected function onRun(ExcelExportServiceable $service) : void
    {
        $this->schema->exportUsing($service);
    }


    /**
     * Create a new instance
     * @param ExcelExportTemplatedSchema $schema
     * @return static
     */
    public static function create(ExcelExportTemplatedSchema $schema) : static
    {
        return new static($schema);
    }
}