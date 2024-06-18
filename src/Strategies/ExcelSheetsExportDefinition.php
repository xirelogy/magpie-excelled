<?php

namespace MagpieLib\Excelled\Strategies;

use MagpieLib\Excelled\Concepts\Services\ExcelExportServiceable;

/**
 * Definition of an Excel export to multiple sheets
 */
class ExcelSheetsExportDefinition extends ExcelExportDefinition
{
    /**
     * @var array<ExcelSheetExportDefinition> Sheets to be exported
     */
    protected array $sheets;


    /**
     * Constructor
     * @param iterable<ExcelSheetExportDefinition> $sheets
     */
    protected function __construct(iterable $sheets)
    {
        $this->sheets = iter_flatten($sheets, false);
    }


    /**
     * @inheritDoc
     */
    protected function onRun(ExcelExportServiceable $service) : void
    {
        foreach ($this->sheets as $sheet) {
            $sheet->onRun($service);
        }
    }


    /**
     * Create an instance
     * @param iterable<ExcelSheetExportDefinition> $sheets
     * @return static
     */
    public static function create(iterable $sheets) : static
    {
        return new static($sheets);
    }
}