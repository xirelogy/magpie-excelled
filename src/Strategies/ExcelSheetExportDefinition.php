<?php

namespace MagpieLib\Excelled\Strategies;

/**
 * Definition of an Excel export to a sheet
 */
abstract class ExcelSheetExportDefinition extends ExcelExportDefinition
{
    /**
     * @var string|null Specific sheet name to be used
     */
    protected ?string $sheetName;


    /**
     * Constructor
     * @param string|null $sheetName
     */
    protected function __construct(?string $sheetName)
    {
        $this->sheetName = $sheetName;
    }


    /**
     * Specific sheet name
     * @param string $sheetName
     * @return $this
     */
    public function withSheetName(string $sheetName) : static
    {
        $this->sheetName = $sheetName;
        return $this;
    }
}