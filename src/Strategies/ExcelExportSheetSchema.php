<?php

namespace MagpieLib\Excelled\Strategies;

use MagpieLib\Excelled\Concepts\ExcelSheetFinalizable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetExportServiceable;

/**
 * A schema to export to Excel sheet
 */
abstract class ExcelExportSheetSchema extends ExcelExportSchema implements ExcelSheetFinalizable
{
    /**
     * @inheritDoc
     */
    public function finalizeSheet(ExcelSheetExportServiceable $sheetService) : void
    {
        // Default NOP
    }
}