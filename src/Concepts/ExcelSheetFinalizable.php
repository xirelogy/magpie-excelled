<?php

namespace MagpieLib\Excelled\Concepts;

use Magpie\Exceptions\SafetyCommonException;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetExportServiceable;

interface ExcelSheetFinalizable
{
    /**
     * Finalize the sheet
     * @param ExcelSheetExportServiceable $sheetService
     * @return void
     * @throws SafetyCommonException
     */
    public function finalizeSheet(ExcelSheetExportServiceable $sheetService) : void;
}