<?php

namespace MagpieLib\Excelled\Concepts\Services;

use Magpie\Exceptions\SafetyCommonException;
use MagpieLib\Excelled\Strategies\ExcelImportOptions;

/**
 * Service interface to import from Excel
 */
interface ExcelImportServiceable
{
    /**
     * Get specific sheet
     * @param ExcelImportOptions $options
     * @return ExcelSheetImportServiceable
     * @throws SafetyCommonException
     */
    public function getSheet(ExcelImportOptions $options) : ExcelSheetImportServiceable;
}