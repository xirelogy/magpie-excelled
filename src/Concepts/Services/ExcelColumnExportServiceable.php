<?php

namespace MagpieLib\Excelled\Concepts\Services;

use Magpie\Exceptions\SafetyCommonException;
use MagpieLib\Excelled\Objects\Shims\ExcelColumnAutoSize;
use MagpieLib\Excelled\Objects\Shims\ExcelColumnDefaultSize;

/**
 * Service interface to export to Excel column
 */
interface ExcelColumnExportServiceable extends ExcelGeneralExportServiceable
{
    /**
     * Set column width
     * @param float|ExcelColumnAutoSize|ExcelColumnDefaultSize $spec
     * @return void
     * @throws SafetyCommonException
     */
    public function setWidth(float|ExcelColumnAutoSize|ExcelColumnDefaultSize $spec) : void;
}