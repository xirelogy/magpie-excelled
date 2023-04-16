<?php

namespace MagpieLib\Excelled\Concepts\Services;

use Magpie\Exceptions\SafetyCommonException;
use MagpieLib\Excelled\Objects\Styles\ExcelStyle;

/**
 * Service interface to export to Excel, in general
 */
interface ExcelGeneralExportServiceable
{
    /**
     * Set cell/cells format
     * @param string $formatString
     * @return void
     * @throws SafetyCommonException
     */
    public function setFormat(string $formatString) : void;


    /**
     * Apply cell/cells style(s)
     * @param ExcelStyle ...$styles
     * @return void
     * @throws SafetyCommonException
     */
    public function applyStyle(ExcelStyle ...$styles) : void;
}