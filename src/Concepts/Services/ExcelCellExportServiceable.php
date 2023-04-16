<?php

namespace MagpieLib\Excelled\Concepts\Services;

use Magpie\Exceptions\SafetyCommonException;

/**
 * Service interface to export to Excel cell
 */
interface ExcelCellExportServiceable extends ExcelGeneralExportServiceable
{
    /**
     * Set cell value
     * @param mixed $value
     * @return void
     * @throws SafetyCommonException
     */
    public function setValue(mixed $value) : void;
}