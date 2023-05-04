<?php

namespace MagpieLib\Excelled\Concepts\Services;

use Magpie\Codecs\Formats\Formatter;
use Magpie\Exceptions\SafetyCommonException;

/**
 * Service interface to export to Excel cell
 */
interface ExcelCellExportServiceable extends ExcelGeneralExportServiceable, ExcelResourceManageable
{
    /**
     * Set cell value
     * @param mixed $value
     * @param Formatter|null $formatter
     * @return void
     * @throws SafetyCommonException
     */
    public function setValue(mixed $value, ?Formatter $formatter = null) : void;
}