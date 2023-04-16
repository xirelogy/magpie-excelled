<?php

namespace MagpieLib\Excelled\Codecs\Formats;

use Magpie\Codecs\Formats\Formatter;

/**
 * Extended Formatter with Excel specific format string provided
 */
interface ExcelFormatter extends Formatter
{
    /**
     * The specific Excel format string
     * @return string|null
     */
    public function getExcelFormatString() : ?string;
}