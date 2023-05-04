<?php

namespace MagpieLib\Excelled\Concepts;

use Magpie\Codecs\Formats\Formatter;
use Magpie\Exceptions\SafetyCommonException;
use MagpieLib\Excelled\Codecs\Formats\ExcelFormatter;

/**
 * May adapt a formatter according to Excel's need
 */
interface ExcelFormatterAdaptable
{
    /**
     * Adapt formatter according to Excel's need
     * @param Formatter $format
     * @return ExcelFormatter
     * @throws SafetyCommonException
     */
    public function adapt(Formatter $format) : ExcelFormatter;
}