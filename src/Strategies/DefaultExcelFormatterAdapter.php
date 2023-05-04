<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Codecs\Formats\Formatter;
use Magpie\Codecs\Formats\StringFormatter;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\UnsupportedValueException;
use MagpieLib\Excelled\Codecs\Formats\ExcelFormatter;
use MagpieLib\Excelled\Codecs\Formats\ExcelUsingFormatter;
use MagpieLib\Excelled\Concepts\ExcelFormatterAdaptable;
use MagpieLib\Excelled\Constants\ExcelCellFormat;

/**
 * Default implementation of ExcelFormatterAdaptable
 */
class DefaultExcelFormatterAdapter implements ExcelFormatterAdaptable
{
    /**
     * @inheritDoc
     */
    public function adapt(Formatter $format) : ExcelFormatter
    {
        if ($format instanceof ExcelFormatter) return $format;

        return $this->onAdapt($format);
    }


    /**
     * Adapt formatter according to Excel's need
     * @param Formatter $format
     * @return ExcelFormatter
     * @throws SafetyCommonException
     */
    protected function onAdapt(Formatter $format) : ExcelFormatter
    {
        if ($format instanceof StringFormatter) return ExcelUsingFormatter::extends($format, ExcelCellFormat::TEXT);

        throw new UnsupportedValueException($format);
    }
}