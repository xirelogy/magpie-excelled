<?php

namespace MagpieLib\Excelled\Codecs\Formats;

use Magpie\General\Traits\StaticCreatable;
use MagpieLib\Excelled\Constants\ExcelCellFormat;

/**
 * Money format (for Excel)
 */
class ExcelMoneyFormatter implements ExcelFormatter
{
    use StaticCreatable;


    /**
     * @inheritDoc
     */
    public function format(mixed $value) : mixed
    {
        if ($value === null) return null;

        if (is_numeric($value)) {
            return $value / 100;
        }

        return $value;
    }


    /**
     * @inheritDoc
     */
    public function getExcelFormatString() : ?string
    {
        return ExcelCellFormat::NUMBER_COMMA_00;
    }
}