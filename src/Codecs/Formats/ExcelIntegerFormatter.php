<?php

namespace MagpieLib\Excelled\Codecs\Formats;

use Magpie\General\Traits\StaticCreatable;
use MagpieLib\Excelled\Constants\ExcelCellFormat;

/**
 * Integer formatter (for Excel)
 */
class ExcelIntegerFormatter implements ExcelFormatter
{
    use StaticCreatable;


    /**
     * @inheritDoc
     */
    public function format(mixed $value) : mixed
    {
        if ($value === null) return null;

        if (is_numeric($value)) {
            return intval(round($value));
        }

        return $value;
    }


    /**
     * @inheritDoc
     */
    public function getExcelFormatString() : ?string
    {
        return ExcelCellFormat::NUMBER;
    }
}