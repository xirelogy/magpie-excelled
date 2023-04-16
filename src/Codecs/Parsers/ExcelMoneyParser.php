<?php

namespace MagpieLib\Excelled\Codecs\Parsers;

use Magpie\Codecs\Parsers\CreatableParser;
use Magpie\Codecs\Parsers\FloatParser;

/**
 * Parse for money value from floating point in Excel
 */
class ExcelMoneyParser extends CreatableParser
{
    /**
     * @inheritDoc
     */
    protected function onParse(mixed $value, ?string $hintName) : int
    {
        $value = FloatParser::create()->withPrecision(2)->parse($value, $hintName);
        return floor($value * 100);
    }
}