<?php

namespace MagpieLib\Excelled\Codecs\Parsers;

use Magpie\Codecs\Parsers\CreatableParser;
use Magpie\Exceptions\InvalidDataException;
use MagpieLib\Excelled\Objects\ExcelImage;
use MagpieLib\Excelled\Objects\ExcelImages;

/**
 * Parse for first valid image from Excel
 */
class ExcelFirstImageParser extends CreatableParser
{
    /**
     * @inheritDoc
     */
    protected function onParse(mixed $value, ?string $hintName) : ExcelImage
    {
        if (!$value instanceof ExcelImages) throw new InvalidDataException();

        return $value->getFirst();
    }
}