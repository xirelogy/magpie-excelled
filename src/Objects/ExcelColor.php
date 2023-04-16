<?php

namespace MagpieLib\Excelled\Objects;

use Magpie\Objects\CommonObject;

/**
 * Excel color specification
 */
abstract class ExcelColor extends CommonObject
{
    /**
     * Corresponding color string
     * @return string
     */
    public abstract function getColorString() : string;


    /**
     * Translate value to corresponding hex
     * @param int $v
     * @return string
     */
    protected static function hex(int $v) : string
    {
        // Control range to be within 0 ~ 255
        if ($v < 0) $v = 0;
        if ($v > 255) $v = 255;

        $h1 = floor($v / 16);
        $h2 = $v % 16;

        $cs = '0123456789ABCDEF';
        return substr($cs, $h1, 1) . substr($cs, $h2, 1);
    }
}