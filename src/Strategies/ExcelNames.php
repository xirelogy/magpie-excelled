<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\General\Traits\StaticClass;

/**
 * Excel naming scheme
 */
class ExcelNames
{
    use StaticClass;


    /**
     * The corresponding range name in alphabet + number pair
     * @param int $row1
     * @param int $col1
     * @param int $row2
     * @param int $col2
     * @return string
     */
    public static function rangeNameOf(int $row1, int $col1, int $row2, int $col2) : string
    {
        return static::cellNameOf($row1, $col1) . ':' . static::cellNameOf($row2, $col2);
    }


    /**
     * The corresponding range name in alphabet + number pair for full row
     * @param int $row1
     * @param int|null $row2
     * @return string
     */
    public static function rowRangeNameOf(int $row1, ?int $row2 = null) : string
    {
        return ($row1 + 1) . ':' . (($row2 ?? $row1) + 1);
    }


    /**
     * The corresponding range name in alphabet + number pair for full column
     * @param int $col1
     * @param int|null $col2
     * @return string
     */
    public static function columnRangeNameOf(int $col1, ?int $col2 = null) : string
    {
        return static::columnNameOf($col1) . ':' . static::columnNameOf($col2 ?? $col1);
    }


    /**
     * The corresponding cell name in alphabet + number format
     * @param int $row
     * @param int $col
     * @return string
     */
    public static function cellNameOf(int $row, int $col) : string
    {
        return static::columnNameOf($col) . ($row + 1);
    }


    /**
     * The corresponding column name in alphabet
     * @param int $col
     * @return string
     */
    public static function columnNameOf(int $col) : string
    {
        $ret = '';
        ++$col;
        do {
            --$col;
            $r = $col % 26;
            $col = intdiv($col, 26);
            /** @noinspection SpellCheckingInspection */
            $ret = substr('ABCDEFGHIJKLMNOPQRSTUVWXYZ', $r, 1) . $ret;
        } while ($col > 0);

        return $ret;
    }
}