<?php

namespace MagpieLib\Excelled\Objects\Styles;

use PhpOffice\PhpSpreadsheet\Style\Font as PhpOfficeFont;

/**
 * Bold style
 */
class ExcelBoldStyle extends ExcelCreatableStyle
{
    /**
     * Current type class
     */
    public const TYPECLASS = 'bold';


    /**
     * @inheritDoc
     */
    protected function onApplyOnFont(PhpOfficeFont $font) : void
    {
        $font->setBold(true);
    }


    /**
     * @inheritDoc
     */
    public static function getTypeClass() : string
    {
        return static::TYPECLASS;
    }
}