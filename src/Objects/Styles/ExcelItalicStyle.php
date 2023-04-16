<?php

namespace MagpieLib\Excelled\Objects\Styles;

use PhpOffice\PhpSpreadsheet\Style\Font as PhpOfficeFont;

/**
 * Italic style
 */
class ExcelItalicStyle extends ExcelCreatableStyle
{
    /**
     * Current type class
     */
    public const TYPECLASS = 'italic';


    /**
     * @inheritDoc
     */
    protected function onApplyOnFont(PhpOfficeFont $font) : void
    {
        $font->setItalic(true);
    }


    /**
     * @inheritDoc
     */
    public static function getTypeClass() : string
    {
        return static::TYPECLASS;
    }
}