<?php

namespace MagpieLib\Excelled\Objects\Styles;

use MagpieLib\Excelled\Objects\ExcelColor;
use PhpOffice\PhpSpreadsheet\Style\Font as PhpOfficeFont;

class ExcelColorStyle extends ExcelStyle
{
    /**
     * Current type class
     */
    public const TYPECLASS = 'color';

    /**
     * @var ExcelColor Associated color
     */
    public readonly ExcelColor $color;


    /**
     * Constructor
     * @param ExcelColor $color
     */
    protected function __construct(ExcelColor $color)
    {
        parent::__construct();

        $this->color = $color;
    }


    /**
     * @inheritDoc
     */
    protected function onApplyOnFont(PhpOfficeFont $font) : void
    {
        $font->getColor()->setARGB($this->color->getColorString());
    }


    /**
     * @inheritDoc
     */
    public static function getTypeClass() : string
    {
        return static::TYPECLASS;
    }


    /**
     * Create an instance
     * @param ExcelColor $color
     * @return static
     */
    public static function create(ExcelColor $color) : static
    {
        return new static($color);
    }
}