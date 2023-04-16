<?php

namespace MagpieLib\Excelled\Objects\Styles;

use Magpie\Exceptions\SafetyCommonException;
use Magpie\General\Concepts\TypeClassable;
use PhpOffice\PhpSpreadsheet\Style\Font as PhpOfficeFont;

/**
 * Excel style
 */
abstract class ExcelStyle implements TypeClassable
{
    /**
     * Constructor
     */
    protected function __construct()
    {

    }


    /**
     * Apply current style on underlying font object
     * @param PhpOfficeFont $font
     * @return void
     * @throws SafetyCommonException
     * @internal
     */
    public final function _applyOnFont(PhpOfficeFont $font) : void
    {
        $this->onApplyOnFont($font);
    }


    /**
     * Apply current style on underlying font object
     * @param PhpOfficeFont $font
     * @return void
     * @throws SafetyCommonException
     */
    protected abstract function onApplyOnFont(PhpOfficeFont $font) : void;
}