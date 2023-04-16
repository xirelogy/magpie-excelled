<?php

namespace MagpieLib\Excelled\Constants;

use Magpie\General\Traits\StaticClass;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat as PhpOfficeNumberFormat;

/**
 * Excel cell format
 */
class ExcelCellFormat
{
    use StaticClass;

    /**
     * General format
     */
    public const GENERAL = PhpOfficeNumberFormat::FORMAT_GENERAL;
    /**
     * Text format
     */
    public const TEXT = PhpOfficeNumberFormat::FORMAT_TEXT;
    /**
     * Number format
     */
    public const NUMBER = PhpOfficeNumberFormat::FORMAT_NUMBER;
    /**
     * Number format with 2 decimals
     */
    public const NUMBER_00 = PhpOfficeNumberFormat::FORMAT_NUMBER_00;
    /**
     * Number format with 2 decimals, comma separated
     */
    public const NUMBER_COMMA_00 = PhpOfficeNumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1;
    /**
     * Percentage format
     */
    public const PERCENTAGE = PhpOfficeNumberFormat::FORMAT_PERCENTAGE;
    /**
     * Percentage format with 2 decimals
     */
    public const PERCENTAGE_00 = PhpOfficeNumberFormat::FORMAT_PERCENTAGE_00;
    /**
     * ISO date
     */
    public const DATE_ISO = 'yyyy-mm-dd';
    /**
     * ISO date-time
     */
    public const DATETIME_ISO = 'yyyy-mm-dd hh:mm:ss';
}