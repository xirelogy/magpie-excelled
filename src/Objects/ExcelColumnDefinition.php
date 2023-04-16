<?php

namespace MagpieLib\Excelled\Objects;

use Magpie\Codecs\Formats\Formatter;
use MagpieLib\Excelled\Constants\ExcelCellFormat;
use MagpieLib\Excelled\Objects\Shims\ExcelColumnAutoSize;
use MagpieLib\Excelled\Objects\Shims\ExcelColumnDefaultSize;

/**
 * Column definition with Excel related information
 */
class ExcelColumnDefinition extends ColumnDefinition
{
    /**
     * @var string Format string used in Excel
     */
    public string $excelFormatString;


    /**
     * Constructor
     * @param string $name
     * @param Formatter|null $format
     * @param string $excelFormatString
     */
    public function __construct(string $name, ?Formatter $format = null, string $excelFormatString = ExcelCellFormat::GENERAL, float|ExcelColumnAutoSize|ExcelColumnDefaultSize|null $setWidth = null)
    {
        parent::__construct($name, $format, $setWidth);

        $this->excelFormatString = $excelFormatString;
    }


    /**
     * Adapt column definition with specific Excel format string
     * @param ColumnDefinition $baseDefinition
     * @param string $excelFormatString
     * @return static
     */
    public static function adapt(ColumnDefinition $baseDefinition, string $excelFormatString) : static
    {
        $ret = new static($baseDefinition->name, $baseDefinition->format, $excelFormatString, $baseDefinition->setWidth);
        $ret->id = $baseDefinition->id; // Copy ID

        return $ret;
    }
}