<?php

namespace MagpieLib\Excelled\Codecs\Formats;

use Magpie\Codecs\Formats\Formatter;

/**
 * Extends existing formatter to be compliant with ExcelFormatter
 */
class ExcelUsingFormatter implements ExcelFormatter
{
    /**
     * @var Formatter Base formatter
     */
    protected readonly Formatter $baseFormatter;
    /**
     * @var string Excel format string
     */
    protected readonly string $excelFormatString;


    /**
     * Constructor
     * @param Formatter $baseFormatter
     * @param string $excelFormatString
     */
    protected function __construct(Formatter $baseFormatter, string $excelFormatString)
    {
        $this->baseFormatter = $baseFormatter;
        $this->excelFormatString = $excelFormatString;
    }


    /**
     * @inheritDoc
     */
    public function format(mixed $value) : mixed
    {
        return $this->baseFormatter->format($value);
    }


    /**
     * @inheritDoc
     */
    public function getExcelFormatString() : ?string
    {
        return $this->excelFormatString;
    }


    /**
     * Create formatter by extension
     * @param Formatter $baseFormatter
     * @param string $excelFormatString
     * @return static
     */
    public static function extends(Formatter $baseFormatter, string $excelFormatString) : static
    {
        return new static($baseFormatter, $excelFormatString);
    }
}