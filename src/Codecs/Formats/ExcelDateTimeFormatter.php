<?php

namespace MagpieLib\Excelled\Codecs\Formats;

use Carbon\CarbonInterface;
use Magpie\General\DateTimes\SystemTimezone;
use MagpieLib\Excelled\Constants\ExcelCellFormat;

/**
 * Date+time format (for Excel)
 */
class ExcelDateTimeFormatter implements ExcelFormatter
{
    /**
     * @var string Timezone to format into
     */
    protected string $timezone;
    /**
     * @var string Excel format string
     */
    protected readonly string $excelFormatString;


    /**
     * Constructor
     * @param string $excelFormatString
     */
    protected function __construct(string $excelFormatString)
    {
        $this->timezone = SystemTimezone::default();
        $this->excelFormatString = $excelFormatString;
    }


    /**
     * Specify the timezone to format
     * @param string $timezone
     * @return $this
     */
    public final function withTimezone(string $timezone) : static
    {
        $this->timezone = $timezone;
        return $this;
    }


    /**
     * @inheritDoc
     */
    public function format(mixed $value) : mixed
    {
        if ($value instanceof CarbonInterface) {
            $value = $value->toImmutable()->setTimezone($this->timezone);
        }

        return $value;
    }


    /**
     * @inheritDoc
     */
    public function getExcelFormatString() : ?string
    {
        return $this->excelFormatString;
    }


    /**
     * Create a new instance
     * @param string $excelFormatString
     * @return static
     */
    public static function create(string $excelFormatString = ExcelCellFormat::DATETIME_ISO) : static
    {
        return new static($excelFormatString);
    }
}