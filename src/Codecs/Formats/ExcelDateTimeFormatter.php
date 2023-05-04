<?php

namespace MagpieLib\Excelled\Codecs\Formats;

use Carbon\Carbon;
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

            return static::getDateSerial($value->year, $value->month, $value->day)
                + static::getTimeSerial($value->hour, $value->minute, $value->second);
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


    /**
     * Convert date to date serial
     * @param int $year
     * @param int $month
     * @param int $day
     * @return int
     */
    protected static function getDateSerial(int $year, int $month, int $day) : int
    {
        // Compatibility with Excel 1900-02-29 bug
        if ($year === 1900 && $month === 2 && $day === 29) {
            return 60;
        }

        $dtRef = Carbon::parse('1900-01-01', tz: 'UTC');
        $dtNow = Carbon::create($year, $month, $day, tz: 'UTC');

        $interval = $dtRef->diff($dtNow);
        $ret = intval($interval->format('%a')) + 2;

        if ($ret < 60) --$ret;  // Date before 1900-02-29: compatibility

        return $ret;
    }


    /**
     * Convert time to time serial
     * @param int $hour
     * @param int $minute
     * @param int $second
     * @return float
     */
    protected static function getTimeSerial(int $hour, int $minute, int $second) : float
    {
        $timeOfDay = ($hour * 3600) + ($minute * 60) + $second;
        if ($timeOfDay < 0) $timeOfDay = 0;
        if ($timeOfDay >= 86400) $timeOfDay = 86399;    // Safety cap

        return $timeOfDay / 86400;
    }
}