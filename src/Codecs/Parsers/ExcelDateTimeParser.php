<?php

namespace MagpieLib\Excelled\Codecs\Parsers;

use Carbon\Carbon;
use Carbon\CarbonInterface;
use Magpie\Codecs\Parsers\DateTimeParser;

/**
 * Parse for date/time from Excel
 */
class ExcelDateTimeParser extends DateTimeParser
{
    /**
     * @inheritDoc
     */
    protected function onParse(mixed $value, ?string $hintName) : ?CarbonInterface
    {
        if (is_float($value)) {
            // Process as date-time serial
            $daySerial = intval(floor($value));
            $timeSerial = $value - $daySerial;

            $value = static::parseDateSerial($daySerial) . ' ' . static::parseDateSerialTime($timeSerial);
        } else if (is_int($value)) {
            $value = static::parseDateSerial($value);
        }

        return parent::onParse($value, $hintName);
    }


    /**
     * Convert date serial into date string
     * @param int $value
     * @return string
     */
    private static function parseDateSerial(int $value) : string
    {
        // Adapt to special case of '1900-02-29'
        if ($value > 60) --$value;

        $dt = Carbon::parse('1900-01-01')->addDays($value - 1);
        return $dt->format('Y-m-d');
    }


    /**
     * Convert date serial's time portion into time string
     * @param float $value
     * @return string
     */
    private static function parseDateSerialTime(float $value) : string
    {
        $timeOfDay = intval(round(86400 * $value));
        if ($timeOfDay < 0) $timeOfDay = 0;
        if ($timeOfDay >= 86400) $timeOfDay = 86399;    // Safety cap

        // Break into hour-minute-seconds
        $hour = intdiv($timeOfDay, 3600);
        $timeOfDay -= ($hour * 3600);
        $minute = intdiv($timeOfDay, 60);
        $timeOfDay -= ($minute * 60);
        $second = $timeOfDay;

        $dt = Carbon::createFromTime($hour, $minute, $second);
        return $dt->format('H:i:s');
    }
}