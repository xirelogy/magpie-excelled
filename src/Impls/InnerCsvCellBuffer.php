<?php

namespace MagpieLib\Excelled\Impls;

use DateTimeInterface;
use Magpie\Codecs\Formats\Formatter;
use Magpie\Exceptions\NullException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\UnsupportedValueException;
use Magpie\General\Sugars\Quote;
use MagpieLib\Excelled\Codecs\Parsers\ExcelDateTimeParser;

/**
 * Cell buffer for CSV
 * @internal
 */
class InnerCsvCellBuffer
{
    /**
     * @var mixed Saved value
     */
    public readonly mixed $value;
    /**
     * @var Formatter|null
     */
    public ?Formatter $formatter;


    /**
     * Constructor
     * @param mixed $value
     * @param Formatter|null $formatter
     */
    public function __construct(mixed $value, ?Formatter $formatter)
    {
        $this->value = $value;
        $this->formatter = $formatter;
    }


    /**
     * Finalize the cell
     * @param string|null $formatString
     * @return string|null
     * @throws SafetyCommonException
     */
    public function finalize(?string $formatString) : ?string
    {
        _throwable() ?? throw new NullException();
        _used($formatString);

        if ($this->value === null) return null;

        if (is_string($this->value)) {
            // String is string
            $value = $this->value;
            if ($this->formatter !== null) {
                $value = $this->formatter->format($value);
            }

            return Quote::double(static::escapeCsvString($value));
        }

        if (is_numeric($this->value)) {
            // Numerical value
            $value = $this->value;
            if ($this->formatter !== null) {
                $value = $this->formatter->format($value);
            }

            return $value;
        }

        if (is_bool($this->value)) {
            // Boolean value
            $value = $this->value;
            if ($this->formatter !== null) {
                $value = $this->formatter->format($value);
            }

            // Hard conversion
            if (is_bool($value)) $value = $value ? 1 : 0;

            return $value;
        }

        if ($this->value instanceof DateTimeInterface) {
            $value = $this->value;
            if ($this->formatter !== null) {
                $value = $this->formatter->format($value);

                if (is_numeric($value)) {
                    $value = ExcelDateTimeParser::create()->parse($value);
                }
            }

            if ($value instanceof DateTimeInterface) {
                $value = $value->format('Y-m-d H:i:s');
            }

            return $value;
        }

        throw new UnsupportedValueException($this->value);
    }


    /**
     * Escape given string according to CSV rules
     * @param string $value
     * @return string
     */
    protected static function escapeCsvString(string $value) : string
    {
        $ret = '';
        $length = strlen($value);

        for ($i = 0; $i < $length; ++$i) {
            $c = substr($value, $i, 1);
            if ($c == '"') {
                $ret .= '""';
            } else {
                $ret .= $c;
            }
        }

        return $ret;
    }
}