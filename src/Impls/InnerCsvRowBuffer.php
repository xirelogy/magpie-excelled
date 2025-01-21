<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Codecs\Formats\Formatter;
use Magpie\Exceptions\SafetyCommonException;

/**
 * Row buffer for CSV
 * @internal
 */
class InnerCsvRowBuffer
{
    /**
     * @var int Maximum column
     */
    protected int $maxCol = -1;
    /**
     * @var array<int, InnerCsvCellBuffer> Column value map
     */
    protected array $columnValues = [];


    /**
     * Set value
     * @param int $col
     * @param mixed $value
     * @param Formatter|null $formatter
     * @return void
     */
    public function setValue(int $col, mixed $value, ?Formatter $formatter) : void
    {
        $this->maxCol = max($this->maxCol, $col);
        $this->columnValues[$col] = new InnerCsvCellBuffer($value, $formatter);
    }


    /**
     * Finalize the current row
     * @param array<int, string> $colFormatStrings
     * @return string
     * @throws SafetyCommonException
     */
    public function finalize(array $colFormatStrings) : string
    {
        $retValues = [];
        $maxCol = $this->maxCol + 1;

        for ($i = 0; $i < $maxCol; ++$i) {
            $cell = $this->columnValues[$i] ?? null;
            $retValues[] = $cell?->finalize($colFormatStrings[$i] ?? null) ?? '';
        }

        return implode(',', $retValues);
    }
}