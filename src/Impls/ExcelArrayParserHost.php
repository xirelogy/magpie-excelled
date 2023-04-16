<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Codecs\ParserHosts\ArrayParserHost;
use MagpieLib\Excelled\Strategies\ExcelNames;

/**
 * Array parser host adapted to Excel
 * @internal
 */
class ExcelArrayParserHost extends ArrayParserHost
{
    /**
     * @var int Reference row index
     */
    protected readonly int $refRowIndex;
    /**
     * @var array|null Named indices
     */
    protected readonly ?array $namedIndices;


    /**
     * Constructor
     * @param array $arr
     * @param int $refRowIndex
     * @param array<string, int>|null $namedIndices
     * @param string|null $prefix
     */
    public function __construct(array $arr, int $refRowIndex, ?array $namedIndices, ?string $prefix = null)
    {
        parent::__construct($arr, $prefix);

        $this->refRowIndex = $refRowIndex;
        $this->namedIndices = $namedIndices;
    }


    /**
     * @inheritDoc
     */
    protected function acceptKey(int|string $key) : int
    {
        if (is_int($key)) return $key;

        if ($this->namedIndices !== null) {
            return $this->namedIndices[$key] ?? -1;
        }

        return -1;
    }


    /**
     * @inheritDoc
     */
    public function fullKey(int|string $key) : string
    {
        if (is_string($key)) return parent::fullKey($key);
        if ($key < 0) return '<err>';

        return $this->prefix . ExcelNames::cellNameOf($this->refRowIndex, $key);
    }
}