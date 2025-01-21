<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Codecs\ParserHosts\ArrayParserHost;
use MagpieLib\Excelled\Strategies\ExcelImportOptions;
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
     * @var ExcelImportOptions|null Import options (Excel)
     */
    protected readonly ?ExcelImportOptions $excelOptions;


    /**
     * Constructor
     * @param array $arr
     * @param int $refRowIndex
     * @param array<string, int>|null $namedIndices
     * @param ExcelImportOptions|null $excelOptions
     * @param string|null $prefix
     */
    public function __construct(array $arr, int $refRowIndex, ?array $namedIndices, ?ExcelImportOptions $excelOptions, ?string $prefix = null)
    {
        parent::__construct($arr, $prefix);

        $this->refRowIndex = $refRowIndex;
        $this->namedIndices = $namedIndices;
        $this->excelOptions = $excelOptions;
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


    /**
     * @inheritDoc
     */
    protected function obtainRaw(int|string $key, int|string $inKey, bool $isMandatory, mixed $default) : mixed
    {
        $ret = parent::obtainRaw($key, $inKey, $isMandatory, $default);

        // Extra string cleanup
        $stringCleanup = $this->excelOptions?->stringCleanup;
        if (is_string($ret) && $stringCleanup !== null) {
            if (str_starts_with($ret, $stringCleanup)) {
                $ret = substr($ret, strlen($stringCleanup));
            }
        }

        return $ret;
    }
}