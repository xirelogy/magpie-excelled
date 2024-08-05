<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Codecs\Parsers\Parser;

/**
 * ExcelImportTableSchema column from attribute
 * @internal
 */
class ExcelImportTableSchemaColumnFromAttribute
{
    /**
     * @var string Property name
     */
    public readonly string $propertyName;
    /**
     * @var string Column name as in Excel
     */
    public readonly string $columnName;
    /**
     * @var array Alternative column names as in Excel
     */
    public readonly array $altColumnNames;
    /**
     * @var bool If column required
     */
    public readonly bool $isRequired;
    /**
     * @var Parser|null Associated parser for value (if any)
     */
    public readonly ?Parser $valueParser;
    /**
     * @var bool|null If value required
     */
    public readonly ?bool $isValueRequired;
    /**
     * @var int|null Index on constructor (if to be hydrated via constructor)
     */
    public ?int $constructorIndex = null;


    /**
     * Constructor
     * @param string $propertyName
     * @param string $columnName
     * @param iterable<string> $altColumnNames
     * @param bool $isRequired
     * @param Parser|null $valueParser
     * @param bool|null $isValueRequired
     */
    public function __construct(string $propertyName, string $columnName, iterable $altColumnNames, bool $isRequired, ?Parser $valueParser, ?bool $isValueRequired)
    {
        $this->propertyName = $propertyName;
        $this->columnName = $columnName;
        $this->altColumnNames = iter_flatten($altColumnNames, false);
        $this->isRequired = $isRequired;
        $this->valueParser = $valueParser;
        $this->isValueRequired = $isValueRequired;
    }
}