<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Exceptions\DuplicatedKeyException;
use Magpie\Exceptions\MissingArgumentException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\UnexpectedException;
use Magpie\General\Str;
use Magpie\General\Traits\StaticCreatable;
use MagpieLib\Excelled\Impls\ImportTableColumnMetadata;
use MagpieLib\Excelled\Objects\ExcelImages;

/**
 * A schema to import Excel as a table according to defined columns
 */
class ExcelImportTableColumnsSchema
{
    use StaticCreatable;


    /**
     * @var array<ImportTableColumnMetadata> Mapped definitions
     */
    protected array $definitions = [];
    /**
     * @var array<string, int> Inwards map
     */
    protected array $inMap = [];


    /**
     * Define a column as required
     * @param string $columnName
     * @param array<string>|string|null $altNames
     * @return $this
     * @throws SafetyCommonException
     */
    public final function requires(string $columnName, array|string|null $altNames = null) : static
    {
        $this->defineColumn(true, $columnName, static::acceptAltNames($altNames));
        return $this;
    }


    /**
     * Define a column as optional
     * @param string $columnName
     * @param array<string>|string|null $altNames
     * @return $this
     * @throws SafetyCommonException
     */
    public final function optional(string $columnName, array|string|null $altNames = null) : static
    {
        $this->defineColumn(false, $columnName, static::acceptAltNames($altNames));
        return $this;
    }


    /**
     * Process header values to produce a map of column name key to given index
     * @param array $headerValues
     * @return array<string, int>
     * @throws SafetyCommonException
     * @internal
     */
    public final function _processHeader(array $headerValues) : array
    {
        /** @var array<string, int> $ret */
        $ret = [];

        foreach ($headerValues as $columnIndex => $headerValue) {
            $columnName = static::flattenHeaderAsColumnName($headerValue);
            if ($columnName === null) continue;
            if (!array_key_exists($columnName, $this->inMap)) continue;

            $definition = $this->definitions[$this->inMap[$columnName]] ?? throw new UnexpectedException();

            if (array_key_exists($definition->key, $ret)) throw new DuplicatedKeyException($columnName, _l('Mapped column name'));
            $ret[$definition->key] = $columnIndex;
        }

        // Check for compulsory requirements
        foreach ($this->definitions as $definition) {
            if (!$definition->isRequired) continue;
            if (!array_key_exists($definition->key, $ret)) throw new MissingArgumentException($definition->key);
        }

        return $ret;
    }


    /**
     * Flatten a cell value from header into a column name
     * @param mixed $headerValue
     * @return string|null
     */
    protected static function flattenHeaderAsColumnName(mixed $headerValue) : ?string
    {
        if ($headerValue === null) return null;
        if ($headerValue instanceof ExcelImages) return null; // Image is not supported

        $ret = "$headerValue";
        if (Str::isNullOrEmpty($ret)) return null;

        return $ret;
    }


    /**
     * Define a column
     * @param bool $isRequired
     * @param string $columnName
     * @param array<string> $altNames
     * @return void
     * @throws SafetyCommonException
     */
    private function defineColumn(bool $isRequired, string $columnName, array $altNames) : void
    {
        $defIndex = count($this->definitions);
        $this->definitions[] = new ImportTableColumnMetadata($columnName, $isRequired, $altNames);

        $this->defineInMap($columnName, $defIndex);
        foreach ($altNames as $altName) {
            $this->defineInMap($altName, $defIndex);
        }
    }


    /**
     * Define in map
     * @param string $text
     * @param int $defIndex
     * @return void
     * @throws SafetyCommonException
     */
    private function defineInMap(string $text, int $defIndex) : void
    {
        if (array_key_exists($text, $this->inMap)) throw new DuplicatedKeyException($text, _l('Candidate column name'));

        $this->inMap[$text] = $defIndex;
    }


    /**
     * Accept alternative names
     * @param array<string>|string|null $altNames
     * @return array<string>
     */
    protected static function acceptAltNames(array|string|null $altNames) : array
    {
        if ($altNames === null) return [];
        if (is_string($altNames)) return [$altNames];

        return $altNames;
    }
}