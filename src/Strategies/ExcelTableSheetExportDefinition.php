<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Codecs\Formats\Formatter;
use Magpie\Exceptions\DuplicatedKeyException;
use Magpie\Exceptions\PersistenceException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\UnsupportedValueException;
use Magpie\Facades\Random;
use Magpie\General\Randoms\RandomCharset;
use Magpie\General\Str;
use MagpieLib\Excelled\Concepts\MappedTranslatable;
use MagpieLib\Excelled\Concepts\Services\ExcelExportServiceable;
use MagpieLib\Excelled\Impls\ColumnMetadata;
use MagpieLib\Excelled\Objects\ColumnDefinition;
use MagpieLib\Excelled\Objects\Shims\ExcelColumnAutoSize;
use MagpieLib\Excelled\Objects\Styles\ExcelBoldStyle;
use MagpieLib\Excelled\Objects\Styles\ExcelStyle;

/**
 * Definition of an Excel table export to a sheet
 */
class ExcelTableSheetExportDefinition extends ExcelSheetExportDefinition
{
    /**
     * @var ExcelExportTableSchema Associated table schema
     */
    protected readonly ExcelExportTableSchema $schema;
    /**
     * @var iterable Rows of items
     */
    protected readonly iterable $rows;


    /**
     * Constructor
     * @param ExcelExportTableSchema $schema
     * @param iterable $rows
     * @param string|null $sheetName
     */
    protected function __construct(ExcelExportTableSchema $schema, iterable $rows, ?string $sheetName)
    {
        parent::__construct($sheetName);

        $this->schema = $schema;
        $this->rows = $rows;
    }


    /**
     * @inheritDoc
     */
    protected function onRun(ExcelExportServiceable $service, ?string &$mimeType = null) : void
    {
        /** @var array<int, string> $columnIds */
        $columnIds = [];
        /** @var array<string, ColumnMetadata> $columns */
        $columns = [];

        // Process the schema
        $columnIndex = 0;
        foreach ($this->schema->getColumns() as $column) {
            if (Str::isNullOrEmpty($column->id)) {
                $column->id = Random::string(8, RandomCharset::LOWER_ALPHANUM);
            }

            if (array_key_exists($column->id, $columns)) throw new DuplicatedKeyException($column->id);

            $column = $service->adaptColumnDefinition($column);
            $columnIds[$columnIndex] = $column->id;
            $columns[$column->id] = new ColumnMetadata($column, $columnIndex);

            ++$columnIndex;
        }


        try {
            $this->schema->_setColumnIndexResolver(function (ColumnDefinition $def) use ($columns) : ?int {
                if ($def->id === null) return null;
                if (!array_key_exists($def->id, $columns)) return null;

                $columnMetadata = $columns[$def->id];
                return $columnMetadata->index;
            });

            // Create sheet
            $sheetService = $service->createSheet($this->sheetName);

            // Output the header
            $currentRowIndex = 0;
            foreach ($columns as $columnMetadata) {
                $sheetService->accessCell($currentRowIndex, $columnMetadata->index)->setValue($columnMetadata->definition->name);
                if (!Str::isNullOrEmpty($columnMetadata->definition->excelFormatString)) $sheetService->accessColumn($columnMetadata->index)->setFormat($columnMetadata->definition->excelFormatString);
            }
            ++$currentRowIndex;

            $totalHeaderRows = $currentRowIndex;

            // Process and output the rows
            foreach ($this->rows as $row) {
                $rowCells = $this->translateRow($columns, $row);
                $currentColumnIndex = 0;
                foreach ($rowCells as $cell) {
                    $columnFormat = static::getColumnFormat($columns, $columnIds, $currentColumnIndex);
                    $sheetService->accessCell($currentRowIndex, $currentColumnIndex)->setValue($cell, $columnFormat);
                    ++$currentColumnIndex;
                }
                ++$currentRowIndex;
            }

            // Resize columns (default to auto-size)
            foreach ($columns as $columnMetadata) {
                $sheetService->accessColumn($columnMetadata->index)->setWidth($columnMetadata->definition->setWidth ?? ExcelColumnAutoSize::create());
            }

            // Set header styles
            $headerStyles = iter_flatten($this->getHeaderStyles(), false);
            for ($row = 0; $row < $totalHeaderRows; ++$row) {
                $sheetService->accessRow($row)->applyStyle(...$headerStyles);
            }

            // Common freeze
            $sheetService->freezePane($totalHeaderRows, 0);

            // Finalize
            $this->schema->finalizeSheet($sheetService);
            $service->finalize($mimeType);
        } finally {
            $this->schema->_setColumnIndexResolver(null);
        }
    }


    /**
     * Get column format
     * @param array<string, ColumnMetadata> $columns
     * @param array<int, string> $columnIds
     * @param int $columnIndex
     * @return Formatter|null
     */
    private static function getColumnFormat(array $columns, array $columnIds, int $columnIndex) : ?Formatter
    {
        $columnId = $columnIds[$columnIndex] ?? null;
        if ($columnId === null) return null;

        $column = $columns[$columnId] ?? null;
        return $column?->definition?->format;
    }


    /**
     * Header styles
     * @return iterable<ExcelStyle>
     */
    protected function getHeaderStyles() : iterable
    {
        yield ExcelBoldStyle::create();
    }


    /**
     * Translate a row of data
     * @param array<string, ColumnMetadata> $columns
     * @param mixed $row
     * @return array
     * @throws SafetyCommonException
     * @throws PersistenceException
     */
    private function translateRow(array $columns, mixed $row) : array
    {
        // Handle mapped translation
        if ($this->schema instanceof MappedTranslatable) {
            $retCells = static::generateEmptyCells($columns);

            /**
             * @var ColumnDefinition $column
             * @var mixed $value
             */
            foreach ($this->schema->mappedTranslate($row) as $column => $value) {
                $columnMetadata = $columns[$column->id ?? ''] ?? throw new UnsupportedValueException($column, _l('column'));
                $retCells[$columnMetadata->index] = $value;
            }

            return $retCells;
        }

        // Otherwise, fallback
        return iter_flatten($this->schema->translate($row), false);
    }


    /**
     * Generate empty cells
     * @param array<string, ColumnMetadata> $columns
     * @return array
     */
    private static function generateEmptyCells(array $columns) : array
    {
        $ret = [];
        foreach ($columns as $column) {
            _used($column);
            $ret[] = null;
        }

        return $ret;
    }


    /**
     * Create a new instance
     * @param ExcelExportTableSchema $schema
     * @param iterable $rows
     * @param string|null $sheetName
     * @return static
     */
    public static function create(ExcelExportTableSchema $schema, iterable $rows, ?string $sheetName = null) : static
    {
        return new static($schema, $rows, $sheetName);
    }
}