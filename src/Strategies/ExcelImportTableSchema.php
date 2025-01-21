<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Exceptions\InvalidStateException;
use Magpie\Exceptions\NullException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\General\Traits\StaticCreatable;
use MagpieLib\Excelled\Concepts\TableImportable;
use MagpieLib\Excelled\Impls\ExcelArrayParserHost;

/**
 * A schema to import Excel as a table
 * @template T Type of the resulting import
 * @implements TableImportable<T>
 */
abstract class ExcelImportTableSchema implements TableImportable
{
    use StaticCreatable;

    /**
     * @var array|null Header map
     */
    private ?array $headerMap = null;


    /**
     * The starting row's index
     * @return int
     */
    protected function getStartRowIndex() : int
    {
        return 0;
    }


    /**
     * The starting column's index
     * @return int
     */
    protected function getStartColumnIndex() : int
    {
        return 0;
    }


    /**
     * The size of header row
     * @return int
     */
    protected function getHeaderRowSize() : int
    {
        return 1;
    }


    /**
     * Import from given source
     * @param CommonImporter $source
     * @param ExcelImportOptions|null $options
     * @return iterable<T>
     * @throws SafetyCommonException
     */
    public final function import(CommonImporter $source, ?ExcelImportOptions $options = null) : iterable
    {
        $options = $options ?? ExcelImportOptions::default();

        $service = $source->_getService();
        $sheetService = $service->getSheet($options);

        $startRowIndex = $this->getStartRowIndex();
        $startColIndex = $this->getStartColumnIndex();

        $thisRowIndex = $startRowIndex - 1;
        $headerRowCount = $this->getHeaderRowSize();
        foreach ($sheetService->getRows($startRowIndex, $startColIndex, null, $headerRowCount) as $rowArray) {
            ++$thisRowIndex;
            if ($headerRowCount > 0) {
                $this->onHeaderRow($rowArray);
                --$headerRowCount;
            } else {
                $parserHost = new ExcelArrayParserHost($rowArray, $thisRowIndex, $this->headerMap, $options);
                yield $this->parseRow($parserHost);
            }
        }
    }


    /**
     * Get notified on header row
     * @param array $values
     * @return void
     * @throws SafetyCommonException
     */
    protected function onHeaderRow(array $values) : void
    {
        _throwable() ?? throw new NullException();

        _used($values);
    }


    /**
     * Map header rows according to column schema
     * @param ExcelImportTableColumnsSchema $columns
     * @param array $headerValues
     * @return void
     * @throws SafetyCommonException
     */
    protected function mapHeaderRow(ExcelImportTableColumnsSchema $columns, array $headerValues) : void
    {
        if ($this->headerMap !== null) throw new InvalidStateException();

        $this->headerMap = $columns->_processHeader($headerValues);
    }
}
