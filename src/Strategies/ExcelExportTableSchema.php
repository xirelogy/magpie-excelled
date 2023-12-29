<?php

namespace MagpieLib\Excelled\Strategies;

use Closure;
use Magpie\Exceptions\NullException;
use Magpie\Exceptions\SafetyCommonException;
use MagpieLib\Excelled\Concepts\TableExportable;
use MagpieLib\Excelled\Objects\ColumnDefinition;

/**
 * A schema to export Excel as a table
 */
abstract class ExcelExportTableSchema extends ExcelExportSheetSchema implements TableExportable
{
    /**
     * @var Closure|null Column index resolver
     */
    private ?Closure $columnIndexResolverFn = null;


    /**
     * Get column index for given definition
     * @param ColumnDefinition $definition
     * @return int|null
     * @throws SafetyCommonException
     */
    protected function getColumnIndex(ColumnDefinition $definition) : ?int
    {
        if ($this->columnIndexResolverFn === null) return null;

        _throwable() ?? throw new NullException();

        return ($this->columnIndexResolverFn)($definition);
    }


    /**
     * Set the column index resolver
     * @param (callable(ColumnDefinition):int|null)|null $fn
     * @return void
     * @internal
     */
    public function _setColumnIndexResolver(?callable $fn) : void
    {
        $this->columnIndexResolverFn = $fn;
    }
}
