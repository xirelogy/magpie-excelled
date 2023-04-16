<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Exceptions\UnsupportedException;
use MagpieLib\Excelled\Concepts\MappedTranslatable;

/**
 * A schema to export Excel as a table, with column mappings
 */
abstract class ExcelMappedExportTableSchema extends ExcelExportTableSchema implements MappedTranslatable
{
    /**
     * @inheritDoc
     */
    public final function translate(mixed $row) : iterable
    {
        // When mapped, the original translate is no longer supported
        throw new UnsupportedException();
    }
}
