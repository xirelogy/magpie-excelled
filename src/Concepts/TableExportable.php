<?php

namespace MagpieLib\Excelled\Concepts;

use MagpieLib\Excelled\Objects\ColumnDefinition;

/**
 * Table-like export schema
 */
interface TableExportable extends Translatable
{
    /**
     * Column definitions
     * @return iterable<ColumnDefinition>
     */
    public function getColumns() : iterable;
}