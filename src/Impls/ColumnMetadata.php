<?php

namespace MagpieLib\Excelled\Impls;

use MagpieLib\Excelled\Objects\ExcelColumnDefinition;

/**
 * Column related metadata
 * @internal
 */
class ColumnMetadata
{
    /**
     * @var ExcelColumnDefinition Column definition
     */
    public readonly ExcelColumnDefinition $definition;
    /**
     * @var int Column index
     */
    public readonly int $index;


    /**
     * Constructor
     * @param ExcelColumnDefinition $definition
     * @param int $index
     */
    public function __construct(ExcelColumnDefinition $definition, int $index)
    {
        $this->definition = $definition;
        $this->index = $index;
    }
}
