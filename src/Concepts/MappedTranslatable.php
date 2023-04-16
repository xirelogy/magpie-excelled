<?php

namespace MagpieLib\Excelled\Concepts;

use Magpie\Exceptions\PersistenceException;
use Magpie\Exceptions\SafetyCommonException;
use MagpieLib\Excelled\Objects\ColumnDefinition;

/**
 * May translate an item (row) with mappings
 */
interface MappedTranslatable
{
    /**
     * Translate item with mappings
     * @param mixed $row
     * @return iterable<ColumnDefinition, mixed>
     * @throws SafetyCommonException
     * @throws PersistenceException
     */
    public function mappedTranslate(mixed $row) : iterable;
}