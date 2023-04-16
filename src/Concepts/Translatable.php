<?php

namespace MagpieLib\Excelled\Concepts;

use Magpie\Exceptions\SafetyCommonException;

/**
 * May translate an item (row)
 */
interface Translatable
{
    /**
     * Translate item
     * @param mixed $row
     * @return iterable
     * @throws SafetyCommonException
     */
    public function translate(mixed $row) : iterable;
}