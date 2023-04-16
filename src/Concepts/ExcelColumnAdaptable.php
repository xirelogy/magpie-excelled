<?php

namespace MagpieLib\Excelled\Concepts;

use Magpie\Exceptions\SafetyCommonException;
use MagpieLib\Excelled\Objects\ColumnDefinition;
use MagpieLib\Excelled\Objects\ExcelColumnDefinition;

/**
 * May adapt a column definition according to Excel's need
 */
interface ExcelColumnAdaptable
{
    /**
     * Adapt column definition according to Excel's need
     * @param ColumnDefinition $column
     * @return ExcelColumnDefinition
     * @throws SafetyCommonException
     */
    public function adapt(ColumnDefinition $column) : ExcelColumnDefinition;
}