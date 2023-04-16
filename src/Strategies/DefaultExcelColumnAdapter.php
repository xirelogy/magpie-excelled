<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Codecs\Formats\StringFormatter;
use Magpie\Exceptions\UnsupportedValueException;
use MagpieLib\Excelled\Codecs\Formats\ExcelFormatter;
use MagpieLib\Excelled\Concepts\ExcelColumnAdaptable;
use MagpieLib\Excelled\Constants\ExcelCellFormat;
use MagpieLib\Excelled\Objects\ColumnDefinition;
use MagpieLib\Excelled\Objects\ExcelColumnDefinition;

/**
 * Default implementation of ExcelColumnAdaptable
 */
class DefaultExcelColumnAdapter implements ExcelColumnAdaptable
{
    /**
     * @inheritDoc
     */
    public function adapt(ColumnDefinition $column) : ExcelColumnDefinition
    {
        $format = $column->format;
        if ($format === null) {
            return ExcelColumnDefinition::adapt($column, ExcelCellFormat::GENERAL);
        }

        if ($format instanceof ExcelFormatter) {
            return ExcelColumnDefinition::adapt($column, $format->getExcelFormatString() ?? ExcelCellFormat::GENERAL);
        }

        if ($format instanceof StringFormatter) {
            return ExcelColumnDefinition::adapt($column, ExcelCellFormat::TEXT);
        }

        throw new UnsupportedValueException($column);
    }
}