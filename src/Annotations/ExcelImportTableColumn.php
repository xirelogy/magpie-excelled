<?php

namespace MagpieLib\Excelled\Annotations;

use Attribute;

/**
 * Column header expected during Excel import
 */
#[Attribute(Attribute::TARGET_PROPERTY)]
class ExcelImportTableColumn
{
    /**
     * @var string Column name expected during import
     */
    public string $columnName;
    /**
     * @var bool If the column required (compulsory)
     */
    public bool $isRequired;


    /**
     * Constructor
     * @param string $columnName Column name expected during import
     * @param bool $isRequired If the column required (compulsory)
     */
    public function __construct(
        string $columnName,
        bool $isRequired,
    ) {
        $this->columnName = $columnName;
        $this->isRequired = $isRequired;
    }
}