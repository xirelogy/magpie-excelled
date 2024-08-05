<?php

namespace MagpieLib\Excelled\Annotations;

use Attribute;

/**
 * Alternative column name during Excel import
 */
#[Attribute(Attribute::TARGET_PROPERTY | Attribute::IS_REPEATABLE)]
class ExcelImportTableColumnAlt
{
    /**
     * @var string Alternative column name usable during import
     */
    public string $altColumnName;


    /**
     * Constructor
     * @param string $altColumnName Alternative column name usable during import
     */
    public function __construct(
        string $altColumnName,
    ) {
        $this->altColumnName = $altColumnName;
    }
}