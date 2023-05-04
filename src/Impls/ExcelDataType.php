<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\General\Traits\StaticClass;
use MagpieLib\Excelled\Constants\ExcelCellFormat;
use PhpOffice\PhpSpreadsheet\Cell\DataType as PhpOfficeCellDataType;
use PhpOffice\PhpSpreadsheet\Cell\IValueBinder as PhpOfficeIValueBinder;
use PhpOffice\PhpSpreadsheet\Cell\StringValueBinder as PhpOfficeStringValueBinder;

/**
 * Support for translation to Excel's data type
 * @internal
 */
class ExcelDataType
{
    use StaticClass;


    /**
     * Create a strict value binder for strings
     * @return PhpOfficeIValueBinder
     */
    public static function createStrictStringValueBinder() : PhpOfficeIValueBinder
    {
        $ret = new PhpOfficeStringValueBinder();
        $ret->setNumericConversion(true);
        $ret->setBooleanConversion(true);

        return $ret;
    }


    /**
     * Translate data type
     * @param string|null $excelFormatString
     * @return string|null
     */
    public static function getDataType(?string $excelFormatString) : ?string
    {
        if ($excelFormatString === null) return null;

        return match ($excelFormatString) {
            ExcelCellFormat::TEXT,
                => PhpOfficeCellDataType::TYPE_STRING2,
            ExcelCellFormat::DATE_ISO,
            ExcelCellFormat::DATETIME_ISO,
            ExcelCellFormat::NUMBER,
            ExcelCellFormat::NUMBER_00,
            ExcelCellFormat::NUMBER_COMMA_00,
                => PhpOfficeCellDataType::TYPE_NUMERIC,
            default,
                => null,
        };
    }
}