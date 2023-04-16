<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Exceptions\SafetyCommonException;
use MagpieLib\Excelled\Concepts\Services\ExcelGeneralExportServiceable;
use MagpieLib\Excelled\Objects\Styles\ExcelStyle;
use PhpOffice\PhpSpreadsheet\Style\Style as PhpOfficeStyle;

/**
 * Default export service (general)
 * @internal
 */
abstract class DefaultExcelGeneralExportService implements ExcelGeneralExportServiceable
{
    /**
     * @inheritDoc
     */
    public final function setFormat(string $formatString) : void
    {
        OfficeExcepts::protect(
            fn () => $this->getStyle()->getNumberFormat()->setFormatCode($formatString)
        );
    }


    /**
     * @inheritDoc
     */
    public final function applyStyle(ExcelStyle ...$styles) : void
    {
        OfficeExcepts::protect(function () use ($styles) {
            foreach ($styles as $style) {
                $style->_applyOnFont($this->getStyle()->getFont());
            }
        });
    }


    /**
     * The associated style object
     * @return PhpOfficeStyle
     * @throws SafetyCommonException
     */
    protected abstract function getStyle() : PhpOfficeStyle;
}