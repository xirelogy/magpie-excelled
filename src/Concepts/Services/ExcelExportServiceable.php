<?php

namespace MagpieLib\Excelled\Concepts\Services;

use Magpie\Exceptions\PersistenceException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\StreamException;
use MagpieLib\Excelled\Objects\ColumnDefinition;
use MagpieLib\Excelled\Objects\ExcelColumnDefinition;

/**
 * Service interface to export to Excel
 */
interface ExcelExportServiceable extends ExcelResourceManageable
{
    /**
     * Create a new sheet
     * @param string|null $sheetName
     * @return ExcelSheetExportServiceable
     * @throws SafetyCommonException
     */
    public function createSheet(?string $sheetName) : ExcelSheetExportServiceable;


    /**
     * Finalize the output
     * @param string|null $mimeType
     * @return void
     * @throws SafetyCommonException
     * @throws PersistenceException
     * @throws StreamException
     */
    public function finalize(?string &$mimeType = null) : void;


    /**
     * Adapt column definition according to Excel's need
     * @param ColumnDefinition $column
     * @return ExcelColumnDefinition
     * @throws SafetyCommonException
     */
    public function adaptColumnDefinition(ColumnDefinition $column) : ExcelColumnDefinition;
}