<?php

namespace MagpieLib\Excelled\Concepts\Services;

use Magpie\Exceptions\PersistenceException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\StreamException;

/**
 * Service interface to export to Excel sheet
 */
interface ExcelSheetExportServiceable
{
    /**
     * Access a row
     * @param int $row
     * @return ExcelRowExportServiceable
     * @throws SafetyCommonException
     */
    public function accessRow(int $row) : ExcelRowExportServiceable;


    /**
     * Access a column
     * @param int $col
     * @return ExcelColumnExportServiceable
     * @throws SafetyCommonException
     */
    public function accessColumn(int $col) : ExcelColumnExportServiceable;


    /**
     * Access a cell
     * @param int $row
     * @param int $col
     * @param int|null $row2
     * @param int|null $col2
     * @return ExcelCellExportServiceable
     * @throws SafetyCommonException
     */
    public function accessCell(int $row, int $col, ?int $row2 = null, ?int $col2 = null) : ExcelCellExportServiceable;


    /**
     * Freeze pane at given location
     * @param int $row
     * @param int $col
     * @return void
     * @throws SafetyCommonException
     */
    public function freezePane(int $row, int $col) : void;


    /**
     * Finalize the sheet
     * @return void
     * @throws SafetyCommonException
     */
    public function finalize() : void;
}