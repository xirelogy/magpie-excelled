<?php

namespace MagpieLib\Excelled\Concepts\Services;

use Magpie\Exceptions\SafetyCommonException;

/**
 * Service interface to import from Excel sheet
 */
interface ExcelSheetImportServiceable
{
    /**
     * Get all rows from given scope
     * @param int $startRowIndex
     * @param int $startColIndex
     * @param int|null $endColIndex
     * @param int $lockRowCount
     * @return iterable<array>
     * @throws SafetyCommonException
     */
    public function getRows(int $startRowIndex, int $startColIndex, ?int $endColIndex = null, int $lockRowCount = 1) : iterable;
}