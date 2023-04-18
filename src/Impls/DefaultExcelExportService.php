<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Exceptions\GeneralPersistenceException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\UnsupportedValueException;
use Magpie\Facades\FileSystem\Providers\Local\LocalFileWriteTarget;
use Magpie\Facades\Mime\Mime;
use Magpie\General\Concepts\TargetWritable;
use Magpie\General\Contexts\ScopedCollection;
use MagpieLib\Excelled\Concepts\ExcelColumnAdaptable;
use MagpieLib\Excelled\Concepts\Services\ExcelExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetExportServiceable;
use MagpieLib\Excelled\Objects\ColumnDefinition;
use MagpieLib\Excelled\Objects\ExcelColumnDefinition;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PhpOfficeSpreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception as PhpOfficeWriterException;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as PhpOfficeWriterXlsx;

/**
 * Default export service
 * @internal
 */
class DefaultExcelExportService implements ExcelExportServiceable
{
    /**
     * @var PhpOfficeSpreadsheet Workbook
     */
    protected PhpOfficeSpreadsheet $workbook;
    /**
     * @var ExcelColumnAdaptable Column adapter
     */
    protected readonly ExcelColumnAdaptable $columnAdapter;
    /**
     * @var TargetWritable Write target
     */
    protected readonly TargetWritable $target;


    /**
     * Constructor
     * @param ExcelColumnAdaptable $columnAdapter
     * @param TargetWritable $target
     * @throws SafetyCommonException
     */
    public function __construct(ExcelColumnAdaptable $columnAdapter, TargetWritable $target)
    {
        $this->workbook = new PhpOfficeSpreadsheet();

        // Clear any existing default worksheets
        while (count($this->workbook->getAllSheets()) > 0) {
            OfficeExcepts::protect(fn () => $this->workbook->removeSheetByIndex(0));
        }

        $this->columnAdapter = $columnAdapter;
        $this->target = $target;
    }


    /**
     * @inheritDoc
     */
    public function createSheet(?string $sheetName) : ExcelSheetExportServiceable
    {
        $worksheet = $this->workbook->createSheet();
        if ($sheetName !== null) $worksheet->setTitle($sheetName);
        return new DefaultExcelSheetExportService($worksheet);
    }


    /**
     * @inheritDoc
     */
    public function finalize(?string &$mimeType = null) : void
    {
        $writer = new PhpOfficeWriterXlsx($this->workbook);
        $mimeType = Mime::getMimeType('xlsx');

        if (!$this->target instanceof LocalFileWriteTarget) {
            throw new UnsupportedValueException($this->target, _l('write target'));
        }

        $scoped = new ScopedCollection($this->target->getScopes());
        _used($scoped);

        try {
            $writer->save($this->target->path);
        } catch (PhpOfficeWriterException $ex) {
            throw new GeneralPersistenceException(previous: $ex);
        }
    }


    /**
     * @inheritDoc
     */
    public function adaptColumnDefinition(ColumnDefinition $column) : ExcelColumnDefinition
    {
        if ($column instanceof ExcelColumnDefinition) return $column;

        return $this->columnAdapter->adapt($column);
    }
}