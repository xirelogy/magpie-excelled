<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Exceptions\GeneralPersistenceException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\UnsupportedValueException;
use Magpie\Facades\FileSystem\Providers\Local\LocalFileWriteTarget;
use Magpie\Facades\Mime\Mime;
use Magpie\General\Concepts\TargetReadable;
use Magpie\General\Concepts\TargetWritable;
use Magpie\General\Contexts\ScopedCollection;
use MagpieLib\Excelled\Concepts\ExcelFormatterAdaptable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetExportServiceable;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PhpOfficeSpreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception as PhpOfficeWriterException;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as PhpOfficeWriterXlsx;

/**
 * Default export service
 * @internal
 */
class DefaultExcelExportService extends CommonExcelExportService
{
    /**
     * @var PhpOfficeSpreadsheet Workbook
     */
    protected PhpOfficeSpreadsheet $workbook;


    /**
     * Constructor
     * @param ExcelFormatterAdaptable $formatAdapter
     * @param TargetWritable $target
     * @throws SafetyCommonException
     */
    public function __construct(ExcelFormatterAdaptable $formatAdapter, TargetWritable $target)
    {
        $this->workbook = new PhpOfficeSpreadsheet();

        // Clear any existing default worksheets
        while (count($this->workbook->getAllSheets()) > 0) {
            OfficeExcepts::protect(fn () => $this->workbook->removeSheetByIndex(0));
        }

        parent::__construct($formatAdapter, $target);
    }


    /**
     * @inheritDoc
     */
    public function load(TargetReadable $target) : void
    {
        $this->workbook = ExcelIO::readWorkbookFromTarget($target, false);
    }


    /**
     * @inheritDoc
     */
    public function accessSheet(string $sheetName) : ExcelSheetExportServiceable
    {
        $worksheet = $this->workbook->getSheetByName($sheetName) ?? throw new UnsupportedValueException($sheetName);
        return new DefaultExcelSheetExportService($this, $worksheet, $this->formatAdapter);
    }


    /**
     * @inheritDoc
     */
    public function createSheet(?string $sheetName) : ExcelSheetExportServiceable
    {
        $worksheet = $this->workbook->createSheet();
        if ($sheetName !== null) $worksheet->setTitle($sheetName);
        return new DefaultExcelSheetExportService($this, $worksheet, $this->formatAdapter);
    }


    /**
     * @inheritDoc
     */
    protected function onFinalize(?string &$mimeType = null) : void
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
}