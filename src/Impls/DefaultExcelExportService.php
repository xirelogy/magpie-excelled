<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Exceptions\GeneralPersistenceException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\UnsupportedValueException;
use Magpie\Facades\FileSystem\Providers\Local\LocalFileWriteTarget;
use Magpie\Facades\Mime\Mime;
use Magpie\General\Concepts\Releasable;
use Magpie\General\Concepts\TargetWritable;
use Magpie\General\Contexts\ScopedCollection;
use Magpie\Objects\ReleasableCollection;
use MagpieLib\Excelled\Concepts\ExcelFormatterAdaptable;
use MagpieLib\Excelled\Concepts\Services\ExcelExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetExportServiceable;
use MagpieLib\Excelled\Constants\ExcelCellFormat;
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
     * @var ExcelFormatterAdaptable Format adapter
     */
    protected readonly ExcelFormatterAdaptable $formatAdapter;
    /**
     * @var TargetWritable Write target
     */
    protected readonly TargetWritable $target;
    /**
     * @var ReleasableCollection To be released after finalization
     */
    protected ReleasableCollection $releasedAfterFinalize;


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

        $this->formatAdapter = $formatAdapter;
        $this->target = $target;
        $this->releasedAfterFinalize = new ReleasableCollection();
    }


    /**
     * @inheritDoc
     */
    public function addReleasable(Releasable $resource) : void
    {
        $this->releasedAfterFinalize->add($resource);
    }


    }


    /**
     * @inheritDoc
     */
    public function createSheet(?string $sheetName) : ExcelSheetExportServiceable
    {
        $worksheet = $this->workbook->createSheet();
        if ($sheetName !== null) $worksheet->setTitle($sheetName);
        return new DefaultExcelSheetExportService($this, $worksheet);
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
        } finally {
            $this->releasedAfterFinalize->release();
        }
    }


    /**
     * @inheritDoc
     */
    public function adaptColumnDefinition(ColumnDefinition $column) : ExcelColumnDefinition
    {
        if ($column instanceof ExcelColumnDefinition) return $column;

        $format = $column->format;
        if ($format === null) {
            return ExcelColumnDefinition::adapt($column, ExcelCellFormat::GENERAL);
        }

        $excelFormat = $this->formatAdapter->adapt($format);
        $excelFormatString = $excelFormat->getExcelFormatString() ?? ExcelCellFormat::GENERAL;
        return ExcelColumnDefinition::adapt($column, $excelFormatString);
    }
}