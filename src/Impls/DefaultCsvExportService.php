<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\StreamException;
use Magpie\Exceptions\UnsupportedException;
use Magpie\General\Concepts\TargetReadable;
use Magpie\General\Concepts\TargetWritable;
use MagpieLib\Excelled\Concepts\ExcelFormatterAdaptable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetExportServiceable;
use MagpieLib\Excelled\Strategies\CsvExporterOptions;

/**
 * Default export service (CSV)
 * @internal
 */
class DefaultCsvExportService extends CommonExcelExportService
{
    /**
     * @var DefaultCsvSheetExportService Associated sheet service
     */
    protected DefaultCsvSheetExportService $sheetService;


    /**
     * Constructor
     * @param ExcelFormatterAdaptable $formatAdapter
     * @param TargetWritable $target
     * @param CsvExporterOptions $options
     * @throws SafetyCommonException
     * @throws StreamException
     */
    public function __construct(ExcelFormatterAdaptable $formatAdapter, TargetWritable $target, CsvExporterOptions $options)
    {
        parent::__construct($formatAdapter, $target);

        $targetStream = $target->createStream();
        $this->sheetService = new DefaultCsvSheetExportService($this, $targetStream, $formatAdapter, $options->isOutputUtf8Bom);
    }


    /**
     * @inheritDoc
     */
    public function load(TargetReadable $target) : void
    {
        throw new UnsupportedException(_l('CSV exports do not support loading'));
    }


    /**
     * @inheritDoc
     */
    public function accessSheet(string $sheetName) : ExcelSheetExportServiceable
    {
        return $this->sheetService;
    }


    /**
     * @inheritDoc
     */
    public function createSheet(?string $sheetName) : ExcelSheetExportServiceable
    {
        return $this->sheetService;
    }


    /**
     * @inheritDoc
     */
    protected function onFinalize(?string &$mimeType = null) : void
    {
        $this->sheetService->finalize();
    }
}