<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Exceptions\SafetyCommonException;
use Magpie\General\Concepts\StreamReadable;
use Magpie\General\Concepts\TargetReadable;
use MagpieLib\Excelled\Concepts\Services\ExcelImportServiceable;
use MagpieLib\Excelled\Impls\DefaultCsvImportService;

/**
 * CSV importer instance
 */
class CsvImporter extends CommonImporter
{
    /**
     * @var StreamReadable Reading stream
     */
    protected readonly StreamReadable $stream;


    /**
     * Constructor
     * @param TargetReadable $target
     * @param CsvImporterOptions $options
     * @throws SafetyCommonException
     */
    protected function __construct(TargetReadable $target, CsvImporterOptions $options)
    {
        parent::__construct();

        $this->stream = $target->createStream();

        _used($options);
    }


    /**
     * @inheritDoc
     */
    protected function onGetService() : ExcelImportServiceable
    {
        return new DefaultCsvImportService($this->stream);
    }


    protected static function onCreate(TargetReadable $target, CommonImporterOptions $options) : static
    {
        $options = $options instanceof CsvImporterOptions ? $options : CsvImporterOptions::default();
        return new static($target, $options);
    }
}