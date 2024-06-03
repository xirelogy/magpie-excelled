<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\General\Concepts\StreamReadable;
use MagpieLib\Excelled\Concepts\Services\ExcelImportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetImportServiceable;
use MagpieLib\Excelled\Strategies\ExcelImportOptions;

/**
 * Default import service (for CSV)
 * @internal
 */
class DefaultCsvImportService implements ExcelImportServiceable
{
    /**
     * @var StreamReadable Reading stream
     */
    protected readonly StreamReadable $stream;


    /**
     * Constructor
     * @param StreamReadable $stream
     */
    public function __construct(StreamReadable $stream)
    {
        $this->stream = $stream;
    }


    /**
     * @inheritDoc
     */
    public function getSheet(ExcelImportOptions $options) : ExcelSheetImportServiceable
    {
        return new DefaultCsvSheetImportService($this->stream, $options);
    }
}