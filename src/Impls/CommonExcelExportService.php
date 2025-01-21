<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Exceptions\PersistenceException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\StreamException;
use Magpie\General\Concepts\Releasable;
use Magpie\General\Concepts\TargetWritable;
use Magpie\Objects\ReleasableCollection;
use MagpieLib\Excelled\Concepts\ExcelFormatterAdaptable;
use MagpieLib\Excelled\Concepts\Services\ExcelExportServiceable;
use MagpieLib\Excelled\Constants\ExcelCellFormat;
use MagpieLib\Excelled\Objects\ColumnDefinition;
use MagpieLib\Excelled\Objects\ExcelColumnDefinition;

/**
 * Common implementation of ExcelExportServiceable
 * @internal
 */
abstract class CommonExcelExportService implements ExcelExportServiceable
{
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
     */
    protected function __construct(ExcelFormatterAdaptable $formatAdapter, TargetWritable $target)
    {
        $this->formatAdapter = $formatAdapter;
        $this->target = $target;
        $this->releasedAfterFinalize = new ReleasableCollection();
    }


    /**
     * @inheritDoc
     */
    public final function addReleasable(Releasable $resource) : void
    {
        $this->releasedAfterFinalize->add($resource);
    }


    /**
     * @inheritDoc
     */
    public final function finalize(?string &$mimeType = null) : void
    {
        try {
            $this->onFinalize($mimeType);
        } finally {
            $this->releasedAfterFinalize->release();
        }
    }


    /**
     * Actually finalize the output
     * @param string|null $mimeType
     * @return void
     * @throws SafetyCommonException
     * @throws PersistenceException
     * @throws StreamException
     */
    protected abstract function onFinalize(?string &$mimeType = null) : void;


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