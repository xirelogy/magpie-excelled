<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Codecs\Formats\Formatter;
use Magpie\Exceptions\InvalidStateException;
use Magpie\Exceptions\OperationFailedException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\StreamException;
use Magpie\Exceptions\UnexpectedException;
use Magpie\Exceptions\UnsupportedException;
use Magpie\General\Concepts\StreamWriteable;
use MagpieLib\Excelled\Concepts\ExcelFormatterAdaptable;
use MagpieLib\Excelled\Concepts\Services\ExcelCellExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelColumnExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelRowExportServiceable;

/**
 * Default export service (CSV sheet)
 * @internal
 */
class DefaultCsvSheetExportService extends CommonExcelSheetExportService
{
    /**
     * @var StreamWriteable Target writing stream
     */
    protected readonly StreamWriteable $targetStream;
    /**
     * @var ExcelFormatterAdaptable Format adapter
     */
    protected ExcelFormatterAdaptable $formatAdapter;
    /**
     * @var bool If finalized
     */
    protected bool $isFinalized = false;
    /**
     * @var int Last row
     */
    protected int $lastRow = -1;
    /**
     * @var InnerCsvRowBuffer|null Current row buffer
     */
    protected ?InnerCsvRowBuffer $currentRow = null;
    /**
     * @var array<int, string> Specific row format strings
     */
    protected array $specRowFormatStrings = [];
    /**
     * @var array<int, string> Specific column format strings
     */
    protected array $specColFormatStrings = [];


    /**
     * Constructor
     * @param ExcelExportServiceable $parentService
     * @param StreamWriteable $targetStream
     * @param ExcelFormatterAdaptable $formatAdapter
     * @param bool $isOutputBom
     * @throws StreamException
     */
    public function __construct(ExcelExportServiceable $parentService, StreamWriteable $targetStream, ExcelFormatterAdaptable $formatAdapter, bool $isOutputBom)
    {
        parent::__construct($parentService);

        $this->targetStream = $targetStream;
        $this->formatAdapter = $formatAdapter;

        if ($isOutputBom) {
            $this->targetStream->write(chr(239) . chr(187) . chr(191));
        }
    }


    /**
     * @inheritDoc
     */
    public function activate() : void
    {
        // NOP
    }


    /**
     * @inheritDoc
     */
    public function accessRow(int $row) : ExcelRowExportServiceable
    {
        $formatReactor = function (string $formatString) use ($row) : void {
            $this->specRowFormatStrings[$row] = $formatString;
        };

        return new DefaultCsvRowExportService($this, $formatReactor);
    }


    /**
     * @inheritDoc
     */
    public function accessColumn(int $col) : ExcelColumnExportServiceable
    {
        $formatReactor = function (string $formatString) use ($col) : void {
            $this->specColFormatStrings[$col] = $formatString;
        };

        return new DefaultCsvColumnExportService($this, $formatReactor);
    }


    /**
     * @inheritDoc
     */
    public function accessCell(int $row, int $col, ?int $row2 = null, ?int $col2 = null) : ExcelCellExportServiceable
    {
        $valueReactor = function (mixed $value, ?Formatter $formatter) use ($row, $col, $row2, $col2) {
            $this->setCellValue($row, $col, $row2, $col2, $value, $formatter);
        };

        return new DefaultCsvCellExportService($this, $valueReactor);
    }


    /**
     * Set cell value
     * @param int $row
     * @param int $col
     * @param int|null $row2
     * @param int|null $col2
     * @param mixed $value
     * @param Formatter|null $formatter
     * @return void
     * @throws SafetyCommonException
     */
    protected function setCellValue(int $row, int $col, ?int $row2, ?int $col2, mixed $value, ?Formatter $formatter) : void
    {
        if ($this->isFinalized) throw new InvalidStateException();

        if ($row2 !== null || $col2 !== null) {
            throw new UnsupportedException(_l('CSV exports do not support cell range'));
        }

        $this->getRowBuffer($row)->setValue($col, $value, $formatter);
    }


    /**
     * Get row buffer
     * @param int $row
     * @return InnerCsvRowBuffer
     * @throws SafetyCommonException
     */
    protected function getRowBuffer(int $row) : InnerCsvRowBuffer
    {
        if ($row < $this->lastRow) {
            throw new UnsupportedException(_l('CSV exports do not support row back-tracking'));
        }

        while ($this->lastRow < $row) {
            $content = $this->currentRow?->finalize($this->specColFormatStrings);
            $this->writeCsvLine($content);

            ++$this->lastRow;
            $this->currentRow = new InnerCsvRowBuffer();
        }

        if ($this->currentRow === null) {
            throw new UnexpectedException();
        }

        return $this->currentRow;
    }


    /**
     * @inheritDoc
     */
    public function freezePane(int $row, int $col) : void
    {
        // NOP
    }


    /**
     * @inheritDoc
     */
    public function finalize() : void
    {
        if ($this->isFinalized) throw new InvalidStateException();
        if ($this->currentRow === null) return;

        $content = $this->currentRow->finalize($this->specColFormatStrings);
        $this->writeCsvLine($content);

        $this->targetStream->close();
    }


    /**
     * Write a CSV line
     * @param string|null $content
     * @return void
     * @throws SafetyCommonException
     */
    protected function writeCsvLine(?string $content) : void
    {
        if ($content === null) return;

        try {
            $this->targetStream->write($content . "\r\n");
        } catch (StreamException $ex) {
            throw new OperationFailedException(previous: $ex);
        }
    }
}