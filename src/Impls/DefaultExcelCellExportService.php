<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Codecs\Formats\Formatter;
use Magpie\Exceptions\UnsupportedException;
use Magpie\General\Concepts\Releasable;
use Magpie\General\Str;
use MagpieLib\Excelled\Concepts\ExcelFormatterAdaptable;
use MagpieLib\Excelled\Concepts\Services\ExcelCellExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetExportServiceable;
use MagpieLib\Excelled\Objects\ExcelComment;
use MagpieLib\Excelled\Objects\ExcelImage;
use MagpieLib\Excelled\Objects\ExcelImageBuilder;
use MagpieLib\Excelled\Strategies\ExcelNames;
use PhpOffice\PhpSpreadsheet\Style\Style as PhpOfficeStyle;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet as PhpOfficeWorksheet;

/**
 * Default export service (cell)
 * @internal
 */
class DefaultExcelCellExportService extends DefaultExcelGeneralExportService implements ExcelCellExportServiceable
{
    /**
     * @var ExcelSheetExportServiceable Parent service
     */
    protected ExcelSheetExportServiceable $parentService;
    /**
     * @var PhpOfficeWorksheet Associated worksheet
     */
    protected readonly PhpOfficeWorksheet $worksheet;
    /**
     * @var ExcelFormatterAdaptable Format adapter
     */
    protected readonly ExcelFormatterAdaptable $formatAdapter;
    /**
     * @var string Associated cell name
     */
    protected readonly string $cellName;
    /**
     * @var bool If specification is a range
     */
    protected readonly bool $isRange;


    /**
     * Constructor
     * @param ExcelSheetExportServiceable $parentService
     * @param PhpOfficeWorksheet $worksheet
     * @param ExcelFormatterAdaptable $formatAdapter
     * @param int $row
     * @param int $col
     * @param int|null $row2
     * @param int|null $col2
     */
    public function __construct(ExcelSheetExportServiceable $parentService, PhpOfficeWorksheet $worksheet, ExcelFormatterAdaptable $formatAdapter, int $row, int $col, ?int $row2, ?int $col2)
    {
        $this->parentService = $parentService;
        $this->worksheet = $worksheet;
        $this->formatAdapter = $formatAdapter;
        $this->cellName = static::formatCellName($row, $col, $row2, $col2);
        $this->isRange = $row2 !== null || $col2 !== null;
    }


    /**
     * @inheritDoc
     */
    public function addReleasable(Releasable $resource) : void
    {
        $this->parentService->addReleasable($resource);
    }


    /**
     * @inheritDoc
     */
    public function setValue(mixed $value, ?Formatter $formatter = null) : void
    {
        if ($this->isRange) throw new UnsupportedException();

        if ($formatter !== null) {
            $excelFormatter = $this->formatAdapter->adapt($formatter);
            $excelFormatString = $excelFormatter->getExcelFormatString();
            if ($excelFormatString !== null) $this->setFormat($excelFormatString);

            $excelDataType = ExcelDataType::getDataType($excelFormatString);
            $formattedValue = $formatter->format($value);

            if ($excelDataType !== null) {
                $this->worksheet->setCellValueExplicit($this->cellName, $formattedValue, $excelDataType);
            } else {
                $this->worksheet->setCellValue($this->cellName, $formattedValue);
            }
        } else {
            $this->worksheet->setCellValue($this->cellName, $value);
        }
    }


    /**
     * @inheritDoc
     */
    public function addImage(ExcelImageBuilder $builder) : ExcelImage
    {
        if ($this->isRange) throw new UnsupportedException();

        return OfficeExcepts::protect(function () use ($builder) {
            $excelDrawing = $builder->_build($this->parentService);
            $excelDrawing->setCoordinates($this->cellName);
            $excelDrawing->setWorksheet($this->worksheet);
            $excelDrawing->setOffsetX2($excelDrawing->getImageWidth());
            $excelDrawing->setOffsetY2($excelDrawing->getImageHeight());
            return new ExcelDrawingImage($excelDrawing);
        });
    }


    /**
     * @inheritDoc
     */
    public function setComment(ExcelComment $comment) : void
    {
        OfficeExcepts::protect(function () use ($comment) {
            $excelComment = $this->worksheet->getComment($this->cellName);

            if (!Str::isNullOrEmpty($comment->author)) {
                $excelComment->setAuthor($comment->author);
            }

            $excelComment->getText()->createTextRun($comment->content);
        });
    }


    /**
     * @inheritDoc
     */
    protected function getStyle() : PhpOfficeStyle
    {
        return $this->worksheet->getStyle($this->cellName);
    }


    /**
     * Format cell name
     * @param int $row
     * @param int $col
     * @param int|null $row2
     * @param int|null $col2
     * @return string
     */
    protected static function formatCellName(int $row, int $col, ?int $row2, ?int $col2) : string
    {
        if ($row2 !== null && $col2 !== null) {
            // Assume cell range
            return ExcelNames::rangeNameOf($row, $col, $row2, $col2);
        }

        if ($col2 !== null) {
            // Assume column range
            return ExcelNames::columnRangeNameOf($col, $col2);
        }

        if ($row2 !== null) {
            // Assume row range
            return ExcelNames::rowRangeNameOf($row, $row2);
        }

        // Assume single cell
        return ExcelNames::cellNameOf($row, $col);
    }
}