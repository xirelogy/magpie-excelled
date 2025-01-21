<?php

namespace MagpieLib\Excelled\Impls;

use Closure;
use Magpie\Codecs\Formats\Formatter;
use Magpie\Exceptions\UnsupportedException;
use Magpie\General\Concepts\Releasable;
use MagpieLib\Excelled\Concepts\Services\ExcelCellExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetExportServiceable;
use MagpieLib\Excelled\Objects\ExcelComment;
use MagpieLib\Excelled\Objects\ExcelImage;
use MagpieLib\Excelled\Objects\ExcelImageBuilder;
use MagpieLib\Excelled\Objects\Styles\ExcelStyle;

/**
 * Default export service (CSV cell)
 * @internal
 */
class DefaultCsvCellExportService implements ExcelCellExportServiceable
{
    /**
     * @var ExcelSheetExportServiceable Parent service
     */
    protected readonly ExcelSheetExportServiceable $parentService;
    /**
     * @var Closure(mixed,Formatter|null):void Value reactor function
     */
    protected readonly Closure $valueReactor;
    /**
     * @var int Row
     */
    protected readonly int $row;
    /**
     * @var int Column
     */
    protected readonly int $col;


    /**
     * Constructor
     * @param ExcelSheetExportServiceable $parentService
     * @param callable(mixed,Formatter|null):void $valueReactor
     */
    public function __construct(ExcelSheetExportServiceable $parentService, callable $valueReactor)
    {
        $this->parentService = $parentService;
        $this->valueReactor = $valueReactor;
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
    public final function setFormat(string $formatString) : void
    {
        // FIXME
    }


    /**
     * @inheritDoc
     */
    public final function applyStyle(ExcelStyle ...$styles) : void
    {
        // FIXME
    }


    /**
     * @inheritDoc
     */
    public function setValue(mixed $value, ?Formatter $formatter = null) : void
    {
        ($this->valueReactor)($value, $formatter);
    }


    /**
     * @inheritDoc
     */
    public function addImage(ExcelImageBuilder $builder) : ExcelImage
    {
        throw new UnsupportedException();
    }


    /**
     * @inheritDoc
     */
    public function setComment(ExcelComment $comment) : void
    {
        // NOP
    }
}