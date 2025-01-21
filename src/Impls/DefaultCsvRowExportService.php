<?php

namespace MagpieLib\Excelled\Impls;

use Closure;
use Magpie\General\Concepts\Releasable;
use MagpieLib\Excelled\Concepts\Services\ExcelRowExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetExportServiceable;
use MagpieLib\Excelled\Objects\Styles\ExcelStyle;

/**
 * Default export service (CSV row)
 * @internal
 */
class DefaultCsvRowExportService implements ExcelRowExportServiceable
{
    /**
     * @var ExcelSheetExportServiceable Parent service
     */
    protected ExcelSheetExportServiceable $parentService;
    /**
     * @var Closure(string):void Format reactor
     */
    protected readonly Closure $formatReactor;


    /**
     * Constructor
     * @param ExcelSheetExportServiceable $parentService
     * @param callable(string):void $formatReactor
     */
    public function __construct(ExcelSheetExportServiceable $parentService, callable $formatReactor)
    {
        $this->parentService = $parentService;
        $this->formatReactor = $formatReactor;
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
        ($this->formatReactor)($formatString);
    }


    /**
     * @inheritDoc
     */
    public final function applyStyle(ExcelStyle ...$styles) : void
    {
        // FIXME
    }
}