<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\General\Concepts\Releasable;
use MagpieLib\Excelled\Concepts\Services\ExcelExportServiceable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetExportServiceable;

/**
 * Common implementation of ExcelSheetExportServiceable
 * @internal
 */
abstract class CommonExcelSheetExportService implements ExcelSheetExportServiceable
{
    /**
     * @var ExcelExportServiceable Parent service
     */
    protected readonly ExcelExportServiceable $parentService;


    /**
     * Constructor
     * @param ExcelExportServiceable $parentService
     */
    protected function __construct(ExcelExportServiceable $parentService)
    {
        $this->parentService = $parentService;
    }


    /**
     * @inheritDoc
     */
    public function addReleasable(Releasable $resource) : void
    {
        $this->parentService->addReleasable($resource);
    }
}