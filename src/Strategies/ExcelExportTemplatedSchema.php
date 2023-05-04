<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Exceptions\PersistenceException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\General\Concepts\TargetReadable;
use MagpieLib\Excelled\Concepts\Services\ExcelExportServiceable;

/**
 * A schema to export Excel from template
 */
abstract class ExcelExportTemplatedSchema extends ExcelExportSchema
{
    /**
     * Export using given service
     * @param ExcelExportServiceable $service
     * @return void
     * @throws SafetyCommonException
     * @throws PersistenceException
     */
    public final function exportUsing(ExcelExportServiceable $service) : void
    {
        $service->load($this->getTemplateTarget());

        $this->onExport($service);
    }


    /**
     * Get target for template
     * @return TargetReadable
     * @throws SafetyCommonException
     */
    protected abstract function getTemplateTarget() : TargetReadable;


    /**
     * Handle export
     * @param ExcelExportServiceable $service
     * @return void
     * @throws SafetyCommonException
     * @throws PersistenceException
     */
    protected abstract function onExport(ExcelExportServiceable $service) : void;
}