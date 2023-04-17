<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Exceptions\PersistenceException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\StreamException;
use MagpieLib\Excelled\Concepts\Services\ExcelExportServiceable;

/**
 * Definition of an Excel export
 */
abstract class ExcelExportDefinition
{
    /**
     * Run the export
     * @param ExcelExportServiceable $service
     * @param string|null $mimeType
     * @return void
     * @throws SafetyCommonException
     * @throws PersistenceException
     * @throws StreamException
     * @internal
     */
    public final function _run(ExcelExportServiceable $service, ?string &$mimeType = null) : void
    {
        $this->onRun($service, $mimeType);
    }


    /**
     * Run the export
     * @param ExcelExportServiceable $service
     * @param string|null $mimeType
     * @return void
     * @throws SafetyCommonException
     * @throws PersistenceException
     * @throws StreamException
     */
    protected abstract function onRun(ExcelExportServiceable $service, ?string &$mimeType = null) : void;
}