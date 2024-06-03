<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\UnsupportedValueException;
use Magpie\General\Concepts\PathTargetReadable;
use Magpie\General\Concepts\TargetReadable;
use Magpie\General\FilePath;
use MagpieLib\Excelled\Concepts\Services\ExcelImportServiceable;

/**
 * Importer instance
 */
abstract class CommonImporter
{
    /**
     * Constructor
     */
    protected function __construct()
    {

    }


    /**
     * Access to the service interface
     * @return ExcelImportServiceable
     * @internal
     */
    public final function _getService() : ExcelImportServiceable
    {
        return $this->onGetService();
    }


    /**
     * Access to the service interface (actual)
     * @return ExcelImportServiceable
     */
    protected abstract function onGetService() : ExcelImportServiceable;


    /**
     * Create an instance
     * @param TargetReadable $target
     * @param CommonImporterOptions|null $options
     * @return static
     * @throws SafetyCommonException
     */
    public static final function create(TargetReadable $target, ?CommonImporterOptions $options = null) : static
    {
        $options = $options ?? CommonImporterOptions::default();
        return static::onCreate($target, $options);
    }


    /**
     * Actually create an instance
     * @param TargetReadable $target
     * @param CommonImporterOptions $options
     * @return static
     * @throws SafetyCommonException
     */
    protected static function onCreate(TargetReadable $target, CommonImporterOptions $options) : static
    {
        if (!$target instanceof PathTargetReadable) throw new UnsupportedValueException($target);

        $targetPath = $target->getPath();
        $extension = FilePath::getExtension($targetPath);

        // Optimize CSV handling using very efficient importer
        if (strtolower($extension) === 'csv') {
            return CsvImporter::onCreate($target, $options);
        }

        return ExcelImporter::onCreate($target, $options);
    }
}