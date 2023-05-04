<?php

namespace MagpieLib\Excelled\Objects;

use Magpie\Exceptions\SafetyCommonException;
use Magpie\General\Concepts\BinaryDataProvidable;
use Magpie\General\Concepts\Releasable;
use Magpie\Objects\CommonObject;
use MagpieLib\Excelled\Concepts\Services\ExcelResourceManageable;
use MagpieLib\Excelled\Impls\ExcelDrawingImageBuilder;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing as PhpOfficeDrawing;

/**
 * An Excel image builder (adding image to Excel)
 */
abstract class ExcelImageBuilder extends CommonObject
{
    /**
     * @var array<Releasable> Associated resources
     */
    protected array $resources;


    /**
     * Constructor
     * @param array<Releasable> $resources
     */
    protected function __construct(array $resources = [])
    {
        $this->resources = $resources;
    }


    /**
     * Build into underlying object
     * @param ExcelResourceManageable $service
     * @return PhpOfficeDrawing
     * @throws SafetyCommonException
     * @internal
     */
    public final function _build(ExcelResourceManageable $service) : PhpOfficeDrawing
    {
        $ret = $this->onBuild();

        foreach ($this->resources as $resource) {
            $service->addReleasable($resource);
        }

        return $ret;
    }


    /**
     * Build into underlying object
     * @return PhpOfficeDrawing
     * @throws SafetyCommonException
     */
    protected abstract function onBuild() : PhpOfficeDrawing;


    /**
     * Create a new instance from binary
     * @param BinaryDataProvidable $data
     * @return static
     * @throws SafetyCommonException
     */
    public static function fromBinary(BinaryDataProvidable $data) : static
    {
        return ExcelDrawingImageBuilder::createFromBinary($data);
    }
}