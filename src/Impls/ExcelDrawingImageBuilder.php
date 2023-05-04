<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Exceptions\SafetyCommonException;
use Magpie\General\Concepts\BinaryDataProvidable;
use Magpie\General\Concepts\Releasable;
use Magpie\General\Contents\TemporaryBinaryContent;
use MagpieLib\Excelled\Objects\ExcelImageBuilder;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing as PhpOfficeDrawing;

/**
 * An Excel image builder with underlying 'Drawing'
 * @internal
 */
class ExcelDrawingImageBuilder extends ExcelImageBuilder
{
    /**
     * @var PhpOfficeDrawing Underlying object
     */
    protected readonly PhpOfficeDrawing $drawing;


    /**
     * Constructor
     * @param PhpOfficeDrawing $drawing
     * @param array<Releasable> $resources
     */
    protected function __construct(PhpOfficeDrawing $drawing, array $resources = [])
    {
        parent::__construct($resources);

        $this->drawing = $drawing;
    }


    /**
     * @inheritDoc
     */
    protected function onBuild() : PhpOfficeDrawing
    {
        return $this->drawing;
    }


    /**
     * Create a new instance from binary
     * @param BinaryDataProvidable $data
     * @return static
     * @throws SafetyCommonException
     */
    public static function createFromBinary(BinaryDataProvidable $data) : static
    {
        return OfficeExcepts::protect(function () use ($data) {
            $tempContent = TemporaryBinaryContent::fromContent($data);
            $drawing = new PhpOfficeDrawing();
            $drawing->setPath($tempContent->getFileSystemPath());
            return new static($drawing, [$tempContent]);
        });
    }
}