<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Exceptions\UnsupportedException;
use Magpie\Facades\Http\HttpClient;
use Magpie\Facades\Mime\Mime;
use Magpie\General\Concepts\BinaryDataProvidable;
use Magpie\General\Contents\PrimitiveFileBinaryContent;
use Magpie\General\Str;
use MagpieLib\Excelled\Objects\ExcelImage;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing as PhpOfficeDrawing;

/**
 * An Excel image with underlying 'Drawing'
 * @internal
 */
class ExcelDrawingImage extends ExcelImage
{
    /**
     * @var PhpOfficeDrawing Underlying object
     */
    protected readonly PhpOfficeDrawing $drawing;


    /**
     * Constructor
     * @param PhpOfficeDrawing $drawing
     */
    public function __construct(PhpOfficeDrawing $drawing)
    {
        $this->drawing = $drawing;
    }


    /**
     * @inheritDoc
     */
    public function getName() : string
    {
        return $this->drawing->getName();
    }


    /**
     * @inheritDoc
     */
    public function getMimeType() : ?string
    {
       return Mime::getMimeType($this->drawing->getExtension());
    }


    /**
     * @inheritDoc
     */
    public function export() : BinaryDataProvidable
    {
        return OfficeExcepts::protect(function () {
            $drawingPath = $this->drawing->getPath();
            if (!Str::isNullOrEmpty($drawingPath)) {
                if ($this->drawing->getIsURL()) {
                    return HttpClient::initialize()->get($drawingPath);
                } else {
                    $mimeType = $this->getMimeType();
                    $filename = $this->drawing->getFilename();
                    return new PrimitiveFileBinaryContent($drawingPath, $mimeType, $filename);
                }
            } else {
                throw new UnsupportedException();
            }
        });
    }
}