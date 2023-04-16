<?php

namespace MagpieLib\Excelled\Objects;

use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\UnexpectedException;
use Magpie\Objects\CommonObject;
use Magpie\Objects\Traits\CommonObjectPackAll;

/**
 * A list of Excel images
 */
class ExcelImages extends CommonObject
{
    use CommonObjectPackAll;

    /**
     * @var array<ExcelImage> All images
     */
    public readonly array $images;


    /**
     * Constructor
     * @param iterable<ExcelImage> $images
     * @throws SafetyCommonException
     */
    public function __construct(iterable $images)
    {
        $this->images = iter_flatten($images, false);
        if (count($this->images) <= 0) throw new UnexpectedException();
    }


    /**
     * First image
     * @return ExcelImage
     * @throws SafetyCommonException
     */
    public function getFirst() : ExcelImage
    {
        return iter_first($this->images) ?? throw new UnexpectedException();
    }
}
