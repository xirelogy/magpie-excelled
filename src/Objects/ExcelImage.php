<?php

namespace MagpieLib\Excelled\Objects;

use Magpie\Exceptions\PersistenceException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\StreamException;
use Magpie\General\Concepts\BinaryDataProvidable;
use Magpie\General\Packs\PackContext;
use Magpie\Objects\CommonObject;

/**
 * An Excel image
 */
abstract class ExcelImage extends CommonObject
{
    /**
     * Image name
     * @return string
     * @throws SafetyCommonException
     */
    public abstract function getName() : string;


    /**
     * Image MIME type
     * @return string|null
     */
    public abstract function getMimeType() : ?string;


    /**
     * Export the image
     * @return BinaryDataProvidable
     * @throws SafetyCommonException
     * @throws PersistenceException
     * @throws StreamException
     */
    public abstract function export() : BinaryDataProvidable;


    /**
     * @inheritDoc
     */
    protected function onPack(object $ret, PackContext $context) : void
    {
        parent::onPack($ret, $context);

        $ret->name = $this->getName();
        $ret->mimeType = $this->getMimeType();
    }
}
