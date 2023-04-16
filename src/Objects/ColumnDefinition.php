<?php

namespace MagpieLib\Excelled\Objects;

use Magpie\Codecs\Formats\Formatter;
use Magpie\General\Packs\PackContext;
use Magpie\Objects\CommonObject;
use MagpieLib\Excelled\Objects\Shims\ExcelColumnAutoSize;
use MagpieLib\Excelled\Objects\Shims\ExcelColumnDefaultSize;

/**
 * Column definition
 */
class ColumnDefinition extends CommonObject
{
    /**
     * @var string|null Column ID (must be unique)
     */
    public ?string $id = null;
    /**
     * @var string Column name (label)
     */
    public readonly string $name;
    /**
     * @var Formatter|null Associated formatter for values of this column
     */
    public readonly ?Formatter $format;
    /**
     * @var float|ExcelColumnAutoSize|ExcelColumnDefaultSize|null Specific column width
     */
    public readonly float|ExcelColumnAutoSize|ExcelColumnDefaultSize|null $setWidth;


    /**
     * Constructor
     * @param string $name
     * @param Formatter|null $format
     * @param float|ExcelColumnAutoSize|ExcelColumnDefaultSize|null $setWidth
     */
    public function __construct(string $name, ?Formatter $format = null, float|ExcelColumnAutoSize|ExcelColumnDefaultSize|null $setWidth = null)
    {
        $this->name = $name;
        $this->format = $format;
        $this->setWidth = $setWidth;
    }


    /**
     * @inheritDoc
     */
    protected function onPack(object $ret, PackContext $context) : void
    {
        parent::onPack($ret, $context);

        $ret->name = $this->name;
    }
}