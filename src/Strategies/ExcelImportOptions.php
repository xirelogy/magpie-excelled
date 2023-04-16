<?php

namespace MagpieLib\Excelled\Strategies;

/**
 * Options when importing
 */
class ExcelImportOptions
{
    /**
     * @var string|null Sheet name to be imported from
     */
    public ?string $sheetName = null;
    /**
     * @var bool If import images
     */
    public bool $isImportImages = false;


    /**
     * Specify sheet name
     * @param string $sheetName
     * @return $this
     */
    public function withSheetName(string $sheetName) : static
    {
        $this->sheetName = $sheetName;
        return $this;
    }


    /**
     * Specify if import images
     * @param bool $isImportImages
     * @return $this
     */
    public function withImportImages(bool $isImportImages = true) : static
    {
        $this->isImportImages = $isImportImages;
        return $this;
    }


    /**
     * Default options
     * @return static
     */
    public static function default() : static
    {
        return new static();
    }
}