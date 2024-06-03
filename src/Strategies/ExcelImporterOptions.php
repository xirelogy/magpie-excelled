<?php

namespace MagpieLib\Excelled\Strategies;

/**
 * Options for importer
 */
class ExcelImporterOptions
{
    /**
     * @var bool If save memory
     */
    public bool $isSaveMemory = false;


    /**
     * Specify if save memory
     * @param bool $isSave
     * @return $this
     */
    public function withSaveMemory(bool $isSave = true) : static
    {
        $this->isSaveMemory = $isSave;
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