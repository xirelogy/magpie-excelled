<?php

namespace MagpieLib\Excelled\Strategies;

/**
 * Options for exporter
 */
class CommonExporterOptions
{
    /**
     * Default options
     * @return static
     */
    public static function default() : static
    {
        return new static();
    }
}