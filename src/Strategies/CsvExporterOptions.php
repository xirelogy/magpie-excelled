<?php

namespace MagpieLib\Excelled\Strategies;

/**
 * Options for exporter (CSV)
 */
class CsvExporterOptions extends CommonImporterOptions
{
    /**
     * @var bool If output UTF-8 BOM
     */
    public bool $isOutputUtf8Bom = false;


    /**
     * With output UTF-8 BOM
     * @param bool $isOutput
     * @return $this
     */
    public function withOutputUtf8Bom(bool $isOutput = true) : static
    {
        $this->isOutputUtf8Bom = $isOutput;
        return $this;
    }
}