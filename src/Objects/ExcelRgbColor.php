<?php

namespace MagpieLib\Excelled\Objects;

use Magpie\Objects\Traits\CommonObjectPackAll;

/**
 * Excel color specification in RGB
 */
class ExcelRgbColor extends ExcelColor
{
    use CommonObjectPackAll;


    /**
     * @var int Red components
     */
    public int $r;
    /**
     * @var int Green components
     */
    public int $g;
    /**
     * @var int Blue components
     */
    public int $b;


    /**
     * Constructor
     * @param int $r
     * @param int $g
     * @param int $b
     */
    public function __construct(int $r, int $g, int $b)
    {
        $this->r = $r;
        $this->g = $g;
        $this->b = $b;
    }


    /**
     * Corresponding color string
     * @return string
     */
    public function getColorString() : string
    {
        return static::hex(255) . static::hex($this->r) . static::hex($this->g) . static::hex($this->b);
    }
}