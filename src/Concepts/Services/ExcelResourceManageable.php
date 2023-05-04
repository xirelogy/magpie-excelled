<?php

namespace MagpieLib\Excelled\Concepts\Services;

use Magpie\General\Concepts\Releasable;

/**
 * May manage resources for services related to Excel
 */
interface ExcelResourceManageable
{
    /**
     * Add a resource to be released after finalization
     * @param Releasable $resource
     * @return void
     */
    public function addReleasable(Releasable $resource) : void;
}