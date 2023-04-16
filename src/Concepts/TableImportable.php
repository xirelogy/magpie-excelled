<?php

namespace MagpieLib\Excelled\Concepts;

use Magpie\Codecs\ParserHosts\ParserHost;
use Magpie\Exceptions\SafetyCommonException;

/**
 * Table-like import schema
 */
interface TableImportable
{
    /**
     * Parse data
     * @param ParserHost $parserHost
     * @return mixed
     * @throws SafetyCommonException
     */
    public function parseRow(ParserHost $parserHost) : mixed;
}