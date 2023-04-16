<?php

namespace MagpieLib\Excelled\Impls;

/**
 * Metadata to import Excel as a table according to defined column
 * @internal
 */
class ImportTableColumnMetadata
{
    /**
     * @var string Column key
     */
    public readonly string $key;
    /**
     * @var bool If required (compulsory)
     */
    public readonly bool $isRequired;
    /**
     * @var array<string> Alternative names
     */
    public readonly array $altNames;


    /**
     * Constructor
     * @param string $key
     * @param bool $isRequired
     * @param array<string> $altNames
     */
    public function __construct(string $key, bool $isRequired, array $altNames)
    {
        $this->key = $key;
        $this->isRequired = $isRequired;
        $this->altNames = $altNames;
    }
}