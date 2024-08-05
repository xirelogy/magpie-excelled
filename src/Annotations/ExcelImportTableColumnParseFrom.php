<?php

namespace MagpieLib\Excelled\Annotations;

use Attribute;
use Magpie\Codecs\Concepts\ObjectParseable;
use Magpie\Codecs\Parsers\CreatableParser;

/**
 * A required column during parsing of Excel import
 */
#[Attribute(Attribute::TARGET_PROPERTY)]
class ExcelImportTableColumnParseFrom
{
    /**
     * @var class-string<CreatableParser>|class-string<ObjectParseable>|null Specification for parser instance to be used
     */
    public ?string $parser;
    /**
     * @var bool If value required from current field
     */
    public bool $isRequired;


    /**
     * Constructor
     * @param class-string<CreatableParser>|class-string<ObjectParseable>|null $parser Specification for parser instance to be used
     * @param bool $isRequired If value required from current field
     */
    public function __construct(
        string|null $parser,
        bool $isRequired,
    ) {
        $this->parser = $parser;
        $this->isRequired = $isRequired;
    }
}