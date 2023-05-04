<?php

namespace MagpieLib\Excelled\Objects;

/**
 * Comments
 */
class ExcelComment
{
    /**
     * @var string|null Comment author
     */
    public ?string $author;
    /**
     * @var string Content of the comment
     */
    public string $content;


    /**
     * Constructor
     * @param string $content
     */
    protected function __construct(string $content)
    {
        $this->author = null;
        $this->content = $content;
    }


    /**
     * Specify author
     * @param string|null $author
     * @return $this
     */
    public function withAuthor(?string $author) : static
    {
        $this->author = $author;
        return $this;
    }


    /**
     * Create an instance
     * @param string $content
     * @return static
     */
    public static function create(string $content) : static
    {
        return new static($content);
    }
}