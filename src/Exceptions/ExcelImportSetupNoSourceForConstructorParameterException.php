<?php

namespace MagpieLib\Excelled\Exceptions;

use Throwable;

/**
 * No available source for object constructor when performing ExcelImport from attributes
 */
class ExcelImportSetupNoSourceForConstructorParameterException extends ExcelImportSetupFromAttributeException
{
    /**
     * Constructor
     * @param string $paramName
     * @param Throwable|null $previous
     * @param int $code
     */
    public function __construct(string $paramName, ?Throwable $previous = null, int $code = 0)
    {
        $message = _format_l('No source for constructor parameter', 'No source for constructor parameter \'{{0}}\'', $paramName);

        parent::__construct($message, $previous, $code);
    }
}