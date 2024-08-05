<?php

namespace MagpieLib\Excelled\Exceptions;

use Magpie\Exceptions\SafetyCommonException;
use Throwable;

/**
 * Exception during setup for ExcelImport from attributes
 */
class ExcelImportSetupFromAttributeException extends SafetyCommonException
{
    /**
     * Constructor
     * @param string|null $message
     * @param Throwable|null $previous
     * @param int $code
     */
    public function __construct(?string $message, ?Throwable $previous = null, int $code = 0)
    {
        $message = $message ?? _l('Cannot properly setup Excel import from attributes');

        parent::__construct($message, $previous, $code);
    }
}