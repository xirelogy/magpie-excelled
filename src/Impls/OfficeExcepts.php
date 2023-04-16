<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Exceptions\OperationFailedException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\General\Traits\StaticClass;
use Throwable;

/**
 * Handle exceptions gracefully
 * @internal
 */
class OfficeExcepts
{
    use StaticClass;


    /**
     * Protect and converge exception
     * @param callable():mixed $fn
     * @return mixed
     * @throws SafetyCommonException
     */
    public static function protect(callable $fn) : mixed
    {
        try {
            return $fn();
        } catch (SafetyCommonException $ex) {
            throw $ex;
        } catch (Throwable $ex) {
            throw new OperationFailedException(previous: $ex);
        }
    }
}
