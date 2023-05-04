<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\UnsupportedValueException;
use Magpie\General\Concepts\PathTargetReadable;
use Magpie\General\Concepts\TargetReadable;
use Magpie\General\Contexts\ScopedCollection;
use Magpie\General\Traits\StaticClass;
use PhpOffice\PhpSpreadsheet\IOFactory as PhpOfficeIOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PhpOfficeSpreadsheet;

/**
 * Excel input/output helper
 * @internal
 */
class ExcelIO
{
    use StaticClass;


    /**
     * Read workbook from target
     * @param TargetReadable $target
     * @return PhpOfficeSpreadsheet
     * @throws SafetyCommonException
     */
    public static function readWorkbookFromTarget(TargetReadable $target) : PhpOfficeSpreadsheet
    {
        if (!$target instanceof PathTargetReadable) throw new UnsupportedValueException($target);

        $scoped = new ScopedCollection($target->getScopes());
        _used($scoped);

        $path = $target->getPath();

        return OfficeExcepts::protect(function () use ($path) {
            $type = PhpOfficeIOFactory::identify($path);
            $reader = PhpOfficeIOFactory::createReader($type);

            return $reader->load($path);
        });
    }
}