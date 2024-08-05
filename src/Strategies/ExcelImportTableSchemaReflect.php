<?php

namespace MagpieLib\Excelled\Strategies;

use Magpie\Codecs\Concepts\ObjectParseable;
use Magpie\Codecs\ParserHosts\ParserHost;
use Magpie\Codecs\Parsers\CreatableParser;
use Magpie\Codecs\Parsers\Parser;
use Magpie\Exceptions\DuplicatedKeyException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\UnsupportedValueException;
use Magpie\General\Traits\StaticClass;
use MagpieLib\Excelled\Annotations\ExcelImportTableColumn;
use MagpieLib\Excelled\Annotations\ExcelImportTableColumnAlt;
use MagpieLib\Excelled\Annotations\ExcelImportTableColumnParseFrom;
use MagpieLib\Excelled\Exceptions\ExcelImportSetupFromAttributeException;
use MagpieLib\Excelled\Exceptions\ExcelImportSetupNoSourceForConstructorParameterException;
use MagpieLib\Excelled\Impls\ExcelImportTableSchemaColumnFromAttribute;
use ReflectionClass;
use ReflectionException;
use ReflectionProperty;

/**
 * ExcelImportTableSchema created from reflection attributes
 */
class ExcelImportTableSchemaReflect
{
    use StaticClass;


    /**
     * Create instance of ExcelImportTableSchema from object attributes
     * @param class-string<T> $className
     * @return ExcelImportTableSchema<T>
     * @template T of object
     * @throws SafetyCommonException
     * @throws ReflectionException
     */
    public static function create(string $className) : ExcelImportTableSchema
    {
        $reflectClass = new ReflectionClass($className);

        /** @var array<string, ExcelImportTableSchemaColumnFromAttribute> $properties */
        $properties = [];

        // Collect properties import
        foreach ($reflectClass->getProperties() as $reflectionProperty) {
            $importColumn = static::getExcelImportTableColumnAttribute($reflectionProperty);
            if ($importColumn === null) continue;

            $parseColumn = static::getExcelImportTableColumnParseFromAttribute($reflectionProperty);
            $parser = static::getExcelImportTableColumnParser($parseColumn);

            $propertyName = $reflectionProperty->name;

            $properties[$propertyName] = new ExcelImportTableSchemaColumnFromAttribute(
                $propertyName,
                $importColumn->columnName,
                static::listExcelImportTableColumnAltColumnAttributes($reflectionProperty),
                $importColumn->isRequired,
                $parser,
                $parseColumn?->isRequired,
            );
        }

        // Create columns definition
        $columns = ExcelImportTableColumnsSchema::create();
        foreach ($properties as $propertyData) {
            if ($propertyData->isRequired) {
                $columns->requires($propertyData->columnName, $propertyData->altColumnNames);
            } else {
                $columns->optional($propertyData->columnName, $propertyData->altColumnNames);
            }
        }

        // Create hydration function using object's constructor
        if (!$reflectClass->hasMethod('__construct')) {
            throw new ExcelImportSetupFromAttributeException(_l('Missing object constructor'));
        }

        $constructor = $reflectClass->getMethod('__construct');
        if (!$constructor->isPublic()) {
            throw new ExcelImportSetupFromAttributeException(_l('Object constructor must be public'));
        }

        /** @var array<callable(ParserHost):mixed> $ctorValues */
        $ctorValues = [];

        // Analyze constructor parameters
        $currentParamIndex = -1;
        foreach ($constructor->getParameters() as $reflectParam) {
            ++$currentParamIndex;
            $paramName = $reflectParam->name;
            if (!array_key_exists($paramName, $properties)) {
                throw new ExcelImportSetupNoSourceForConstructorParameterException($paramName);
            }

            if ($properties[$paramName]->constructorIndex !== null) {
                throw new DuplicatedKeyException($paramName);
            }

            $paramPropertyData = $properties[$paramName];
            $paramPropertyData->constructorIndex = $currentParamIndex;
            $ctorValues[] = static::createValueAccessor($paramPropertyData);
        }

        /** @var array<string, callable(ParserHost):mixed> $propValues */
        $propValues = [];

        // Analyze loose properties
        foreach ($properties as $propertyName => $propertyData) {
            if ($propertyData->constructorIndex !== null) continue;

            $propValues[$propertyName] = static::createValueAccessor($propertyData);
        }

        // Return instance
        return new class($reflectClass, $columns, $ctorValues, $propValues) extends ExcelImportTableSchema {
            /**
             * Constructor
             * @param ReflectionClass $reflectClass
             * @param ExcelImportTableColumnsSchema $columns
             * @param array<callable(ParserHost):mixed> $ctorValues
             * @param array<string, callable(ParserHost):mixed> $propValues
             */
            public function __construct(
                protected readonly ReflectionClass $reflectClass,
                protected readonly ExcelImportTableColumnsSchema $columns,
                protected readonly array $ctorValues,
                protected readonly array $propValues,
            ) {
                parent::__construct();
            }


            /**
             * @inheritDoc
             */
            protected function onHeaderRow(array $values) : void
            {
                static::mapHeaderRow($this->columns, $values);
            }


            /**
             * @inheritDoc
             */
            public function parseRow(ParserHost $parserHost) : mixed
            {
                // Handle constructor properties
                $results = [];
                foreach ($this->ctorValues as $ctorValue) {
                    $results[] = $ctorValue($parserHost);
                }
                $ret = $this->reflectClass->newInstance(...$results);

                // Handle loose properties
                foreach ($this->propValues as $propertyName => $propValue) {
                    $ret->{$propertyName} = $propValue($parserHost);
                }

                // And return
                return $ret;
            }
        };
    }


    /**
     * Create value accessor for given property
     * @param ExcelImportTableSchemaColumnFromAttribute $propertyData
     * @return callable(ParserHost):mixed
     */
    protected static function createValueAccessor(ExcelImportTableSchemaColumnFromAttribute $propertyData) : callable
    {
        return function (ParserHost $parserHost) use ($propertyData) {
            if ($propertyData->isValueRequired === null) return null;

            if ($propertyData->isValueRequired) {
                return $parserHost->requires($propertyData->columnName, $propertyData->valueParser);
            } else {
                return $parserHost->optional($propertyData->columnName, $propertyData->valueParser);
            }
        };
    }


    /**
     * Get first ExcelImportTableColumn
     * @param ReflectionProperty $reflect
     * @return ExcelImportTableColumn|null
     */
    protected static function getExcelImportTableColumnAttribute(ReflectionProperty $reflect) : ?ExcelImportTableColumn
    {
        foreach ($reflect->getAttributes(ExcelImportTableColumn::class) as $attr) {
            return $attr->newInstance();
        }

        return null;
    }


    /**
     * List all attributes of ExcelImportTableColumnAlt
     * @param ReflectionProperty $reflect
     * @return iterable<string>
     */
    protected static function listExcelImportTableColumnAltColumnAttributes(ReflectionProperty $reflect) : iterable
    {
        foreach ($reflect->getAttributes(ExcelImportTableColumnAlt::class) as $attr) {
            /** @var ExcelImportTableColumnAlt $inst */
            $inst = $attr->newInstance();
            yield $inst->altColumnName;
        }
    }


    /**
     * Get first ExcelImportTableColumnParseFrom
     * @param ReflectionProperty $reflect
     * @return ExcelImportTableColumnParseFrom|null
     */
    protected static function getExcelImportTableColumnParseFromAttribute(ReflectionProperty $reflect) : ?ExcelImportTableColumnParseFrom
    {
        foreach ($reflect->getAttributes(ExcelImportTableColumnParseFrom::class) as $attr) {
            return $attr->newInstance();
        }

        return null;
    }


    /**
     * @param ExcelImportTableColumnParseFrom|null $attr
     * @return Parser|null
     * @throws SafetyCommonException
     */
    protected static function getExcelImportTableColumnParser(?ExcelImportTableColumnParseFrom $attr) : ?Parser
    {
        if ($attr === null) return null;

        $parserSpec = $attr->parser;
        if ($parserSpec === null) return null;

        if (is_subclass_of($parserSpec, CreatableParser::class)) {
            return $parserSpec::create();
        }

        if (is_subclass_of($parserSpec, ObjectParseable::class)) {
            return $parserSpec::createParser();
        }

        throw new UnsupportedValueException($parserSpec, _l('parser specification'));
    }
}