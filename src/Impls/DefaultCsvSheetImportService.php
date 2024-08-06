<?php

namespace MagpieLib\Excelled\Impls;

use Magpie\Exceptions\InvalidDataException;
use Magpie\Exceptions\OperationFailedException;
use Magpie\Exceptions\SafetyCommonException;
use Magpie\Exceptions\StreamException;
use Magpie\Exceptions\UnexpectedException;
use Magpie\General\Concepts\StreamReadable;
use MagpieLib\Excelled\Concepts\Services\ExcelSheetImportServiceable;
use MagpieLib\Excelled\Strategies\ExcelImportOptions;


/**
 * Default import service (CSV sheet)
 * @internal
 */
class DefaultCsvSheetImportService implements ExcelSheetImportServiceable
{
    /**
     * @var StreamReadable Reading stream
     */
    protected readonly StreamReadable $stream;
    /**
     * @var string Field delimiter (so call the 'comma')
     */
    protected readonly string $delimiter;
    /**
     * @var int Block size per read
     */
    protected readonly int $readBlockSize;
    /**
     * @var string Read buffer
     */
    public string $buffer = '';


    /**
     * Constructor
     * @param StreamReadable $stream
     * @param ExcelImportOptions $options
     */
    public function __construct(StreamReadable $stream, ExcelImportOptions $options)
    {
        $this->stream = $stream;
        $this->delimiter = ',';
        $this->readBlockSize = 128;

        _used($options);
    }


    /**
     * @inheritDoc
     */
    public function getRows(int $startRowIndex, int $startColIndex, ?int $endColIndex = null, int $lockRowCount = 1) : iterable
    {
        $currentRowIndex = -1;
        foreach ($this->readRowColumns() as $row) {
            ++$currentRowIndex;
            if ($currentRowIndex < $startRowIndex) continue;

            yield static::filterRow($row, $startColIndex, $endColIndex);
        }
    }


    /**
     * Read stream as stream of columns
     * @return iterable<array>
     * @throws SafetyCommonException
     */
    protected function readRowColumns() : iterable
    {
        $state = DefaultCsvSheetImportServiceParseState::UTF_HEADER_0;

        $retBom = '';
        $retBuffer = '';
        $retRow = [];

        try {
            foreach ($this->readCharacters() as $c) {
                $isReparse = true;
                while ($isReparse) {
                    // Default do not reparse
                    $isReparse = false;

                    switch ($state) {
                        case DefaultCsvSheetImportServiceParseState::UTF_HEADER_0:
                            // Expecting BOM (1)
                            if (ord($c) === 0xEF) {
                                $retBom .= $c;
                                $state = DefaultCsvSheetImportServiceParseState::UTF_HEADER_1;
                            } else {
                                $retBuffer = '';
                                $state = DefaultCsvSheetImportServiceParseState::CONTENT_NORMAL;
                                $isReparse = true;
                            }
                            break;
                        case DefaultCsvSheetImportServiceParseState::UTF_HEADER_1:
                            // Expecting BOM (2)
                            if (ord($c) === 0xBB) {
                                $retBom .= $c;
                                $state = DefaultCsvSheetImportServiceParseState::UTF_HEADER_2;
                            } else {
                                $retBuffer = '';
                                $state = DefaultCsvSheetImportServiceParseState::CONTENT_NORMAL;
                                $isReparse = true;
                            }
                            break;
                        case DefaultCsvSheetImportServiceParseState::UTF_HEADER_2:
                            // Expecting BOM (3)
                            if (ord($c) === 0xBF) {
                                $state = DefaultCsvSheetImportServiceParseState::INITIAL;
                            } else {
                                $retBuffer = $retBom;
                                $state = DefaultCsvSheetImportServiceParseState::CONTENT_NORMAL;
                                $isReparse = true;
                            }
                            $retBom = '';
                            break;

                        case DefaultCsvSheetImportServiceParseState::INITIAL:
                            // Beginning state
                            switch ($c) {
                                case '"';
                                    // Start of quote
                                    $state = DefaultCsvSheetImportServiceParseState::CONTENT_QUOTED;
                                    break;

                                case $this->delimiter:
                                    // Encounter delimiter, treat as a blank column
                                    $retRow[] = '';
                                    break;

                                case "\r":
                                case "\n":
                                    yield $retRow;
                                    $retRow = [];
                                    $state = static::getLineBreakState($c);
                                    break;

                                default:
                                    $retBuffer = '';
                                    $state = DefaultCsvSheetImportServiceParseState::CONTENT_NORMAL;
                                    $isReparse = true;
                                    break;
                            }
                            break;

                        case DefaultCsvSheetImportServiceParseState::CONTENT_NORMAL:
                            // Normal content, almost everything accepted ad-verbatim
                            switch ($c) {
                                case $this->delimiter:
                                    // Content terminated by delimiter
                                    $retRow[] = $retBuffer;
                                    $retBuffer = '';
                                    $state = DefaultCsvSheetImportServiceParseState::INITIAL;
                                    break;

                                case "\r":
                                case "\n":
                                    // Content terminated by line break
                                    $retRow[] = $retBuffer;
                                    $retBuffer = '';
                                    yield $retRow;
                                    $retRow = [];
                                    $state = static::getLineBreakState($c);
                                    break;

                                default:
                                    $retBuffer .= $c;
                                    break;
                            }
                            break;

                        case DefaultCsvSheetImportServiceParseState::CONTENT_QUOTED:
                            // Quoted content
                            switch ($c) {
                                case '"':
                                    $state = DefaultCsvSheetImportServiceParseState::CONTENT_QUOTED_QUOTE;
                                    break;
                                default:
                                    $retBuffer .= $c;
                                    break;
                            }
                            break;

                        case DefaultCsvSheetImportServiceParseState::CONTENT_QUOTED_QUOTE:
                            // Quote in quoted content, escaped quote or end of quote?
                            switch ($c) {
                                case '"':
                                    // This is an escaped quote
                                    $retBuffer .= '"';
                                    $state = DefaultCsvSheetImportServiceParseState::CONTENT_QUOTED;
                                    break;
                                case $this->delimiter:
                                    // Content closed properly
                                    $retRow[] = $retBuffer;
                                    $retBuffer = '';
                                    $state = DefaultCsvSheetImportServiceParseState::INITIAL;
                                    break;
                                case "\r":
                                case "\n":
                                    // Content closed by EOL
                                    $retRow[] = $retBuffer;
                                    $retBuffer = '';
                                    yield $retRow;
                                    $retRow = [];
                                    $state = static::getLineBreakState($c);
                                    break;
                                default:
                                    throw new InvalidDataException(_format_l('Invalid character after end of quote', 'Invalid character \'{{0}}\' after end of quote', $c));
                            }
                            break;

                        case DefaultCsvSheetImportServiceParseState::LINE_BREAK_R:
                            // Line closing using "\r"
                            switch ($c) {
                                case "\n":
                                    // Fulfilled line break pair
                                    $state = DefaultCsvSheetImportServiceParseState::INITIAL;
                                    break;
                                case "\r":
                                default:
                                    // Consecutive "\r" treated as multiple line
                                    // Otherwise also reset the line
                                    $state = DefaultCsvSheetImportServiceParseState::INITIAL;
                                    $isReparse = true;
                                    break;
                            }
                            break;

                        case DefaultCsvSheetImportServiceParseState::LINE_BREAK_N:
                            // Line closing using "\n"
                            switch ($c) {
                                case "\r":
                                    // Fulfilled line break pair
                                    $state = DefaultCsvSheetImportServiceParseState::INITIAL;
                                    break;
                                case "\n":
                                default:
                                    // Consecutive "\n" treated as multiple line
                                    // Otherwise also reset the line
                                    $state = DefaultCsvSheetImportServiceParseState::INITIAL;
                                    $isReparse = true;
                                    break;
                            }
                            break;

                        /** @noinspection PhpUnusedSwitchBranchInspection */
                        default:
                            // Wrong state
                            throw new UnexpectedException(_l('Invalid parse state'));
                    }
                }
            }

            // Closing state processing
            switch ($state) {
                case DefaultCsvSheetImportServiceParseState::CONTENT_QUOTED:
                    // Unterminated quote content!
                    throw new InvalidDataException('Unexpected end of stream');

                case DefaultCsvSheetImportServiceParseState::CONTENT_QUOTED_QUOTE:
                    // Healthy end quote, always merge the buffer
                    $retRow[] = $retBuffer;
                    $retBuffer = '';
                    break;

                default:
                    break;
            }

            // Process any tail data
            if (strlen($retBuffer) > 0) $retRow[] = $retBuffer;
            if (count($retRow) > 0) yield $retRow;
        } catch (StreamException $ex) {
            throw new OperationFailedException(previous: $ex);
        }
    }


    /**
     * Derive line breaking style according to line break character
     * @param string $c
     * @return DefaultCsvSheetImportServiceParseState
     */
    protected static function getLineBreakState(string $c) : DefaultCsvSheetImportServiceParseState
    {
        if ($c === "\r") return DefaultCsvSheetImportServiceParseState::LINE_BREAK_R;
        return DefaultCsvSheetImportServiceParseState::LINE_BREAK_N;
    }


    /**
     * Read stream of characters
     * @return iterable<string>
     * @throws StreamException
     */
    protected function readCharacters() : iterable
    {
        while ($this->stream->hasData()) {
            $buffer = $this->stream->read($this->readBlockSize);
            $bufferLength = strlen($buffer);
            for ($i = 0; $i < $bufferLength; ++$i) {
                yield substr($buffer, $i, 1);
            }
        }
    }


    /**
     * Filter a row according to start and end column specification
     * @param array $row
     * @param int $startColIndex
     * @param int|null $endColIndex
     * @return array
     */
    protected static function filterRow(array $row, int $startColIndex, ?int $endColIndex) : array
    {
        if ($startColIndex === 0 && $endColIndex === null) return $row; // Shortcut function

        $ret = [];
        $currentIndex = -1;
        foreach ($row as $cell) {
            ++$currentIndex;
            if ($currentIndex < $startColIndex) continue;
            if ($endColIndex !== null && $currentIndex > $endColIndex) break;
            $ret[] = $cell;
        }

        return $ret;
    }
}