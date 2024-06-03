<?php

namespace MagpieLib\Excelled\Impls;

/**
 * Parser state for DefaultCsvSheetImportService
 * @internal
 */
enum DefaultCsvSheetImportServiceParseState : int {
    /**
     * Beginning state
     */
    case INITIAL = 0;
    /**
     * Normal content
     */
    case CONTENT_NORMAL = 1;
    /**
     * Quoted content
     */
    case CONTENT_QUOTED = 2;
    /**
     * Quoted content with quote discovered
     */
    case CONTENT_QUOTED_QUOTE = 3;
    /**
     * Handling line break initiated by leading R
     */
    case LINE_BREAK_R = 6;
    /**
     * Handling line break initiated by leading N
     */
    case LINE_BREAK_N = 7;
}