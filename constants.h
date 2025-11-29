// constants.h - Named constants to replace magic numbers throughout the codebase
#ifndef CONSTANTS_H
#define CONSTANTS_H

// Sheet dimensions and limits
#define DEFAULT_COLUMN_WIDTH        10
#define MIN_COLUMN_WIDTH            1
#define MAX_COLUMN_WIDTH            50
#define DEFAULT_ROW_HEIGHT          1
#define MIN_ROW_HEIGHT              1
#define MAX_ROW_HEIGHT              10

// Formula evaluation
#define MAX_RANGE_VALUES            1000
#define FLOAT_COMPARISON_EPSILON    1e-10

// Chart dimensions
#define DEFAULT_CHART_WIDTH         100
#define DEFAULT_CHART_HEIGHT        40
#define MIN_CHART_WIDTH             40
#define MAX_CHART_WIDTH             300
#define MIN_CHART_HEIGHT            15
#define MAX_CHART_HEIGHT            100
#define CHART_AXIS_LABEL_SPACE      12
#define CHART_LEGEND_WIDTH          25

// String buffer sizes
#define MAX_CELL_FORMULA_LENGTH     1024
#define MAX_CELL_DISPLAY_LENGTH     256
#define MAX_ERROR_MESSAGE_LENGTH    128
#define MAX_CSV_LINE_LENGTH         4096
#define MAX_CELL_REF_LENGTH         32
#define MAX_FUNCTION_NAME_LENGTH    32

// Undo/Redo
#define MAX_UNDO_ACTIONS            100

// Autosave
#define AUTOSAVE_INTERVAL_MS        180000  // 3 minutes in milliseconds

// UI and rendering
#define CURSOR_BLINK_RATE_MS        500     // Cursor blink rate in milliseconds
#define FRAME_SLEEP_MS              16      // Sleep time per frame (~60 FPS)
#define CHART_POPUP_SLEEP_MS        50      // Sleep time while chart popup is open
#define DEFAULT_SLEEP_MS            10      // Default sleep time

// Navigation
#define PAGE_JUMP_ROWS              10      // Number of rows to jump on Page Up/Down
#define VISIBLE_COLS_DIVISOR        10      // Divisor for calculating visible columns

// Cell reference sizes
#define CELL_REF_BUFFER_SIZE        16      // Size for cell reference strings like "A1"
#define COLUMN_NAME_BUFFER_SIZE     10      // Size for column name strings
#define ROW_NUMBER_BUFFER_SIZE      10      // Size for row number strings

// Buffer sizes for various operations
#define STATUS_BUFFER_SIZE          256     // Status message buffer size
#define INPUT_BUFFER_SIZE           256     // Input buffer size
#define TITLE_BUFFER_SIZE           128     // Title buffer size
#define MESSAGE_BUFFER_SIZE         256     // General message buffer size
#define DESCRIPTION_BUFFER_SIZE     128     // Description buffer size

// Date/Time constants
#define SECONDS_PER_DAY             86400   // Seconds in a day
#define SECONDS_PER_HOUR            3600    // Seconds in an hour
#define SECONDS_PER_MINUTE          60      // Seconds in a minute
#define HOURS_12                    12      // 12-hour clock constant
#define HOURS_24                    24      // 24-hour clock constant
#define EXCEL_BASE_YEAR             1900    // Excel date serial base year
#define EXCEL_BASE_TIME             -2209161600LL  // 1900-01-01 00:00:00 UTC (adjusted for Excel bug)
#define TM_YEAR_BASE                1900    // struct tm year base (tm_year + 1900 = actual year)
#define CENTURY_MODULO              100     // For getting 2-digit year

// Chart constants
#define CHART_LEGEND_X              12      // X position for chart legend
#define CHART_BAR_X_BASE            12      // Base X position for bar charts
#define MAX_BAR_WIDTH               12      // Maximum bar width in bar charts

// Color formatting (already defined in console.h but repeated here for clarity)
#define NUM_CONSOLE_COLORS          16

// Result codes for standardized error handling
typedef enum {
    LL_OK = 0,
    LL_ERR_MEMORY,
    LL_ERR_INVALID_PARAM,
    LL_ERR_FILE_NOT_FOUND,
    LL_ERR_FILE_IO,
    LL_ERR_PARSE,
    LL_ERR_OVERFLOW,
    LL_ERR_CIRCULAR_REF
} LLResult;

#endif // CONSTANTS_H
