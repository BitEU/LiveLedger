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
