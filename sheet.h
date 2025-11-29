// sheet.h - Spreadsheet data structures and operations
#ifndef SHEET_H
#define SHEET_H

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <ctype.h>
#include <time.h>
#include "constants.h"

// Cell types
typedef enum {
    CELL_EMPTY,
    CELL_NUMBER,
    CELL_STRING,
    CELL_FORMULA,
    CELL_ERROR
} CellType;

// Error types
typedef enum {
    ERROR_NONE,
    ERROR_DIV_ZERO,
    ERROR_REF,
    ERROR_VALUE,
    ERROR_PARSE,
    ERROR_NA
} ErrorType;

// Data formatting types
typedef enum {
    FORMAT_GENERAL,
    FORMAT_NUMBER,
    FORMAT_PERCENTAGE,
    FORMAT_CURRENCY,
    FORMAT_DATE,
    FORMAT_TIME,
    FORMAT_DATETIME
} DataFormat;

// Date/Time formatting styles
typedef enum {
    DATE_STYLE_MM_DD_YYYY,    // 12/25/2023
    DATE_STYLE_DD_MM_YYYY,    // 25/12/2023  
    DATE_STYLE_YYYY_MM_DD,    // 2023-12-25
    DATE_STYLE_MON_DD_YYYY,   // Dec 25, 2023
    DATE_STYLE_DD_MON_YYYY,   // 25 Dec 2023
    DATE_STYLE_YYYY_MON_DD,   // 2023 Dec 25
    DATE_STYLE_SHORT_DATE,    // 12/25/23
    TIME_STYLE_12HR,          // 2:30 PM
    TIME_STYLE_24HR,          // 14:30
    TIME_STYLE_SECONDS,       // 14:30:45
    TIME_STYLE_12HR_SECONDS,  // 2:30:45 PM
    DATETIME_STYLE_SHORT,     // 12/25/23 2:30 PM
    DATETIME_STYLE_LONG,      // Dec 25, 2023 2:30:45 PM
    DATETIME_STYLE_ISO        // 2023-12-25T14:30:45
} FormatStyle;

// Cell structure
typedef struct Cell {
    CellType type;
    union {
        double number;
        char* string;
        struct {
            char* expression;
            double cached_value;
            char* cached_string;    // For IF function string results
            int is_string_result;   // Flag indicating if result is a string
            ErrorType error;
        } formula;
    } data;
    
    // Display properties
    int width;
    int precision;
    int align;  // 0=left, 1=center, 2=right
    
    // Formatting properties
    DataFormat format;
    FormatStyle format_style;
    
    // Color formatting properties
    int text_color;         // Foreground color (0-15 or -1 for default)
    int background_color;   // Background color (0-15 or -1 for default)
    
    // Size properties
    int row_height;         // Custom row height (-1 for default)
    
    // Dependencies
    struct Cell** depends_on;    // Cells this cell depends on
    int depends_count;
    struct Cell** dependents;    // Cells that depend on this cell
    int dependents_count;
    
    // Position (for dependency tracking)
    int row;
    int col;
} Cell;

// Dependency graph structure for optimized recalculation
typedef struct {
    int* in_degree;          // Number of dependencies for each cell
    int** dependents;        // List of cells that depend on each cell
    int* dependent_count;    // Count of dependents for each cell
} DependencyGraph;

// Range selection structure
typedef struct {
    int start_row, start_col;
    int end_row, end_col;
    int is_active;
} RangeSelection;

// Range clipboard structure
typedef struct {
    Cell*** cells;  // 2D array of copied cells
    int rows, cols;
    int is_active;
} RangeClipboard;

// Sheet structure
typedef struct Sheet {
    Cell*** cells;      // 2D array of cell pointers
    int rows;
    int cols;
    int* col_widths;
    int* row_heights;   // Array of row heights
    char* name;
    
    // Calculation state
    int needs_recalc;
    Cell** calc_order;  // Topological sort of cells for calculation
    int calc_count;
    DependencyGraph dep_graph;  // Dependency tracking
    
    // Range operations
    RangeSelection selection;
    RangeClipboard range_clipboard;
} Sheet;

// Function prototypes
Sheet* sheet_new(int rows, int cols);
void sheet_free(Sheet* sheet);
Cell* sheet_get_cell(Sheet* sheet, int row, int col);
Cell* sheet_get_or_create_cell(Sheet* sheet, int row, int col);
void sheet_set_number(Sheet* sheet, int row, int col, double value);
void sheet_set_string(Sheet* sheet, int row, int col, const char* str);
void sheet_set_formula(Sheet* sheet, int row, int col, const char* formula);
void sheet_clear_cell(Sheet* sheet, int row, int col);
char* sheet_get_display_value(Sheet* sheet, int row, int col);
void sheet_recalculate(Sheet* sheet);
void sheet_recalculate_smart(Sheet* sheet);

// Copy/paste operations
void sheet_copy_cell(Sheet* sheet, int src_row, int src_col, int dest_row, int dest_col);
Cell* sheet_get_clipboard_cell(void);
void sheet_set_clipboard_cell(Cell* cell);

// CSV operations
int sheet_save_csv(Sheet* sheet, const char* filename, int preserve_formulas);
int sheet_load_csv(Sheet* sheet, const char* filename, int preserve_formulas);

// Cell operations
Cell* cell_new(int row, int col);
void cell_free(Cell* cell);
void cell_set_number(Cell* cell, double value);
void cell_set_string(Cell* cell, const char* str);
void cell_set_formula(Cell* cell, const char* formula);
void cell_clear(Cell* cell);
char* cell_get_display_value(Cell* cell);

// Formula evaluation
double evaluate_formula(Sheet* sheet, const char* formula, ErrorType* error);
double evaluate_expression(Sheet* sheet, const char* expr, ErrorType* error);
double evaluate_comparison(Sheet* sheet, const char* expr, ErrorType* error);
int parse_cell_reference(const char* ref, int* row, int* col);
char* cell_reference_to_string(int row, int col);

// Skip whitespace in expression
void skip_whitespace(const char** expr);

// Range copy/paste operations
void sheet_start_range_selection(Sheet* sheet, int row, int col);
void sheet_extend_range_selection(Sheet* sheet, int row, int col);
void sheet_copy_range(Sheet* sheet);
void sheet_paste_range(Sheet* sheet, int start_row, int start_col);
void sheet_clear_range_selection(Sheet* sheet);
int sheet_is_in_selection(Sheet* sheet, int row, int col);

// Cell formatting functions
void cell_set_format(Cell* cell, DataFormat format, FormatStyle style);
char* format_cell_value(Cell* cell);
char* format_number_as_percentage(double value, int precision);
char* format_number_as_currency(double value);
char* format_number_as_date(double value, FormatStyle style);
char* format_number_as_time(double value, FormatStyle style);
char* format_number_as_datetime(double value, FormatStyle date_style, FormatStyle time_style);
char* format_number_as_enhanced_datetime(double value, FormatStyle style);

// Cell color formatting functions
void cell_set_text_color(Cell* cell, int color);
void cell_set_background_color(Cell* cell, int color);
int parse_color(const char* color_str);

// Column/Row resizing functions
void sheet_set_column_width(Sheet* sheet, int col, int width);
void sheet_set_row_height(Sheet* sheet, int row, int height);
int sheet_get_column_width(Sheet* sheet, int col);
int sheet_get_row_height(Sheet* sheet, int row);
void sheet_resize_columns_in_range(Sheet* sheet, int start_col, int end_col, int delta);
void sheet_resize_rows_in_range(Sheet* sheet, int start_row, int end_row, int delta);

// Insert/Delete Row and Column functions
void sheet_insert_row(Sheet* sheet, int row);
void sheet_insert_column(Sheet* sheet, int col);
void sheet_delete_row(Sheet* sheet, int row);
void sheet_delete_column(Sheet* sheet, int col);

// Demo function
void demo_spreadsheet(void);

#endif // SHEET_H
