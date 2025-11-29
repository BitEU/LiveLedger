// sheet.c - Spreadsheet implementation
#include "sheet.h"
#include "console.h"

// Global variables for IF function string results
char g_if_result_string[256] = {0};
int g_if_result_is_string = 0;
static Cell* g_current_evaluating_cell = NULL;

// Forward declarations for expression parsing
double parse_arithmetic_expression(Sheet* sheet, const char** expr, ErrorType* error);
double parse_term(Sheet* sheet, const char** expr, ErrorType* error);
double parse_factor(Sheet* sheet, const char** expr, ErrorType* error);
double parse_function(Sheet* sheet, const char** expr, ErrorType* error);

// Range parsing structures
typedef struct {
    int start_row, start_col;
    int end_row, end_col;
} CellRange;

int parse_range(const char* range_str, CellRange* range);
int get_range_values(Sheet* sheet, const CellRange* range, double* values, int max_values);

// Function implementations
double func_sum(const double* values, int count);
double func_avg(const double* values, int count);
double func_max(const double* values, int count);
double func_min(const double* values, int count);
double func_median(double* values, int count);
double func_mode(const double* values, int count);
double func_if(double condition, double true_val, double false_val);
double func_if_enhanced(double condition, double true_val, double false_val, 
                       const char* true_str, const char* false_str);
double func_power(double base, double exponent);
double func_xlookup(Sheet* sheet, double lookup_value, const char* lookup_str,
                   const char* lookup_array, const char* return_array, int exact_match, ErrorType* error);
int compare_double(const void* a, const void* b);
char* escape_csv_string(const char* str);
char* parse_csv_field(const char** csv_line, int* is_end);
char* get_cell_string_value(Sheet* sheet, const char* ref);

// Implementation

Sheet* sheet_new(int rows, int cols) {
    Sheet* sheet = (Sheet*)calloc(1, sizeof(Sheet));
    if (!sheet) return NULL;
    
    sheet->rows = rows;
    sheet->cols = cols;
    sheet->name = _strdup("Sheet1");
    if (!sheet->name) {
        free(sheet);
        return NULL;
    }
    
    // Allocate cell array
    sheet->cells = (Cell***)calloc(rows, sizeof(Cell**));
    if (!sheet->cells) {
        free(sheet);
        return NULL;
    }
    
    for (int i = 0; i < rows; i++) {
        sheet->cells[i] = (Cell**)calloc(cols, sizeof(Cell*));
        if (!sheet->cells[i]) {
            // Cleanup on failure
            for (int j = 0; j < i; j++) {
                free(sheet->cells[j]);
            }
            free(sheet->cells);
            free(sheet);
            return NULL;
        }
    }    // Initialize column widths
    sheet->col_widths = (int*)calloc(cols, sizeof(int));
    for (int i = 0; i < cols; i++) {
        sheet->col_widths[i] = DEFAULT_COLUMN_WIDTH;
    }
    
    // Initialize row heights
    sheet->row_heights = (int*)calloc(rows, sizeof(int));
    for (int i = 0; i < rows; i++) {
        sheet->row_heights[i] = 1;  // Default height
    }
    
    // Initialize range selection and clipboard
    sheet->selection.is_active = 0;
    sheet->range_clipboard.is_active = 0;
    sheet->range_clipboard.cells = NULL;
    
    return sheet;
}

void sheet_free(Sheet* sheet) {
    if (!sheet) return;
    
    // Free all cells
    for (int i = 0; i < sheet->rows; i++) {
        for (int j = 0; j < sheet->cols; j++) {
            if (sheet->cells[i][j]) {
                cell_free(sheet->cells[i][j]);
            }
        }
        free(sheet->cells[i]);    }
    free(sheet->cells);
    
    // Free range clipboard
    if (sheet->range_clipboard.cells) {
        for (int i = 0; i < sheet->range_clipboard.rows; i++) {
            for (int j = 0; j < sheet->range_clipboard.cols; j++) {
                if (sheet->range_clipboard.cells[i][j]) {
                    cell_free(sheet->range_clipboard.cells[i][j]);
                }
            }
            free(sheet->range_clipboard.cells[i]);
        }
        free(sheet->range_clipboard.cells);
    }
      free(sheet->col_widths);
    free(sheet->row_heights);  // Free row heights
    free(sheet->name);
    free(sheet->calc_order);
    free(sheet);
}

Cell* sheet_get_cell(Sheet* sheet, int row, int col) {
    if (row < 0 || row >= sheet->rows || col < 0 || col >= sheet->cols) {
        return NULL;
    }
    return sheet->cells[row][col];
}

Cell* sheet_get_or_create_cell(Sheet* sheet, int row, int col) {
    if (row < 0 || row >= sheet->rows || col < 0 || col >= sheet->cols) {
        return NULL;
    }
    
    if (!sheet->cells[row][col]) {
        sheet->cells[row][col] = cell_new(row, col);
    }
    
    return sheet->cells[row][col];
}

Cell* cell_new(int row, int col) {
    Cell* cell = (Cell*)calloc(1, sizeof(Cell));
    if (!cell) return NULL;    cell->type = CELL_EMPTY;
    cell->row = row;
    cell->col = col;
    cell->width = 10;
    cell->precision = 2;
    cell->align = 2;  // Right align for numbers by default
    
    // Initialize formatting
    cell->format = FORMAT_GENERAL;
    cell->format_style = 0;
    
    // Initialize color formatting
    cell->text_color = -1;        // Default text color
    cell->background_color = -1;  // Default background color
    cell->row_height = -1;        // Default row height
    
    return cell;
}

void cell_free(Cell* cell) {
    if (!cell) return;
      // Free string data
    if (cell->type == CELL_STRING && cell->data.string) {
        free(cell->data.string);
    } else if (cell->type == CELL_FORMULA) {
        if (cell->data.formula.expression) {
            free(cell->data.formula.expression);
        }
        if (cell->data.formula.cached_string) {
            free(cell->data.formula.cached_string);
        }
    }
    
    // Free dependency arrays
    free(cell->depends_on);
    free(cell->dependents);
    
    free(cell);
}

void cell_clear(Cell* cell) {
    if (!cell) return;
      // Save the type and set to CELL_EMPTY first to prevent double-free on re-entry
    CellType old_type = cell->type;
    cell->type = CELL_EMPTY;
    
    // Free existing data
    if (old_type == CELL_STRING && cell->data.string) {
        free(cell->data.string);
        cell->data.string = NULL;
    } else if (old_type == CELL_FORMULA) {
        if (cell->data.formula.expression) {
            free(cell->data.formula.expression);
            cell->data.formula.expression = NULL;
        }
        if (cell->data.formula.cached_string) {
            free(cell->data.formula.cached_string);
            cell->data.formula.cached_string = NULL;
        }    }

    // Keep formatting when clearing
}

void cell_set_number(Cell* cell, double value) {
    if (!cell) return;
    
    cell_clear(cell);
    cell->type = CELL_NUMBER;
    cell->data.number = value;
}

void cell_set_string(Cell* cell, const char* str) {
    if (!cell) return;
    
    cell_clear(cell);
    cell->type = CELL_STRING;
    cell->data.string = _strdup(str);
    if (!cell->data.string) {
        cell->type = CELL_EMPTY;  // Revert on allocation failure
        return;
    }
    cell->align = 0;  // Left align for strings
}

void cell_set_formula(Cell* cell, const char* formula) {
    if (!cell) return;
      cell_clear(cell);
    cell->type = CELL_FORMULA;
    cell->data.formula.expression = _strdup(formula);
    if (!cell->data.formula.expression) {
        cell->type = CELL_EMPTY;  // Revert on allocation failure
        return;
    }
    cell->data.formula.cached_value = 0.0;
    cell->data.formula.cached_string = NULL;
    cell->data.formula.is_string_result = 0;
    cell->data.formula.error = ERROR_NONE;
}

void sheet_set_number(Sheet* sheet, int row, int col, double value) {
    Cell* cell = sheet_get_or_create_cell(sheet, row, col);
    if (cell) {
        cell_set_number(cell, value);
        sheet->needs_recalc = 1;
    }
}

void sheet_set_string(Sheet* sheet, int row, int col, const char* str) {
    Cell* cell = sheet_get_or_create_cell(sheet, row, col);
    if (cell) {
        cell_set_string(cell, str);
    }
}

void sheet_set_formula(Sheet* sheet, int row, int col, const char* formula) {
    Cell* cell = sheet_get_or_create_cell(sheet, row, col);
    if (cell) {
        cell_set_formula(cell, formula);
        sheet->needs_recalc = 1;
    }
}

void sheet_clear_cell(Sheet* sheet, int row, int col) {
    Cell* cell = sheet_get_cell(sheet, row, col);
    if (cell) {
        cell_clear(cell);
        sheet->needs_recalc = 1;
    }
}

char* cell_get_display_value(Cell* cell) {
    // Use the formatting function
    return format_cell_value(cell);
}

char* sheet_get_display_value(Sheet* sheet, int row, int col) {
    Cell* cell = sheet_get_cell(sheet, row, col);
    return cell_get_display_value(cell);
}

// Convert column number to letter(s) (0 -> A, 25 -> Z, 26 -> AA, etc.)
// Uses rotating static buffers to avoid issues with multiple calls in same expression
char* cell_reference_to_string(int row, int col) {
    static char buffers[4][16];
    static int current = 0;
    char* buffer = buffers[current];
    current = (current + 1) % 4;
    
    char colStr[8];
    int idx = 0;
    
    // Convert column number to letters
    col++;  // Make it 1-based for conversion
    while (col > 0) {
        col--;
        colStr[idx++] = 'A' + (col % 26);
        col /= 26;
    }
    
    // Reverse the string
    for (int i = 0; i < idx / 2; i++) {
        char temp = colStr[i];
        colStr[i] = colStr[idx - 1 - i];
        colStr[idx - 1 - i] = temp;
    }
    colStr[idx] = '\0';
    
    snprintf(buffer, sizeof(buffers[0]), "%s%d", colStr, row + 1);
    return buffer;
}

// Parse cell reference like "A1" or "AB23"
int parse_cell_reference(const char* ref, int* row, int* col) {
    if (!ref || !row || !col) return 0;
    
    const char* p = ref;
    *col = 0;
    
    // Skip whitespace
    while (*p && isspace(*p)) p++;
    
    // Parse column letters
    if (!isalpha(*p)) return 0;
    
    while (*p && isalpha(*p)) {
        *col = *col * 26 + (toupper(*p) - 'A' + 1);
        p++;
    }
    (*col)--;  // Convert to 0-based
    
    // Parse row number
    if (!isdigit(*p)) return 0;
    
    *row = 0;
    while (*p && isdigit(*p)) {
        *row = *row * 10 + (*p - '0');
        p++;
    }
    (*row)--;  // Convert to 0-based
    
    // Check for trailing garbage
    while (*p && isspace(*p)) p++;
    if (*p) return 0;
    
    return 1;
}

// Forward declarations for expression parsing
double parse_arithmetic_expression(Sheet* sheet, const char** expr, ErrorType* error);
double parse_arithmetic_expression_old(Sheet* sheet, const char* expr, ErrorType* error);
double parse_term(Sheet* sheet, const char** expr, ErrorType* error);
double parse_factor(Sheet* sheet, const char** expr, ErrorType* error);
double parse_function(Sheet* sheet, const char** expr, ErrorType* error);
void skip_whitespace(const char** expr);

// Parse range notation like "A1:A3" or "B2:D5"
int parse_range(const char* range_str, CellRange* range) {
    if (!range_str || !range) return 0;
    
    const char* colon = strchr(range_str, ':');
    if (!colon) return 0;
    
    // Parse start cell
    char start_ref[16];
    int start_len = (int)(colon - range_str);
    if (start_len >= sizeof(start_ref)) return 0;
    
    strncpy_s(start_ref, sizeof(start_ref), range_str, start_len);
    start_ref[start_len] = '\0';
    
    if (!parse_cell_reference(start_ref, &range->start_row, &range->start_col)) {
        return 0;
    }
    
    // Parse end cell
    const char* end_ref = colon + 1;
    if (!parse_cell_reference(end_ref, &range->end_row, &range->end_col)) {
        return 0;
    }
    
    // Ensure start <= end
    if (range->start_row > range->end_row) {
        int temp = range->start_row;
        range->start_row = range->end_row;
        range->end_row = temp;
    }
    if (range->start_col > range->end_col) {
        int temp = range->start_col;
        range->start_col = range->end_col;
        range->end_col = temp;
    }
    
    return 1;
}

// Get values from a range of cells
int get_range_values(Sheet* sheet, const CellRange* range, double* values, int max_values) {
    // Optimized: Pass pre-allocated buffer or use iterator pattern
    // Avoid creating temporary array every time
    int count = 0;
    
    // Early exit if invalid range
    if (!range || max_values <= 0) return 0;
    
    for (int row = range->start_row; row <= range->end_row; row++) {
        if (count >= max_values) break;
        
        for (int col = range->start_col; col <= range->end_col; col++) {
            if (count >= max_values) break;
            
            Cell* cell = sheet_get_cell(sheet, row, col);
            if (cell) {
                switch (cell->type) {
                    case CELL_NUMBER:
                        values[count++] = cell->data.number;
                        break;
                    case CELL_FORMULA:
                        if (cell->data.formula.error == ERROR_NONE) {
                            values[count++] = cell->data.formula.cached_value;
                        }
                        break;
                    case CELL_EMPTY:
                        values[count++] = 0.0;
                        break;
                    default:
                        // Skip strings and errors
                        break;
                }
            } else {
                values[count++] = 0.0;
            }
        }
    }
    
    return count;
}

// Function implementations
double func_sum(const double* values, int count) {
    double sum = 0.0;
    for (int i = 0; i < count; i++) {
        sum += values[i];
    }
    return sum;
}

double func_avg(const double* values, int count) {
    if (count == 0) return 0.0;
    return func_sum(values, count) / count;
}

double func_max(const double* values, int count) {
    if (count == 0) return 0.0;
    double max = values[0];
    for (int i = 1; i < count; i++) {
        if (values[i] > max) max = values[i];
    }
    return max;
}

double func_min(const double* values, int count) {
    if (count == 0) return 0.0;
    double min = values[0];
    for (int i = 1; i < count; i++) {
        if (values[i] < min) min = values[i];
    }
    return min;
}

// Helper function for median calculation
int compare_double(const void* a, const void* b) {
    double da = *(const double*)a;
    double db = *(const double*)b;
    if (da < db) return -1;
    if (da > db) return 1;
    return 0;
}

double func_median(double* values, int count) {
    if (count == 0) return 0.0;
    
    // Sort the values
    qsort(values, count, sizeof(double), compare_double);
    
    if (count % 2 == 0) {
        // Even number of values - return average of middle two
        return (values[count/2 - 1] + values[count/2]) / 2.0;
    } else {
        // Odd number of values - return middle value
        return values[count/2];
    }
}

double func_mode(const double* values, int count) {
    if (count == 0) return 0.0;
    
    // Simple mode calculation - find most frequent value
    double mode = values[0];
    int max_count = 1;
    
    for (int i = 0; i < count; i++) {
        int current_count = 1;
        for (int j = i + 1; j < count; j++) {
            if (fabs(values[i] - values[j]) < FLOAT_COMPARISON_EPSILON) {
                current_count++;
            }
        }
        if (current_count > max_count) {
            max_count = current_count;
            mode = values[i];
        }
    }
    
    return mode;
}

double func_if(double condition, double true_val, double false_val) {
    return (condition != 0.0) ? true_val : false_val;
}

double func_if_enhanced(double condition, double true_val, double false_val, 
                       const char* true_str, const char* false_str) {
    // Reset global string result flag
    g_if_result_is_string = 0;
    g_if_result_string[0] = '\0';
    
    // Also reset the current cell's string result
    if (g_current_evaluating_cell) {
        if (g_current_evaluating_cell->data.formula.cached_string) {
            free(g_current_evaluating_cell->data.formula.cached_string);
            g_current_evaluating_cell->data.formula.cached_string = NULL;
        }
        g_current_evaluating_cell->data.formula.is_string_result = 0;
    }
    
    if (condition != 0.0) {
        // True condition
        if (true_str) {
            strcpy_s(g_if_result_string, sizeof(g_if_result_string), true_str);
            g_if_result_is_string = 1;
            
            // Store in current cell if available
            if (g_current_evaluating_cell) {
                char* cached = _strdup(true_str);
                if (cached) {
                    g_current_evaluating_cell->data.formula.cached_string = cached;
                    g_current_evaluating_cell->data.formula.is_string_result = 1;
                }
            }
            
            return 1.0; // Return non-zero to indicate string result available
        } else {
            return true_val;
        }
    } else {
        // False condition
        if (false_str) {
            strcpy_s(g_if_result_string, sizeof(g_if_result_string), false_str);
            g_if_result_is_string = 1;
            
            // Store in current cell if available
            if (g_current_evaluating_cell) {
                char* cached = _strdup(false_str);
                if (cached) {
                    g_current_evaluating_cell->data.formula.cached_string = cached;
                    g_current_evaluating_cell->data.formula.is_string_result = 1;
                }
            }
            
            return 0.0; // Return zero to indicate string result available
        } else {
            return false_val;
        }
    }
}

double func_power(double base, double exponent) {
    return pow(base, exponent);
}

// Skip whitespace in expression
void skip_whitespace(const char** expr) {
    while (**expr && isspace(**expr)) {
        (*expr)++;
    }
}

// Clipboard functionality
static Cell* clipboard_cell = NULL;

Cell* sheet_get_clipboard_cell(void) {
    return clipboard_cell;
}

void sheet_set_clipboard_cell(Cell* cell) {
    // Free existing clipboard cell
    if (clipboard_cell) {
        cell_free(clipboard_cell);
    }
    
    // Create a copy of the cell
    if (cell) {
        clipboard_cell = cell_new(cell->row, cell->col);
        switch (cell->type) {
            case CELL_NUMBER:
                cell_set_number(clipboard_cell, cell->data.number);
                break;
            case CELL_STRING:
                cell_set_string(clipboard_cell, cell->data.string);
                break;
            case CELL_FORMULA:
                cell_set_formula(clipboard_cell, cell->data.formula.expression);
                clipboard_cell->data.formula.cached_value = cell->data.formula.cached_value;
                clipboard_cell->data.formula.error = cell->data.formula.error;
                break;
            default:
                clipboard_cell->type = CELL_EMPTY;
                break;
        }
          // Copy display properties
        clipboard_cell->width = cell->width;
        clipboard_cell->precision = cell->precision;
        clipboard_cell->align = cell->align;
        
        // Copy color formatting
        clipboard_cell->text_color = cell->text_color;
        clipboard_cell->background_color = cell->background_color;
        clipboard_cell->row_height = cell->row_height;
    } else {
        clipboard_cell = NULL;
    }
}

void sheet_copy_cell(Sheet* sheet, int src_row, int src_col, int dest_row, int dest_col) {
    Cell* src_cell = sheet_get_cell(sheet, src_row, src_col);
    
    if (!src_cell) {
        // Clear destination cell if source is empty
        sheet_clear_cell(sheet, dest_row, dest_col);
        return;
    }
    
    switch (src_cell->type) {
        case CELL_NUMBER:
            sheet_set_number(sheet, dest_row, dest_col, src_cell->data.number);
            break;
        case CELL_STRING:
            sheet_set_string(sheet, dest_row, dest_col, src_cell->data.string);
            break;
        case CELL_FORMULA:
            sheet_set_formula(sheet, dest_row, dest_col, src_cell->data.formula.expression);
            break;
        default:
            sheet_clear_cell(sheet, dest_row, dest_col);
            break;
    }
      // Copy display properties
    Cell* dest_cell = sheet_get_cell(sheet, dest_row, dest_col);
    if (dest_cell) {
        dest_cell->width = src_cell->width;
        dest_cell->precision = src_cell->precision;
        dest_cell->align = src_cell->align;
        
        // Copy color formatting
        dest_cell->text_color = src_cell->text_color;
        dest_cell->background_color = src_cell->background_color;
        dest_cell->row_height = src_cell->row_height;
    }
    
    sheet_recalculate(sheet);
}

// Range selection functions
void sheet_start_range_selection(Sheet* sheet, int row, int col) {
    sheet->selection.start_row = row;
    sheet->selection.start_col = col;
    sheet->selection.end_row = row;
    sheet->selection.end_col = col;
    sheet->selection.is_active = 1;
}

void sheet_extend_range_selection(Sheet* sheet, int row, int col) {
    if (sheet->selection.is_active) {
        sheet->selection.end_row = row;
        sheet->selection.end_col = col;
    }
}

void sheet_clear_range_selection(Sheet* sheet) {
    sheet->selection.is_active = 0;
}

int sheet_is_in_selection(Sheet* sheet, int row, int col) {
    if (!sheet->selection.is_active) return 0;
    
    int min_row = sheet->selection.start_row < sheet->selection.end_row ? 
                  sheet->selection.start_row : sheet->selection.end_row;
    int max_row = sheet->selection.start_row > sheet->selection.end_row ? 
                  sheet->selection.start_row : sheet->selection.end_row;
    int min_col = sheet->selection.start_col < sheet->selection.end_col ? 
                  sheet->selection.start_col : sheet->selection.end_col;
    int max_col = sheet->selection.start_col > sheet->selection.end_col ? 
                  sheet->selection.start_col : sheet->selection.end_col;
    
    return (row >= min_row && row <= max_row && col >= min_col && col <= max_col);
}

// Copy range to clipboard
void sheet_copy_range(Sheet* sheet) {
    if (!sheet->selection.is_active) return;
    
    // Free existing clipboard
    if (sheet->range_clipboard.cells) {
        for (int i = 0; i < sheet->range_clipboard.rows; i++) {
            for (int j = 0; j < sheet->range_clipboard.cols; j++) {
                if (sheet->range_clipboard.cells[i][j]) {
                    cell_free(sheet->range_clipboard.cells[i][j]);
                }
            }
            free(sheet->range_clipboard.cells[i]);
        }
        free(sheet->range_clipboard.cells);
    }
    
    // Calculate range dimensions
    int min_row = sheet->selection.start_row < sheet->selection.end_row ? 
                  sheet->selection.start_row : sheet->selection.end_row;
    int max_row = sheet->selection.start_row > sheet->selection.end_row ? 
                  sheet->selection.start_row : sheet->selection.end_row;
    int min_col = sheet->selection.start_col < sheet->selection.end_col ? 
                  sheet->selection.start_col : sheet->selection.end_col;
    int max_col = sheet->selection.start_col > sheet->selection.end_col ? 
                  sheet->selection.start_col : sheet->selection.end_col;
    
    int rows = max_row - min_row + 1;
    int cols = max_col - min_col + 1;
    
    // Allocate clipboard
    sheet->range_clipboard.rows = rows;
    sheet->range_clipboard.cols = cols;
    sheet->range_clipboard.cells = (Cell***)calloc(rows, sizeof(Cell**));
    if (!sheet->range_clipboard.cells) {
        // Handle allocation failure
        return;
    }
    
    for (int i = 0; i < rows; i++) {
        sheet->range_clipboard.cells[i] = (Cell**)calloc(cols, sizeof(Cell*));
        if (!sheet->range_clipboard.cells[i]) {
            // Handle allocation failure
            return;
        }
    }
    
    // Copy cells
    for (int i = 0; i < rows; i++) {
        for (int j = 0; j < cols; j++) {
            Cell* src_cell = sheet_get_cell(sheet, min_row + i, min_col + j);
            if (src_cell) {
                Cell* copied_cell = cell_new(min_row + i, min_col + j);
                switch (src_cell->type) {
                    case CELL_NUMBER:
                        cell_set_number(copied_cell, src_cell->data.number);
                        break;
                    case CELL_STRING:
                        cell_set_string(copied_cell, src_cell->data.string);
                        break;
                    case CELL_FORMULA:
                        cell_set_formula(copied_cell, src_cell->data.formula.expression);
                        copied_cell->data.formula.cached_value = src_cell->data.formula.cached_value;
                        copied_cell->data.formula.error = src_cell->data.formula.error;
                        break;
                    default:
                        break;
                }
                // Copy display and formatting properties
                copied_cell->width = src_cell->width;
                copied_cell->precision = src_cell->precision;
                copied_cell->align = src_cell->align;
                copied_cell->format = src_cell->format;
                copied_cell->format_style = src_cell->format_style;
                
                sheet->range_clipboard.cells[i][j] = copied_cell;
            }
        }
    }
    
    sheet->range_clipboard.is_active = 1;
}

// Paste range from clipboard
void sheet_paste_range(Sheet* sheet, int start_row, int start_col) {
    if (!sheet->range_clipboard.is_active) return;
    
    for (int i = 0; i < sheet->range_clipboard.rows; i++) {
        for (int j = 0; j < sheet->range_clipboard.cols; j++) {
            int dest_row = start_row + i;
            int dest_col = start_col + j;
            
            if (dest_row >= sheet->rows || dest_col >= sheet->cols) continue;
            
            Cell* src_cell = sheet->range_clipboard.cells[i][j];
            if (src_cell) {
                Cell* dest_cell = sheet_get_or_create_cell(sheet, dest_row, dest_col);
                if (dest_cell) {
                    switch (src_cell->type) {
                        case CELL_NUMBER:
                            cell_set_number(dest_cell, src_cell->data.number);
                            break;
                        case CELL_STRING:
                            cell_set_string(dest_cell, src_cell->data.string);
                            break;
                        case CELL_FORMULA:
                            cell_set_formula(dest_cell, src_cell->data.formula.expression);
                            break;
                        default:
                            cell_clear(dest_cell);
                            break;
                    }
                    // Copy display and formatting properties
                    dest_cell->width = src_cell->width;
                    dest_cell->precision = src_cell->precision;
                    dest_cell->align = src_cell->align;
                    dest_cell->format = src_cell->format;
                    dest_cell->format_style = src_cell->format_style;
                }
            } else {
                sheet_clear_cell(sheet, dest_row, dest_col);
            }
        }
    }
    
    sheet_recalculate(sheet);
}

// Cell formatting functions
void cell_set_format(Cell* cell, DataFormat format, FormatStyle style) {
    if (!cell) return;
    
    cell->format = format;
    cell->format_style = style;
}

char* format_cell_value(Cell* cell) {
    static char buffer[256];
    
    if (!cell || cell->type == CELL_EMPTY) {
        return "";
    }
    
    double value = 0.0;
    char* string_value = NULL;
    
    // Get the value to format
    switch (cell->type) {
        case CELL_NUMBER:
            value = cell->data.number;
            break;
        case CELL_FORMULA:
            if (cell->data.formula.error != ERROR_NONE) {
                switch (cell->data.formula.error) {
                    case ERROR_DIV_ZERO: return "#DIV/0!";
                    case ERROR_REF: return "#REF!";
                    case ERROR_VALUE: return "#VALUE!";
                    case ERROR_PARSE: return "#PARSE!";
                    case ERROR_NA: return "#N/A!";
                    default: return "#ERROR!";
                }
            }
            if (cell->data.formula.is_string_result && cell->data.formula.cached_string) {
                return cell->data.formula.cached_string;
            }
            value = cell->data.formula.cached_value;
            break;
        case CELL_STRING:
            return cell->data.string;
        default:
            return "";
    }
    
    // Apply formatting based on format type
    switch (cell->format) {
        case FORMAT_PERCENTAGE:
            return format_number_as_percentage(value, cell->precision);
            
        case FORMAT_CURRENCY:
            return format_number_as_currency(value);
            
        case FORMAT_DATE:
            return format_number_as_date(value, cell->format_style);
            
        case FORMAT_TIME:
            return format_number_as_time(value, cell->format_style);
              case FORMAT_DATETIME:
            // Check if it's one of the enhanced datetime styles
            if (cell->format_style == DATETIME_STYLE_SHORT || 
                cell->format_style == DATETIME_STYLE_LONG || 
                cell->format_style == DATETIME_STYLE_ISO) {
                return format_number_as_enhanced_datetime(value, cell->format_style);
            } else {
                return format_number_as_datetime(value, DATE_STYLE_MM_DD_YYYY, TIME_STYLE_12HR);
            }
            
        case FORMAT_NUMBER:
        case FORMAT_GENERAL:
        default:
            // Standard number formatting
            snprintf(buffer, sizeof(buffer), "%.*f", cell->precision, value);
            // Remove trailing zeros
            char* dot = strchr(buffer, '.');
            if (dot) {
                char* end = buffer + strlen(buffer) - 1;
                while (end > dot && *end == '0') *end-- = '\0';
                if (*end == '.') *end = '\0';
            }
            return buffer;
    }
}

char* format_number_as_percentage(double value, int precision) {
    static char buffer[64];
    snprintf(buffer, sizeof(buffer), "%.*f%%", precision, value * 100.0);
    return buffer;
}

char* format_number_as_currency(double value) {
    static char buffer[64];
    if (value < 0) {
        snprintf(buffer, sizeof(buffer), "-$%.2f", -value);
    } else {
        snprintf(buffer, sizeof(buffer), "$%.2f", value);
    }
    return buffer;
}

char* format_number_as_date(double value, FormatStyle style) {
    static char buffer[64];
    
    // Convert Excel serial date to time_t (days since 1900-01-01)
    // Excel incorrectly treats 1900 as a leap year, so we need adjustment
    time_t base_time = (time_t)-2209161600LL;  // 1900-01-01 00:00:00 UTC (adjusted)
    time_t date_time = base_time + (time_t)(value * 86400); // 86400 seconds per day
    
    struct tm date_struct;
    if (gmtime_s(&date_struct, &date_time) != 0) {
        strcpy_s(buffer, sizeof(buffer), "#DATE!");
        return buffer;
    }
    
    switch (style) {
        case DATE_STYLE_MM_DD_YYYY:
            strftime(buffer, sizeof(buffer), "%m/%d/%Y", &date_struct);
            break;
        case DATE_STYLE_DD_MM_YYYY:
            strftime(buffer, sizeof(buffer), "%d/%m/%Y", &date_struct);
            break;
        case DATE_STYLE_YYYY_MM_DD:
            strftime(buffer, sizeof(buffer), "%Y-%m-%d", &date_struct);
            break;
        case DATE_STYLE_MON_DD_YYYY:
            strftime(buffer, sizeof(buffer), "%b %d, %Y", &date_struct);
            break;
        case DATE_STYLE_DD_MON_YYYY:
            strftime(buffer, sizeof(buffer), "%d %b %Y", &date_struct);
            break;
        case DATE_STYLE_YYYY_MON_DD:
            strftime(buffer, sizeof(buffer), "%Y %b %d", &date_struct);
            break;
        case DATE_STYLE_SHORT_DATE:
            strftime(buffer, sizeof(buffer), "%m/%d/%y", &date_struct);
            break;
        default:
            strftime(buffer, sizeof(buffer), "%Y-%m-%d", &date_struct);
            break;
    }
    
    return buffer;
}

char* format_number_as_time(double value, FormatStyle style) {
    static char buffer[64];
    
    // Extract time portion (fractional part of the day)
    double time_fraction = value - floor(value);
    if (time_fraction < 0) time_fraction += 1.0;
    
    int total_seconds = (int)(time_fraction * 86400);
    int hours = total_seconds / 3600;
    int minutes = (total_seconds % 3600) / 60;
    int seconds = total_seconds % 60;
    
    switch (style) {
        case TIME_STYLE_12HR:
            {
                int display_hours = hours;
                const char* am_pm = "AM";
                if (hours == 0) {
                    display_hours = 12;
                } else if (hours == 12) {
                    am_pm = "PM";
                } else if (hours > 12) {
                    display_hours = hours - 12;
                    am_pm = "PM";
                }
                snprintf(buffer, sizeof(buffer), "%d:%02d %s", display_hours, minutes, am_pm);
            }
            break;
        case TIME_STYLE_12HR_SECONDS:
            {
                int display_hours = hours;
                const char* am_pm = "AM";
                if (hours == 0) {
                    display_hours = 12;
                } else if (hours == 12) {
                    am_pm = "PM";
                } else if (hours > 12) {
                    display_hours = hours - 12;
                    am_pm = "PM";
                }
                snprintf(buffer, sizeof(buffer), "%d:%02d:%02d %s", display_hours, minutes, seconds, am_pm);
            }
            break;
        case TIME_STYLE_SECONDS:
            snprintf(buffer, sizeof(buffer), "%02d:%02d:%02d", hours, minutes, seconds);
            break;
        case TIME_STYLE_24HR:
        default:
            snprintf(buffer, sizeof(buffer), "%02d:%02d", hours, minutes);
            break;
    }
    
    return buffer;
}

char* format_number_as_datetime(double value, FormatStyle date_style, FormatStyle time_style) {
    static char buffer[128];
    char* date_part = format_number_as_date(value, date_style);
    char* time_part = format_number_as_time(value, time_style);
    snprintf(buffer, sizeof(buffer), "%s %s", date_part, time_part);
    return buffer;
}

// Enhanced datetime formatting for special styles
char* format_number_as_enhanced_datetime(double value, FormatStyle style) {
    static char buffer[128];
    
    // Convert Excel serial date to time_t
    time_t base_time = (time_t)-2209161600LL;  // 1900-01-01 00:00:00 UTC (adjusted)
    time_t date_time = base_time + (time_t)(value * 86400);
    
    struct tm date_struct;
    if (gmtime_s(&date_struct, &date_time) != 0) {
        strcpy_s(buffer, sizeof(buffer), "#DATE!");
        return buffer;
    }
    
    // Extract time portion for time components
    double time_fraction = value - floor(value);
    if (time_fraction < 0) time_fraction += 1.0;
    int total_seconds = (int)(time_fraction * 86400);
    int hours = total_seconds / 3600;
    int minutes = (total_seconds % 3600) / 60;
    int seconds = total_seconds % 60;
    
    switch (style) {
        case DATETIME_STYLE_SHORT:
            {
                int display_hours = hours;
                const char* am_pm = "AM";
                if (hours == 0) {
                    display_hours = 12;
                } else if (hours == 12) {
                    am_pm = "PM";
                } else if (hours > 12) {
                    display_hours = hours - 12;
                    am_pm = "PM";
                }
                snprintf(buffer, sizeof(buffer), "%d/%d/%02d %d:%02d %s", 
                        date_struct.tm_mon + 1, date_struct.tm_mday, date_struct.tm_year % 100,
                        display_hours, minutes, am_pm);
            }
            break;
        case DATETIME_STYLE_LONG:
            {
                int display_hours = hours;
                const char* am_pm = "AM";
                if (hours == 0) {
                    display_hours = 12;
                } else if (hours == 12) {
                    am_pm = "PM";
                } else if (hours > 12) {
                    display_hours = hours - 12;
                    am_pm = "PM";
                }
                strftime(buffer, 80, "%b %d, %Y ", &date_struct);
                char time_buf[48];
                snprintf(time_buf, sizeof(time_buf), "%d:%02d:%02d %s", 
                        display_hours, minutes, seconds, am_pm);
                strcat_s(buffer, sizeof(buffer), time_buf);
            }
            break;
        case DATETIME_STYLE_ISO:
            snprintf(buffer, sizeof(buffer), "%04d-%02d-%02dT%02d:%02d:%02d",
                    date_struct.tm_year + 1900, date_struct.tm_mon + 1, date_struct.tm_mday,
                    hours, minutes, seconds);
            break;
        default:
            // Fall back to regular datetime formatting
            return format_number_as_datetime(value, DATE_STYLE_MM_DD_YYYY, TIME_STYLE_12HR);
    }
    
    return buffer;
}

// XLOOKUP function - more flexible than VLOOKUP
// Searches in lookup_array and returns corresponding value from return_array
double func_xlookup(Sheet* sheet, double lookup_value, const char* lookup_str,
                   const char* lookup_array, const char* return_array, int exact_match, ErrorType* error) {
    *error = ERROR_NONE;
    
    // Parse the lookup array range
    CellRange lookup_range;
    if (!parse_range(lookup_array, &lookup_range)) {
        *error = ERROR_REF;
        return 0.0;
    }
    
    // Parse the return array range
    CellRange return_range;
    if (!parse_range(return_array, &return_range)) {
        *error = ERROR_REF;
        return 0.0;
    }
    
    // Validate ranges have same dimensions
    int lookup_rows = lookup_range.end_row - lookup_range.start_row + 1;
    int lookup_cols = lookup_range.end_col - lookup_range.start_col + 1;
    int return_rows = return_range.end_row - return_range.start_row + 1;
    int return_cols = return_range.end_col - return_range.start_col + 1;
    
    if (lookup_rows != return_rows || lookup_cols != return_cols) {
        *error = ERROR_REF;
        return 0.0;
    }
    
    // Search through the lookup array
    // Support both vertical (column) and horizontal (row) searches
    int is_vertical = (lookup_rows > 1);
    int search_count = is_vertical ? lookup_rows : lookup_cols;
    
    for (int i = 0; i < search_count; i++) {
        int lookup_row = is_vertical ? (lookup_range.start_row + i) : lookup_range.start_row;
        int lookup_col = is_vertical ? lookup_range.start_col : (lookup_range.start_col + i);
        Cell* lookup_cell = sheet_get_cell(sheet, lookup_row, lookup_col);
        if (!lookup_cell) continue;
        
        int match_found = 0;
        
        // Compare based on data type
        if (lookup_str) {
            // String lookup
            if (lookup_cell->type == CELL_STRING) {
                match_found = (strcmp(lookup_cell->data.string, lookup_str) == 0);
            } else if (lookup_cell->type == CELL_FORMULA && 
                      lookup_cell->data.formula.is_string_result && 
                      lookup_cell->data.formula.cached_string) {
                match_found = (strcmp(lookup_cell->data.formula.cached_string, lookup_str) == 0);
            }
        } else {
            // Numeric lookup
            double cell_value = 0.0;
            if (lookup_cell->type == CELL_NUMBER) {
                cell_value = lookup_cell->data.number;
            } else if (lookup_cell->type == CELL_FORMULA && 
                      lookup_cell->data.formula.error == ERROR_NONE) {
                cell_value = lookup_cell->data.formula.cached_value;
            } else {
                continue; // Skip non-numeric cells
            }
            
            if (exact_match) {
                match_found = (fabs(cell_value - lookup_value) < 1e-10);
            } else {
                // Approximate match - find largest value <= lookup_value
                match_found = (cell_value <= lookup_value);
                // For approximate match, we should continue to find the best match
                // This is a simplified implementation
                if (match_found) {
                    // Check if there's a better match later
                    for (int next_i = i + 1; next_i < search_count; next_i++) {
                        int next_lookup_row = is_vertical ? (lookup_range.start_row + next_i) : lookup_range.start_row;
                        int next_lookup_col = is_vertical ? lookup_range.start_col : (lookup_range.start_col + next_i);
                        Cell* next_cell = sheet_get_cell(sheet, next_lookup_row, next_lookup_col);
                        if (next_cell) {
                            double next_value = 0.0;
                            if (next_cell->type == CELL_NUMBER) {
                                next_value = next_cell->data.number;
                            } else if (next_cell->type == CELL_FORMULA && 
                                      next_cell->data.formula.error == ERROR_NONE) {
                                next_value = next_cell->data.formula.cached_value;
                            } else {
                                continue;
                            }
                            
                            if (next_value <= lookup_value && next_value > cell_value) {
                                // Found a better match
                                match_found = 0;
                                break;
                            }
                        }
                    }
                }
            }
        }
        
        if (match_found) {
            // Found a match, return the corresponding value from the return array
            int return_row = is_vertical ? (return_range.start_row + i) : return_range.start_row;
            int return_col = is_vertical ? return_range.start_col : (return_range.start_col + i);
            Cell* result_cell = sheet_get_cell(sheet, return_row, return_col);
            
            if (!result_cell) {
                return 0.0; // Empty cell
            }
            
            switch (result_cell->type) {
                case CELL_NUMBER:
                    return result_cell->data.number;
                case CELL_FORMULA:
                    if (result_cell->data.formula.error == ERROR_NONE) {
                        return result_cell->data.formula.cached_value;
                    }
                    break;
                case CELL_EMPTY:
                    return 0.0;
                default:
                    break;
            }
        }
    }
    
    // No match found
    *error = ERROR_NA;
    return 0.0;
}

// Simple formula evaluator (basic arithmetic and cell references)
double evaluate_formula(Sheet* sheet, const char* formula, ErrorType* error) {
    *error = ERROR_NONE;
    
    // Skip leading '='
    if (*formula == '=') formula++;
    
    // Try comparison expressions first (for IF function support)
    return evaluate_comparison(sheet, formula, error);
}

// Helper function to get string value from a cell reference
char* get_cell_string_value(Sheet* sheet, const char* ref) {
    int row, col;
    if (!parse_cell_reference(ref, &row, &col)) {
        return NULL;
    }
    
    Cell* cell = sheet_get_cell(sheet, row, col);
    if (!cell) return NULL;
    
    switch (cell->type) {
        case CELL_STRING:
            return cell->data.string;
        case CELL_FORMULA:
            if (cell->data.formula.is_string_result && cell->data.formula.cached_string) {
                return cell->data.formula.cached_string;
            }
            break;
        default:
            break;
    }
    return NULL;
}

// Enhanced comparison evaluation with string support
double evaluate_comparison(Sheet* sheet, const char* expr, ErrorType* error) {
    const char* p = expr;
    const char* start_left = p;
    
    // First, try to identify if we have a string comparison
    // Look for pattern: cell_ref = "string" or "string" = cell_ref
    
    // Check if left side is a cell reference
    char left_ref[32] = {0};
    int left_ref_len = 0;
    const char* temp_p = p;
    
    // Extract potential cell reference
    while (*temp_p && (isalpha(*temp_p) || isdigit(*temp_p)) && left_ref_len < 31) {
        left_ref[left_ref_len++] = *temp_p;
        temp_p++;
    }
    left_ref[left_ref_len] = '\0';
    
    // Check if it's a valid cell reference
    int is_left_cell_ref = 0;
    int left_row, left_col;
    if (left_ref_len > 0 && parse_cell_reference(left_ref, &left_row, &left_col)) {
        is_left_cell_ref = 1;
    }
    
    // Skip whitespace and look for comparison operator
    skip_whitespace(&temp_p);
    char comparison_op[3] = {0};
    int op_len = 0;
    
    if (*temp_p == '=' || *temp_p == '<' || *temp_p == '>') {
        comparison_op[op_len++] = *temp_p++;
        if (*temp_p == '=' || (*temp_p == '>' && comparison_op[0] == '<')) {
            comparison_op[op_len++] = *temp_p++;
        }
        comparison_op[op_len] = '\0';
    }
    
    // Check if right side is a string literal
    skip_whitespace(&temp_p);
    int is_right_string = (*temp_p == '"');
    
    // If we have cell_ref = "string", handle as string comparison
    if (is_left_cell_ref && strlen(comparison_op) > 0 && is_right_string) {
        // Get the string value from the cell
        char* left_str = get_cell_string_value(sheet, left_ref);
        
        // Parse the string literal
        temp_p++; // Skip opening quote
        char right_str[256] = {0};
        int right_len = 0;
        while (*temp_p && *temp_p != '"' && right_len < 255) {
            right_str[right_len++] = *temp_p++;
        }
        if (*temp_p == '"') temp_p++; // Skip closing quote
        
        // Perform string comparison
        int cmp_result = 0;
        if (left_str) {
            cmp_result = strcmp(left_str, right_str);
        } else {
            // If cell is empty or not string, compare with empty string
            cmp_result = strcmp("", right_str);
        }
        
        // Apply comparison operator
        if (strcmp(comparison_op, "=") == 0) {
            return (cmp_result == 0) ? 1.0 : 0.0;
        } else if (strcmp(comparison_op, "<>") == 0) {
            return (cmp_result != 0) ? 1.0 : 0.0;
        } else if (strcmp(comparison_op, "<") == 0) {
            return (cmp_result < 0) ? 1.0 : 0.0;
        } else if (strcmp(comparison_op, "<=") == 0) {
            return (cmp_result <= 0) ? 1.0 : 0.0;
        } else if (strcmp(comparison_op, ">") == 0) {
            return (cmp_result > 0) ? 1.0 : 0.0;
        } else if (strcmp(comparison_op, ">=") == 0) {
            return (cmp_result >= 0) ? 1.0 : 0.0;
        }
        
        *error = ERROR_PARSE;
        return 0.0;
    }
    
    // Fall back to numeric comparison
    double left = parse_arithmetic_expression(sheet, &p, error);
    if (*error != ERROR_NONE) return left;
    
    skip_whitespace(&p);
    
    // Check for comparison operators
    if (*p == '>') {
        p++;
        if (*p == '=') {
            p++; // >=
            double right = parse_arithmetic_expression(sheet, &p, error);
            return (left >= right) ? 1.0 : 0.0;
        } else {
            // >
            double right = parse_arithmetic_expression(sheet, &p, error);
            return (left > right) ? 1.0 : 0.0;
        }
    } else if (*p == '<') {
        p++;
        if (*p == '=') {
            p++; // <=
            double right = parse_arithmetic_expression(sheet, &p, error);
            return (left <= right) ? 1.0 : 0.0;
        } else if (*p == '>') {
            p++; // <>
            double right = parse_arithmetic_expression(sheet, &p, error);
            return (left != right) ? 1.0 : 0.0;
        } else {
            // <
            double right = parse_arithmetic_expression(sheet, &p, error);
            return (left < right) ? 1.0 : 0.0;
        }
    } else if (*p == '=') {
        p++; // =
        double right = parse_arithmetic_expression(sheet, &p, error);
        return (fabs(left - right) < FLOAT_COMPARISON_EPSILON) ? 1.0 : 0.0;
    }
    
    // No comparison operator found, return the arithmetic expression result
    return left;
}

// Parse term (handles * and /)
double parse_term(Sheet* sheet, const char** expr, ErrorType* error) {
    double result = parse_factor(sheet, expr, error);
    if (*error != ERROR_NONE) return 0.0;
    
    while (1) {
        skip_whitespace(expr);
        if (**expr == '*') {
            (*expr)++;
            double right = parse_factor(sheet, expr, error);
            if (*error != ERROR_NONE) return 0.0;
            result *= right;
        } else if (**expr == '/') {
            (*expr)++;
            double right = parse_factor(sheet, expr, error);
            if (*error != ERROR_NONE) return 0.0;
            if (right == 0.0) {
                *error = ERROR_DIV_ZERO;
                return 0.0;
            }
            result /= right;
        } else {
            break;
        }
    }
    
    return result;
}

// Parse a factor (number, cell reference, or function call)
double parse_factor(Sheet* sheet, const char** expr, ErrorType* error) {
    skip_whitespace(expr);
    
    if (**expr == '(') {
        (*expr)++; // Skip '('
        double result = parse_arithmetic_expression(sheet, expr, error);
        if (*error != ERROR_NONE) return 0.0;
        
        skip_whitespace(expr);
        if (**expr != ')') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        (*expr)++; // Skip ')'
        
        return result;
    }
    
    // Check if this could be a function call (letters followed by '(')
    const char* start = *expr;
    const char* lookahead = *expr;
    
    // Look ahead to see if this is a function call
    while (*lookahead && isalpha(*lookahead)) lookahead++;
    skip_whitespace(&lookahead);
    
    if (*lookahead == '(') {
        // This is a function call
        return parse_function(sheet, expr, error);
    }
    
    // Try to parse as cell reference or range
    char ref_buf[32];
    int i = 0;
    
    // Extract potential cell reference (letters followed by numbers, possibly with colon)
    while (**expr && (isalpha(**expr) || isdigit(**expr) || **expr == ':') && i < 31) {
        ref_buf[i++] = **expr;
        (*expr)++;
    }
    ref_buf[i] = '\0';
    
    if (i > 0) {
        // Check if it's a range (contains ':')
        if (strchr(ref_buf, ':')) {
            CellRange range;
            if (parse_range(ref_buf, &range)) {
                // For a range without a function, just return the sum
                double values[1000];
                int count = get_range_values(sheet, &range, values, 1000);
                return func_sum(values, count);
            } else {
                *error = ERROR_PARSE;
                return 0.0;
            }
        } else {
            // Single cell reference
            int row, col;
            if (parse_cell_reference(ref_buf, &row, &col)) {
                Cell* cell = sheet_get_cell(sheet, row, col);
                if (!cell) {
                    return 0.0; // Empty cell
                }
                
                switch (cell->type) {
                    case CELL_EMPTY:
                        return 0.0;
                    case CELL_NUMBER:
                        return cell->data.number;
                    case CELL_FORMULA:
                        if (cell->data.formula.error != ERROR_NONE) {
                            *error = cell->data.formula.error;
                            return 0.0;
                        }
                        return cell->data.formula.cached_value;
                    default:
                        *error = ERROR_VALUE;
                        return 0.0;
                }
            }
        }
    }
    
    // Reset and try to parse as number
    *expr = start;
    char* endptr;
    double value = strtod(*expr, &endptr);
    if (endptr != *expr) {
        *expr = endptr;
        return value;
    }
    
    *error = ERROR_PARSE;
    return 0.0;
}

// Enhanced arithmetic expression parser (with updated signature)
double parse_arithmetic_expression(Sheet* sheet, const char** expr, ErrorType* error) {
    double result = parse_term(sheet, expr, error);
    if (*error != ERROR_NONE) return 0.0;
    
    while (1) {
        skip_whitespace(expr);
        if (**expr == '+') {
            (*expr)++;
            double right = parse_term(sheet, expr, error);
            if (*error != ERROR_NONE) return 0.0;
            result += right;
        } else if (**expr == '-') {
            (*expr)++;
            double right = parse_term(sheet, expr, error);
            if (*error != ERROR_NONE) return 0.0;
            result -= right;
        } else {
            break;
        }
    }
    
    return result;
}

// Recalculate all formulas in the sheet
void sheet_recalculate(Sheet* sheet) {
    if (!sheet->needs_recalc) return;
    
    // Use smart recalculation with dependency tracking
    sheet_recalculate_smart(sheet);
    
    sheet->needs_recalc = 0;
}

void sheet_recalculate_smart(Sheet* sheet) {
    // Only recalculate cells that changed and their dependents
    // Use topological sort for correct order
    
    // For now, use a simple queue-based approach
    // In a full implementation, build dependency graph and use topological sort
    
    // Build list of cells to recalculate
    Cell** to_calc = (Cell**)malloc(sheet->rows * sheet->cols * sizeof(Cell*));
    int calc_count = 0;
    
    // Add all formula cells to calculation queue
    for (int row = 0; row < sheet->rows; row++) {
        for (int col = 0; col < sheet->cols; col++) {
            Cell* cell = sheet_get_cell(sheet, row, col);
            if (cell && cell->type == CELL_FORMULA) {
                to_calc[calc_count++] = cell;
            }
        }
    }
    
    // Evaluate formulas (in dependency order if graph is built)
    for (int i = 0; i < calc_count; i++) {
        Cell* cell = to_calc[i];
        ErrorType error;
        g_current_evaluating_cell = cell;
        double value = evaluate_formula(sheet, cell->data.formula.expression, &error);
        cell->data.formula.cached_value = value;
        cell->data.formula.error = error;
        g_current_evaluating_cell = NULL;
    }
    
    free(to_calc);
}

// Escape a string for CSV output (handle quotes and commas)
char* escape_csv_string(const char* str) {
    if (!str) return NULL;
    
    // Check if we need to escape (contains comma, quote, or newline)
    int needs_escape = 0;
    int quote_count = 0;
    
    for (const char* p = str; *p; p++) {
        if (*p == ',' || *p == '\n' || *p == '\r') {
            needs_escape = 1;
        }
        if (*p == '"') {
            needs_escape = 1;
            quote_count++;
        }
    }    if (!needs_escape) {
        // Return a copy of the original string
        char* result = malloc(strlen(str) + 1);
        if (result) {
            strcpy_s(result, strlen(str) + 1, str);
        }
        return result;
    }
    
    // Need to escape - allocate buffer (original + 2 for outer quotes + quotes to double)
    char* result = malloc(strlen(str) + 2 + quote_count + 1);
    if (!result) return NULL;
    
    char* p = result;
    *p++ = '"';  // Opening quote
    
    for (const char* s = str; *s; s++) {
        if (*s == '"') {
            *p++ = '"';  // Double the quote
            *p++ = '"';
        } else {
            *p++ = *s;
        }
    }
    
    *p++ = '"';  // Closing quote
    *p = '\0';
    
    return result;
}

// Parse a CSV field (handles quoted fields with commas and quotes)
char* parse_csv_field(const char** csv_line, int* is_end) {
    const char* start = *csv_line;
    const char* p = start;
    
    *is_end = 0;
    
    // Skip leading whitespace
    while (*p == ' ' || *p == '\t') p++;
    
    if (*p == '\0' || *p == '\n' || *p == '\r') {
        *is_end = 1;
        *csv_line = p;
        return NULL;  // Empty field at end of line
    }
    
    if (*p == '"') {
        // Quoted field
        p++;  // Skip opening quote
        // Dynamically grow buffer or scan first to determine size
        size_t capacity = 256;
        char* result = malloc(capacity);
        if (!result) return NULL;
        
        char* dest = result;
        size_t len = 0;
        
        while (*p && *p != '\n' && *p != '\r') {
            if (*p == '"') {
                if (*(p + 1) == '"') {
                    // Escaped quote
                    if (len >= capacity - 1) {
                        capacity *= 2;
                        char* new_result = realloc(result, capacity);
                        if (!new_result) {
                            free(result);
                            return NULL;
                        }
                        result = new_result;
                        dest = result + len;
                    }
                    *dest++ = '"';
                    len++;
                    p += 2;
                } else {
                    // End of quoted field
                    p++;
                    break;
                }
            } else {
                if (len >= capacity - 1) {
                    capacity *= 2;
                    char* new_result = realloc(result, capacity);
                    if (!new_result) {
                        free(result);
                        return NULL;
                    }
                    result = new_result;
                    dest = result + len;
                }
                *dest++ = *p++;
                len++;
            }
        }
        
        *dest = '\0';
        
        // Skip to next comma or end of line
        while (*p && *p != ',' && *p != '\n' && *p != '\r') p++;
        if (*p == ',') p++;
        else *is_end = 1;
        
        *csv_line = p;
        return result;
    } else {
        // Unquoted field
        const char* field_start = p;
        while (*p && *p != ',' && *p != '\n' && *p != '\r') p++;
        
        int len = (int)(p - field_start);
        char* result = malloc(len + 1);
        if (!result) return NULL;
        
        strncpy_s(result, len + 1, field_start, len);
        result[len] = '\0';
        
        // Trim trailing whitespace
        while (len > 0 && (result[len-1] == ' ' || result[len-1] == '\t')) {
            result[--len] = '\0';
        }
        
        if (*p == ',') p++;
        else *is_end = 1;
        
        *csv_line = p;
        return result;
    }
}

// Save sheet to CSV file
int sheet_save_csv(Sheet* sheet, const char* filename, int preserve_formulas) {
    FILE* file;
    if (fopen_s(&file, filename, "w") != 0) {
        return 0;  // Failed to open file
    }
    
    // Find the actual used range
    int max_row = 0, max_col = 0;
    for (int row = 0; row < sheet->rows; row++) {
        for (int col = 0; col < sheet->cols; col++) {
            Cell* cell = sheet_get_cell(sheet, row, col);
            if (cell && cell->type != CELL_EMPTY) {
                if (row > max_row) max_row = row;
                if (col > max_col) max_col = col;
            }
        }
    }
    
    // Write data
    for (int row = 0; row <= max_row; row++) {
        for (int col = 0; col <= max_col; col++) {
            Cell* cell = sheet_get_cell(sheet, row, col);
            
            if (col > 0) {
                fprintf(file, ",");
            }
            
            if (cell && cell->type != CELL_EMPTY) {
                if (preserve_formulas && cell->type == CELL_FORMULA) {
                    // Save the formula expression
                    char* escaped = escape_csv_string(cell->data.formula.expression);
                    if (escaped) {
                        fprintf(file, "%s", escaped);
                        free(escaped);
                    }
                } else {
                    // Save the display value
                    char* display_value = sheet_get_display_value(sheet, row, col);
                    if (display_value && strlen(display_value) > 0) {
                        char* escaped = escape_csv_string(display_value);
                        if (escaped) {
                            fprintf(file, "%s", escaped);
                            free(escaped);
                        }
                    }
                }
            }
        }
        fprintf(file, "\n");
    }
    
    fclose(file);
    return 1;  // Success
}

// Load sheet from CSV file
int sheet_load_csv(Sheet* sheet, const char* filename, int preserve_formulas) {
    FILE* file;
    if (fopen_s(&file, filename, "r") != 0) {
        return 0;  // Failed to open file
    }
    
    // Clear existing data
    for (int row = 0; row < sheet->rows; row++) {
        for (int col = 0; col < sheet->cols; col++) {
            sheet_clear_cell(sheet, row, col);
        }
    }
    
    char line[4096];  // Buffer for reading lines
    int row = 0;
    
    while (fgets(line, sizeof(line), file) && row < sheet->rows) {
        const char* line_ptr = line;
        int col = 0;
        int is_end = 0;
        
        while (!is_end && col < sheet->cols) {
            char* field = parse_csv_field(&line_ptr, &is_end);
            
            if (field && strlen(field) > 0) {
                // Determine what type of data this is
                if (preserve_formulas && field[0] == '=') {
                    // It's a formula
                    sheet_set_formula(sheet, row, col, field);
                } else {
                    // Check if it's a number
                    char* endptr;
                    double num_value = strtod(field, &endptr);
                    
                    if (*endptr == '\0' || (*endptr == '\0' && endptr != field)) {
                        // It's a valid number
                        sheet_set_number(sheet, row, col, num_value);
                    } else {
                        // It's a string
                        sheet_set_string(sheet, row, col, field);
                    }
                }
                free(field);
            }
            
            col++;
        }
        
        row++;
    }
    
    fclose(file);
    
    // Recalculate if we loaded formulas
    if (preserve_formulas) {
        sheet_recalculate(sheet);
    }
    
    return 1;  // Success
}

// Demo: Create a simple spreadsheet with some data
void demo_spreadsheet() {
    Sheet* sheet = sheet_new(100, 26);
    
    // Add some sample data
    sheet_set_string(sheet, 0, 0, "Item");
    sheet_set_string(sheet, 0, 1, "Quantity");
    sheet_set_string(sheet, 0, 2, "Price");
    sheet_set_string(sheet, 0, 3, "Total");
    
    sheet_set_string(sheet, 1, 0, "Apples");
    sheet_set_number(sheet, 1, 1, 10);
    sheet_set_number(sheet, 1, 2, 0.5);
    sheet_set_formula(sheet, 1, 3, "=B2*C2");  // Would calculate to 5.0
    
    sheet_set_string(sheet, 2, 0, "Oranges");
    sheet_set_number(sheet, 2, 1, 15);
    sheet_set_number(sheet, 2, 2, 0.75);
    sheet_set_formula(sheet, 2, 3, "=B3*C3");  // Would calculate to 11.25
    
    // Recalculate formulas
    sheet_recalculate(sheet);
    
    // Print the sheet
    printf("\nSimple Spreadsheet Demo:\n");
    printf("------------------------\n");
    
    for (int row = 0; row <= 3; row++) {
        for (int col = 0; col <= 3; col++) {
            char* value = sheet_get_display_value(sheet, row, col);
            printf("%-12s", value);
        }
        printf("\n");
    }
      sheet_free(sheet);
}

// Enhanced parse_function to support XLOOKUP and other functions
double parse_function(Sheet* sheet, const char** expr, ErrorType* error) {
    skip_whitespace(expr);
    
    // Extract function name
    const char* start = *expr;
    char func_name[32];
    int i = 0;
    
    while (**expr && isalpha(**expr) && i < 31) {
        func_name[i++] = toupper(**expr);
        (*expr)++;
    }
    func_name[i] = '\0';
    
    skip_whitespace(expr);
    if (**expr != '(') {
        *error = ERROR_PARSE;
        return 0.0;
    }
    (*expr)++; // Skip '('
    
    // Handle XLOOKUP function
    if (strcmp(func_name, "XLOOKUP") == 0) {
        // XLOOKUP(lookup_value, lookup_array, return_array, [match_mode])
        skip_whitespace(expr);
        
        // Parse lookup value (can be number or string)
        double lookup_value = 0.0;
        char lookup_str[256] = {0};
        int is_string_lookup = 0;
        
        if (**expr == '"') {
            // Parse string literal
            (*expr)++; // Skip opening quote
            int j = 0;
            while (**expr && **expr != '"' && j < 255) {
                lookup_str[j++] = **expr;
                (*expr)++;
            }
            if (**expr == '"') {
                (*expr)++; // Skip closing quote
                is_string_lookup = 1;
            } else {
                *error = ERROR_PARSE;
                return 0.0;
            }
        } else {
            // Parse as numeric expression
            lookup_value = parse_arithmetic_expression(sheet, expr, error);
            if (*error != ERROR_NONE) return 0.0;
        }
        
        // Expect comma
        skip_whitespace(expr);
        if (**expr != ',') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        (*expr)++;
        
        // Parse lookup_array range
        skip_whitespace(expr);
        char lookup_array[64];
        int lookup_len = 0;
        
        while (**expr && **expr != ',' && lookup_len < 63) {
            lookup_array[lookup_len++] = **expr;
            (*expr)++;
        }
        lookup_array[lookup_len] = '\0';
        
        // Expect comma
        skip_whitespace(expr);
        if (**expr != ',') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        (*expr)++;
        
        // Parse return_array range
        skip_whitespace(expr);
        char return_array[64];
        int return_len = 0;
        
        while (**expr && **expr != ',' && **expr != ')' && return_len < 63) {
            return_array[return_len++] = **expr;
            (*expr)++;
        }
        return_array[return_len] = '\0';
        
        // Parse optional match_mode parameter (default is exact match)
        int exact_match = 1; // Default to exact match (unlike VLOOKUP)
        skip_whitespace(expr);
        if (**expr == ',') {
            (*expr)++;
            skip_whitespace(expr);
            double mode_f = parse_arithmetic_expression(sheet, expr, error);
            if (*error != ERROR_NONE) return 0.0;
            exact_match = (mode_f == 0.0) ? 1 : 0; // 0 = exact (default), 1 = approximate
        }
        
        // Expect closing parenthesis
        skip_whitespace(expr);
        if (**expr != ')') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        (*expr)++;
        
        // Call XLOOKUP function
        return func_xlookup(sheet, lookup_value, is_string_lookup ? lookup_str : NULL,
                           lookup_array, return_array, exact_match, error);
    }
    
    // Handle other existing functions (simplified for basic functionality)
    double values[MAX_RANGE_VALUES];
    int value_count = 0;
    
    if (strcmp(func_name, "SUM") == 0 || strcmp(func_name, "AVG") == 0 || 
        strcmp(func_name, "MAX") == 0 || strcmp(func_name, "MIN") == 0 ||
        strcmp(func_name, "MEDIAN") == 0 || strcmp(func_name, "MODE") == 0) {
        
        // Parse range or single values
        skip_whitespace(expr);
        const char* arg_start = *expr;
        
        // Find the end of the argument (look for closing parenthesis)
        int paren_count = 1;
        const char* arg_end = *expr;
        while (*arg_end && paren_count > 0) {
            if (*arg_end == '(') paren_count++;
            else if (*arg_end == ')') paren_count--;
            if (paren_count > 0) arg_end++;
        }
        
        // Extract the argument
        int arg_len = (int)(arg_end - arg_start);
        char arg[256];
        if (arg_len >= sizeof(arg)) {
            *error = ERROR_PARSE;
            return 0.0;
        }
        
        strncpy_s(arg, sizeof(arg), arg_start, arg_len);
        arg[arg_len] = '\0';
        
        // Check if it's a range (contains ':')
        if (strchr(arg, ':')) {
            CellRange range;
            if (parse_range(arg, &range)) {
                value_count = get_range_values(sheet, &range, values, MAX_RANGE_VALUES);
            } else {
                *error = ERROR_PARSE;
                return 0.0;
            }
        } else {
            // Single cell reference or value
            int row, col;
            if (parse_cell_reference(arg, &row, &col)) {
                Cell* cell = sheet_get_cell(sheet, row, col);
                if (cell) {
                    switch (cell->type) {
                        case CELL_NUMBER:
                            values[0] = cell->data.number;
                            value_count = 1;
                            break;
                        case CELL_FORMULA:
                            if (cell->data.formula.error == ERROR_NONE) {
                                values[0] = cell->data.formula.cached_value;
                                value_count = 1;
                            }
                            break;
                        case CELL_EMPTY:
                            values[0] = 0.0;
                            value_count = 1;
                            break;
                        default:
                            *error = ERROR_VALUE;
                            return 0.0;
                    }
                } else {
                    values[0] = 0.0;
                    value_count = 1;
                }
            } else {
                // Try to parse as number
                char* endptr;
                double val = strtod(arg, &endptr);
                if (*endptr == '\0') {
                    values[0] = val;
                    value_count = 1;
                } else {
                    *error = ERROR_PARSE;
                    return 0.0;
                }
            }
        }
        
        *expr = arg_end + 1; // Skip closing ')'
        
        // Call appropriate function
        if (strcmp(func_name, "SUM") == 0) {
            return func_sum(values, value_count);
        } else if (strcmp(func_name, "AVG") == 0) {
            return func_avg(values, value_count);
        } else if (strcmp(func_name, "MAX") == 0) {
            return func_max(values, value_count);
        } else if (strcmp(func_name, "MIN") == 0) {
            return func_min(values, value_count);
        } else if (strcmp(func_name, "MEDIAN") == 0) {
            return func_median(values, value_count);
        } else if (strcmp(func_name, "MODE") == 0) {
            return func_mode(values, value_count);
        }
        
    } else if (strcmp(func_name, "POWER") == 0) {
        // POWER function: POWER(base, exponent)
        skip_whitespace(expr);
        
        // Parse first argument (base)
        double base = parse_arithmetic_expression(sheet, expr, error);
        if (*error != ERROR_NONE) return 0.0;
        
        skip_whitespace(expr);
        if (**expr != ',') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        (*expr)++; // Skip comma
        
        // Parse second argument (exponent)  
        double exponent = parse_arithmetic_expression(sheet, expr, error);
        if (*error != ERROR_NONE) return 0.0;
        
        skip_whitespace(expr);
        if (**expr != ')') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        (*expr)++; // Skip closing parenthesis
        
        return func_power(base, exponent);
    } else if (strcmp(func_name, "IF") == 0) {
        // IF function: IF(condition, true_value, false_value)
        skip_whitespace(expr);
        
        // Parse condition
        double condition = evaluate_comparison(sheet, *expr, error);
        if (*error != ERROR_NONE) return 0.0;
        
        // Find the comma after condition
        int paren_depth = 0;
        while (**expr) {
            if (**expr == '(') paren_depth++;
            else if (**expr == ')') paren_depth--;
            else if (**expr == ',' && paren_depth == 0) break;
            (*expr)++;
        }
        
        if (**expr != ',') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        (*expr)++; // Skip comma
        
        // Parse true value (can be string or number)
        skip_whitespace(expr);
        double true_val = 0.0;
        char true_str[256] = {0};
        int has_true_str = 0;
        
        if (**expr == '"') {
            // String literal
            (*expr)++; // Skip opening quote
            int j = 0;
            while (**expr && **expr != '"' && j < 255) {
                if (**expr == '"' && *(*expr + 1) == '"') {
                    // Escaped quote
                    true_str[j++] = '"';
                    (*expr) += 2;
                } else {
                    true_str[j++] = **expr;
                    (*expr)++;
                }
            }
            if (**expr == '"') {
                (*expr)++; // Skip closing quote
                true_str[j] = '\0';
                has_true_str = 1;
            } else {
                *error = ERROR_PARSE;
                return 0.0;
            }
        } else {
            // Numeric expression - need to find the comma
            const char* true_start = *expr;
            paren_depth = 0;
            while (**expr) {
                if (**expr == '(') paren_depth++;
                else if (**expr == ')') paren_depth--;
                else if (**expr == ',' && paren_depth == 0) break;
                (*expr)++;
            }
            
            // Temporarily null-terminate and parse
            char true_expr[256];
            int true_len = (int)(*expr - true_start);
            if (true_len >= sizeof(true_expr)) {
                *error = ERROR_PARSE;
                return 0.0;
            }
            // strncpy_s already handles truncation safely
            strncpy_s(true_expr, sizeof(true_expr), true_start, (true_len < sizeof(true_expr) - 1) ? true_len : sizeof(true_expr) - 1);
            true_expr[(true_len < sizeof(true_expr) - 1) ? true_len : sizeof(true_expr) - 1] = '\0';
            
            const char* true_ptr = true_expr;
            true_val = parse_arithmetic_expression(sheet, &true_ptr, error);
            if (*error != ERROR_NONE) return 0.0;
        }
        
        // Expect comma
        skip_whitespace(expr);
        if (**expr != ',') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        (*expr)++; // Skip comma
        
        // Parse false value (can be string or number)
        skip_whitespace(expr);
        double false_val = 0.0;
        char false_str[256] = {0};
        int has_false_str = 0;
        
        if (**expr == '"') {
            // String literal
            (*expr)++; // Skip opening quote
            int j = 0;
            while (**expr && **expr != '"' && j < 255) {
                if (**expr == '"' && *(*expr + 1) == '"') {
                    // Escaped quote
                    false_str[j++] = '"';
                    (*expr) += 2;
                } else {
                    false_str[j++] = **expr;
                    (*expr)++;
                }
            }
            if (**expr == '"') {
                (*expr)++; // Skip closing quote
                false_str[j] = '\0';
                has_false_str = 1;
            } else {
                *error = ERROR_PARSE;
                return 0.0;
            }
        } else {
            // Numeric expression - need to find the closing paren
            const char* false_start = *expr;
            paren_depth = 0;
            while (**expr) {
                if (**expr == '(') paren_depth++;
                else if (**expr == ')') {
                    if (paren_depth == 0) break;
                    paren_depth--;
                }
                (*expr)++;
            }
            
            // Temporarily null-terminate and parse
            char false_expr[256];
            int false_len = (int)(*expr - false_start);
            if (false_len >= sizeof(false_expr)) {
                *error = ERROR_PARSE;
                return 0.0;
            }
            // strncpy_s already handles truncation safely
            strncpy_s(false_expr, sizeof(false_expr), false_start, (false_len < sizeof(false_expr) - 1) ? false_len : sizeof(false_expr) - 1);
            false_expr[(false_len < sizeof(false_expr) - 1) ? false_len : sizeof(false_expr) - 1] = '\0';
            
            const char* false_ptr = false_expr;
            false_val = parse_arithmetic_expression(sheet, &false_ptr, error);
            if (*error != ERROR_NONE) return 0.0;
        }
        
        // Expect closing parenthesis
        skip_whitespace(expr);
        if (**expr != ')') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        (*expr)++; // Skip closing parenthesis
        
        // Call IF function
        if (has_true_str || has_false_str) {
            return func_if_enhanced(condition, true_val, false_val,
                                   has_true_str ? true_str : NULL,
                                   has_false_str ? false_str : NULL);
        } else {
            return func_if(condition, true_val, false_val);
        }
    }
      *error = ERROR_PARSE;
    return 0.0;
}

// Cell color formatting functions
void cell_set_text_color(Cell* cell, int color) {
    if (cell) {
        cell->text_color = color;
    }
}

void cell_set_background_color(Cell* cell, int color) {
    if (cell) {
        cell->background_color = color;
    }
}

// Parse color name or hex value
int parse_color(const char* color_str) {
    if (!color_str) return -1;
    
    // Check for hex color (e.g., #FFFFFF)
    if (color_str[0] == '#' && strlen(color_str) == 7) {
        // For console applications, we'll map hex colors to the 16 console colors
        // This is a simplified mapping - you could make it more sophisticated
        unsigned int hex_value;
        if (sscanf_s(color_str + 1, "%x", &hex_value) == 1) {
            // Map hex to nearest console color (simplified)
            int r = (hex_value >> 16) & 0xFF;
            int g = (hex_value >> 8) & 0xFF;
            int b = hex_value & 0xFF;
            
            // Simple mapping to 16 colors
            if (r < 128 && g < 128 && b < 128) {
                if (r < 64 && g < 64 && b < 64) return COLOR_BLACK;
                if (b > r && b > g) return COLOR_BLUE;
                if (g > r && g > b) return COLOR_GREEN;
                if (r > g && r > b) return COLOR_RED;
                if (r > 128 && g > 128) return COLOR_YELLOW;
                if (r > 128 && b > 128) return COLOR_MAGENTA;
                if (g > 128 && b > 128) return COLOR_CYAN;
                return COLOR_WHITE;
            } else {
                // Bright colors
                if (b > r && b > g) return COLOR_BLUE | COLOR_BRIGHT;
                if (g > r && g > b) return COLOR_GREEN | COLOR_BRIGHT;
                if (r > g && r > b) return COLOR_RED | COLOR_BRIGHT;
                if (r > 200 && g > 200) return COLOR_YELLOW | COLOR_BRIGHT;
                if (r > 200 && b > 200) return COLOR_MAGENTA | COLOR_BRIGHT;
                if (g > 200 && b > 200) return COLOR_CYAN | COLOR_BRIGHT;
                return COLOR_WHITE | COLOR_BRIGHT;
            }
        }
        return -1; // Invalid hex
    }
    
    // Check for color names
    if (strcmp(color_str, "black") == 0) return COLOR_BLACK;
    if (strcmp(color_str, "blue") == 0) return COLOR_BLUE;
    if (strcmp(color_str, "green") == 0) return COLOR_GREEN;
    if (strcmp(color_str, "cyan") == 0) return COLOR_CYAN;
    if (strcmp(color_str, "red") == 0) return COLOR_RED;
    if (strcmp(color_str, "magenta") == 0) return COLOR_MAGENTA;
    if (strcmp(color_str, "yellow") == 0) return COLOR_YELLOW;
    if (strcmp(color_str, "white") == 0) return COLOR_WHITE;
    
    return -1; // Unknown color
}

// Column/Row resizing functions
void sheet_set_column_width(Sheet* sheet, int col, int width) {
    if (!sheet || col < 0 || col >= sheet->cols || width < 1) return;
    sheet->col_widths[col] = width;
}

void sheet_set_row_height(Sheet* sheet, int row, int height) {
    if (!sheet || row < 0 || row >= sheet->rows || height < 1) return;
    sheet->row_heights[row] = height;
}

int sheet_get_column_width(Sheet* sheet, int col) {
    if (!sheet || col < 0 || col >= sheet->cols) return DEFAULT_COLUMN_WIDTH;
    return sheet->col_widths[col];
}

int sheet_get_row_height(Sheet* sheet, int row) {
    if (!sheet || row < 0 || row >= sheet->rows) return 1; // Default height
    return sheet->row_heights[row];
}

void sheet_resize_columns_in_range(Sheet* sheet, int start_col, int end_col, int delta) {
    if (!sheet || start_col < 0 || end_col >= sheet->cols || start_col > end_col) return;
    
    for (int col = start_col; col <= end_col; col++) {
        int new_width = sheet->col_widths[col] + delta;
        if (new_width < MIN_COLUMN_WIDTH) new_width = MIN_COLUMN_WIDTH;
        if (new_width > MAX_COLUMN_WIDTH) new_width = MAX_COLUMN_WIDTH;
        sheet->col_widths[col] = new_width;
    }
}

void sheet_resize_rows_in_range(Sheet* sheet, int start_row, int end_row, int delta) {
    if (!sheet || start_row < 0 || end_row >= sheet->rows || start_row > end_row) return;
    
    for (int row = start_row; row <= end_row; row++) {
        int new_height = sheet->row_heights[row] + delta;
        if (new_height < MIN_ROW_HEIGHT) new_height = MIN_ROW_HEIGHT;
        if (new_height > MAX_ROW_HEIGHT) new_height = MAX_ROW_HEIGHT;
        sheet->row_heights[row] = new_height;
    }
}

// Insert/Delete Row and Column functions
void sheet_insert_row(Sheet* sheet, int row) {
    if (!sheet || row < 0 || row >= sheet->rows) return;
    
    // Shift all rows down from the insertion point
    for (int r = sheet->rows - 1; r > row; r--) {
        for (int c = 0; c < sheet->cols; c++) {
            if (sheet->cells[r-1][c]) {
                // Move cell from row r-1 to row r
                sheet->cells[r][c] = sheet->cells[r-1][c];
                sheet->cells[r][c]->row = r;
                sheet->cells[r-1][c] = NULL;
            }
        }
        // Move row height
        sheet->row_heights[r] = sheet->row_heights[r-1];
    }
    
    // Clear the inserted row and set default height
    for (int c = 0; c < sheet->cols; c++) {
        sheet->cells[row][c] = NULL;
    }
    sheet->row_heights[row] = 1; // Default height
    
    // Mark sheet for recalculation
    sheet->needs_recalc = 1;
}

void sheet_insert_column(Sheet* sheet, int col) {
    if (!sheet || col < 0 || col >= sheet->cols) return;
    
    // Shift all columns right from the insertion point
    for (int c = sheet->cols - 1; c > col; c--) {
        for (int r = 0; r < sheet->rows; r++) {
            if (sheet->cells[r][c-1]) {
                // Move cell from column c-1 to column c
                sheet->cells[r][c] = sheet->cells[r][c-1];
                sheet->cells[r][c]->col = c;
                sheet->cells[r][c-1] = NULL;
            }
        }
        // Move column width
        sheet->col_widths[c] = sheet->col_widths[c-1];
    }
    
    // Clear the inserted column and set default width
    for (int r = 0; r < sheet->rows; r++) {
        sheet->cells[r][col] = NULL;
    }
    sheet->col_widths[col] = DEFAULT_COLUMN_WIDTH;
    
    // Mark sheet for recalculation
    sheet->needs_recalc = 1;
}

void sheet_delete_row(Sheet* sheet, int row) {
    if (!sheet || row < 0 || row >= sheet->rows) return;
    
    // Clear and free all cells in the row to be deleted
    for (int c = 0; c < sheet->cols; c++) {
        if (sheet->cells[row][c]) {
            cell_free(sheet->cells[row][c]);
            sheet->cells[row][c] = NULL;
        }
    }
    
    // Shift all rows up from the deletion point
    for (int r = row; r < sheet->rows - 1; r++) {
        for (int c = 0; c < sheet->cols; c++) {
            if (sheet->cells[r+1][c]) {
                // Move cell from row r+1 to row r
                sheet->cells[r][c] = sheet->cells[r+1][c];
                sheet->cells[r][c]->row = r;
                sheet->cells[r+1][c] = NULL;
            }
        }
        // Move row height
        sheet->row_heights[r] = sheet->row_heights[r+1];
    }
    
    // Clear the last row and set default height
    for (int c = 0; c < sheet->cols; c++) {
        sheet->cells[sheet->rows - 1][c] = NULL;
    }
    sheet->row_heights[sheet->rows - 1] = 1; // Default height
    
    // Mark sheet for recalculation
    sheet->needs_recalc = 1;
}

void sheet_delete_column(Sheet* sheet, int col) {
    if (!sheet || col < 0 || col >= sheet->cols) return;
    
    // Clear and free all cells in the column to be deleted
    for (int r = 0; r < sheet->rows; r++) {
        if (sheet->cells[r][col]) {
            cell_free(sheet->cells[r][col]);
            sheet->cells[r][col] = NULL;
        }
    }
    
    // Shift all columns left from the deletion point
    for (int c = col; c < sheet->cols - 1; c++) {
        for (int r = 0; r < sheet->rows; r++) {
            if (sheet->cells[r][c+1]) {
                // Move cell from column c+1 to column c
                sheet->cells[r][c] = sheet->cells[r][c+1];
                sheet->cells[r][c]->col = c;
                sheet->cells[r][c+1] = NULL;
            }
        }
        // Move column width
        sheet->col_widths[c] = sheet->col_widths[c+1];
    }
    
    // Clear the last column and set default width
    for (int r = 0; r < sheet->rows; r++) {
        sheet->cells[r][sheet->cols - 1] = NULL;
    }
    sheet->col_widths[sheet->cols - 1] = DEFAULT_COLUMN_WIDTH;
    
    // Mark sheet for recalculation
    sheet->needs_recalc = 1;
}


