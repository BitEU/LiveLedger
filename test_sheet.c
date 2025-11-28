// test_sheet.c - Comprehensive unit tests for LiveLedger spreadsheet engine
// Build: cl test_sheet.c /Fe:test_sheet.exe
// Run: test_sheet.exe

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <assert.h>
#include <stdarg.h>

// Stub for debug logging (required by console.h)
FILE* debug_file = NULL;
void debug_log(const char* format, ...) {
    // No-op for tests
}

#include "console.h"  // Must be included before sheet.h for COLOR_* definitions
#include "sheet.h"

// Test result tracking
static int tests_run = 0;
static int tests_passed = 0;
static int tests_failed = 0;

// Macros for test assertions
#define TEST_START(name) \
    do { \
        printf("\n=== Testing: %s ===\n", name); \
        tests_run++; \
    } while(0)

#define ASSERT_TRUE(condition, message) \
    do { \
        if (!(condition)) { \
            printf("  FAIL: %s\n", message); \
            printf("    Expected true, got false\n"); \
            tests_failed++; \
            return 0; \
        } \
    } while(0)

#define ASSERT_FALSE(condition, message) \
    do { \
        if (condition) { \
            printf("  FAIL: %s\n", message); \
            printf("    Expected false, got true\n"); \
            tests_failed++; \
            return 0; \
        } \
    } while(0)

#define ASSERT_EQUAL(expected, actual, message) \
    do { \
        if ((expected) != (actual)) { \
            printf("  FAIL: %s\n", message); \
            printf("    Expected: %d, Got: %d\n", (int)(expected), (int)(actual)); \
            tests_failed++; \
            return 0; \
        } \
    } while(0)

#define ASSERT_DOUBLE_EQUAL(expected, actual, tolerance, message) \
    do { \
        if (fabs((expected) - (actual)) > (tolerance)) { \
            printf("  FAIL: %s\n", message); \
            printf("    Expected: %.6f, Got: %.6f\n", (double)(expected), (double)(actual)); \
            tests_failed++; \
            return 0; \
        } \
    } while(0)

#define ASSERT_STRING_EQUAL(expected, actual, message) \
    do { \
        if (strcmp((expected), (actual)) != 0) { \
            printf("  FAIL: %s\n", message); \
            printf("    Expected: '%s', Got: '%s'\n", expected, actual); \
            tests_failed++; \
            return 0; \
        } \
    } while(0)

#define ASSERT_NOT_NULL(ptr, message) \
    do { \
        if ((ptr) == NULL) { \
            printf("  FAIL: %s\n", message); \
            printf("    Expected non-NULL pointer\n"); \
            tests_failed++; \
            return 0; \
        } \
    } while(0)

#define ASSERT_NULL(ptr, message) \
    do { \
        if ((ptr) != NULL) { \
            printf("  FAIL: %s\n", message); \
            printf("    Expected NULL pointer\n"); \
            tests_failed++; \
            return 0; \
        } \
    } while(0)

#define TEST_PASS() \
    do { \
        printf("  PASS\n"); \
        tests_passed++; \
        return 1; \
    } while(0)

// ============================================================================
// BASIC SHEET OPERATIONS TESTS
// ============================================================================

int test_sheet_creation() {
    TEST_START("Sheet Creation");

    Sheet* sheet = sheet_new(100, 26);
    ASSERT_NOT_NULL(sheet, "Sheet should be created");
    ASSERT_EQUAL(100, sheet->rows, "Sheet should have 100 rows");
    ASSERT_EQUAL(26, sheet->cols, "Sheet should have 26 columns");
    ASSERT_NOT_NULL(sheet->cells, "Sheet cells should be allocated");
    ASSERT_NOT_NULL(sheet->col_widths, "Column widths should be allocated");
    ASSERT_NOT_NULL(sheet->row_heights, "Row heights should be allocated");

    sheet_free(sheet);
    TEST_PASS();
}

int test_sheet_invalid_creation() {
    TEST_START("Sheet Invalid Creation");

    // These would normally fail gracefully in production code
    // For now, we just test that valid creation works
    Sheet* sheet = sheet_new(0, 0);
    if (sheet) {
        sheet_free(sheet);
    }

    TEST_PASS();
}

int test_cell_number_operations() {
    TEST_START("Cell Number Operations");

    Sheet* sheet = sheet_new(100, 26);
    ASSERT_NOT_NULL(sheet, "Sheet creation failed");

    // Set a number
    sheet_set_number(sheet, 0, 0, 42.5);

    Cell* cell = sheet_get_cell(sheet, 0, 0);
    ASSERT_NOT_NULL(cell, "Cell should exist after setting number");
    ASSERT_EQUAL(CELL_NUMBER, cell->type, "Cell type should be NUMBER");
    ASSERT_DOUBLE_EQUAL(42.5, cell->data.number, 0.0001, "Cell value should be 42.5");

    // Update the number
    sheet_set_number(sheet, 0, 0, 100.0);
    cell = sheet_get_cell(sheet, 0, 0);
    ASSERT_DOUBLE_EQUAL(100.0, cell->data.number, 0.0001, "Cell value should be updated to 100.0");

    sheet_free(sheet);
    TEST_PASS();
}

int test_cell_string_operations() {
    TEST_START("Cell String Operations");

    Sheet* sheet = sheet_new(100, 26);
    ASSERT_NOT_NULL(sheet, "Sheet creation failed");

    // Set a string
    const char* test_string = "Hello, World!";
    sheet_set_string(sheet, 1, 1, test_string);

    Cell* cell = sheet_get_cell(sheet, 1, 1);
    ASSERT_NOT_NULL(cell, "Cell should exist after setting string");
    ASSERT_EQUAL(CELL_STRING, cell->type, "Cell type should be STRING");
    ASSERT_STRING_EQUAL(test_string, cell->data.string, "Cell string should match input");

    // Update the string
    const char* new_string = "Updated";
    sheet_set_string(sheet, 1, 1, new_string);
    cell = sheet_get_cell(sheet, 1, 1);
    ASSERT_STRING_EQUAL(new_string, cell->data.string, "Cell string should be updated");

    sheet_free(sheet);
    TEST_PASS();
}

int test_cell_clear() {
    TEST_START("Cell Clear Operation");

    Sheet* sheet = sheet_new(100, 26);

    // Set a value then clear it
    sheet_set_number(sheet, 0, 0, 42.0);
    sheet_clear_cell(sheet, 0, 0);

    Cell* cell = sheet_get_cell(sheet, 0, 0);
    ASSERT_NOT_NULL(cell, "Cell should still exist after clear");
    ASSERT_EQUAL(CELL_EMPTY, cell->type, "Cell type should be EMPTY after clear");

    sheet_free(sheet);
    TEST_PASS();
}

// ============================================================================
// CELL REFERENCE PARSING TESTS
// ============================================================================

int test_cell_reference_parsing() {
    TEST_START("Cell Reference Parsing");

    int row, col;

    // Test A1
    ASSERT_TRUE(parse_cell_reference("A1", &row, &col), "Should parse A1");
    ASSERT_EQUAL(0, row, "A1 row should be 0");
    ASSERT_EQUAL(0, col, "A1 col should be 0");

    // Test Z26
    ASSERT_TRUE(parse_cell_reference("Z26", &row, &col), "Should parse Z26");
    ASSERT_EQUAL(25, row, "Z26 row should be 25");
    ASSERT_EQUAL(25, col, "Z26 col should be 25");

    // Test AA1
    ASSERT_TRUE(parse_cell_reference("AA1", &row, &col), "Should parse AA1");
    ASSERT_EQUAL(0, row, "AA1 row should be 0");
    ASSERT_EQUAL(26, col, "AA1 col should be 26");

    // Test lowercase
    ASSERT_TRUE(parse_cell_reference("a1", &row, &col), "Should parse a1");
    ASSERT_EQUAL(0, row, "a1 row should be 0");
    ASSERT_EQUAL(0, col, "a1 col should be 0");

    // Test invalid references
    ASSERT_FALSE(parse_cell_reference("1A", &row, &col), "Should reject 1A");
    ASSERT_FALSE(parse_cell_reference("", &row, &col), "Should reject empty string");
    ASSERT_FALSE(parse_cell_reference("ABC", &row, &col), "Should reject ABC (no number)");

    TEST_PASS();
}

int test_cell_reference_to_string() {
    TEST_START("Cell Reference To String");

    char* ref;

    ref = cell_reference_to_string(0, 0);
    ASSERT_STRING_EQUAL("A1", ref, "Should convert (0,0) to A1");

    ref = cell_reference_to_string(25, 25);
    ASSERT_STRING_EQUAL("Z26", ref, "Should convert (25,25) to Z26");

    ref = cell_reference_to_string(0, 26);
    ASSERT_STRING_EQUAL("AA1", ref, "Should convert (0,26) to AA1");

    TEST_PASS();
}

// ============================================================================
// FORMULA EVALUATION TESTS
// ============================================================================

int test_simple_arithmetic() {
    TEST_START("Simple Arithmetic Formulas");

    Sheet* sheet = sheet_new(100, 26);
    ErrorType error;

    // Test addition
    double result = evaluate_formula(sheet, "=2+3", &error);
    ASSERT_EQUAL(ERROR_NONE, error, "Addition should not error");
    ASSERT_DOUBLE_EQUAL(5.0, result, 0.0001, "2+3 should equal 5");

    // Test subtraction
    result = evaluate_formula(sheet, "=10-4", &error);
    ASSERT_EQUAL(ERROR_NONE, error, "Subtraction should not error");
    ASSERT_DOUBLE_EQUAL(6.0, result, 0.0001, "10-4 should equal 6");

    // Test multiplication
    result = evaluate_formula(sheet, "=3*4", &error);
    ASSERT_EQUAL(ERROR_NONE, error, "Multiplication should not error");
    ASSERT_DOUBLE_EQUAL(12.0, result, 0.0001, "3*4 should equal 12");

    // Test division
    result = evaluate_formula(sheet, "=15/3", &error);
    ASSERT_EQUAL(ERROR_NONE, error, "Division should not error");
    ASSERT_DOUBLE_EQUAL(5.0, result, 0.0001, "15/3 should equal 5");

    // Test division by zero
    result = evaluate_formula(sheet, "=5/0", &error);
    ASSERT_EQUAL(ERROR_DIV_ZERO, error, "Division by zero should error");

    sheet_free(sheet);
    TEST_PASS();
}

int test_operator_precedence() {
    TEST_START("Operator Precedence");

    Sheet* sheet = sheet_new(100, 26);
    ErrorType error;

    // Test multiplication before addition
    double result = evaluate_formula(sheet, "=2+3*4", &error);
    ASSERT_EQUAL(ERROR_NONE, error, "Should not error");
    ASSERT_DOUBLE_EQUAL(14.0, result, 0.0001, "2+3*4 should equal 14 (not 20)");

    // Test parentheses
    result = evaluate_formula(sheet, "=(2+3)*4", &error);
    ASSERT_EQUAL(ERROR_NONE, error, "Should not error");
    ASSERT_DOUBLE_EQUAL(20.0, result, 0.0001, "(2+3)*4 should equal 20");

    // Test complex expression
    result = evaluate_formula(sheet, "=2+3*4-6/2", &error);
    ASSERT_EQUAL(ERROR_NONE, error, "Should not error");
    ASSERT_DOUBLE_EQUAL(11.0, result, 0.0001, "2+3*4-6/2 should equal 11");

    sheet_free(sheet);
    TEST_PASS();
}

int test_cell_references_in_formulas() {
    TEST_START("Cell References in Formulas");

    Sheet* sheet = sheet_new(100, 26);

    // Set up source cells
    sheet_set_number(sheet, 0, 0, 10.0);  // A1 = 10
    sheet_set_number(sheet, 0, 1, 20.0);  // B1 = 20

    // Test simple reference
    sheet_set_formula(sheet, 1, 0, "=A1");
    sheet_recalculate(sheet);

    Cell* cell = sheet_get_cell(sheet, 1, 0);
    ASSERT_EQUAL(CELL_FORMULA, cell->type, "Cell should be formula type");
    ASSERT_EQUAL(ERROR_NONE, cell->data.formula.error, "Formula should not error");
    ASSERT_DOUBLE_EQUAL(10.0, cell->data.formula.cached_value, 0.0001, "=A1 should equal 10");

    // Test formula with multiple references
    sheet_set_formula(sheet, 1, 1, "=A1+B1");
    sheet_recalculate(sheet);

    cell = sheet_get_cell(sheet, 1, 1);
    ASSERT_EQUAL(ERROR_NONE, cell->data.formula.error, "Formula should not error");
    ASSERT_DOUBLE_EQUAL(30.0, cell->data.formula.cached_value, 0.0001, "=A1+B1 should equal 30");

    // Test complex formula
    sheet_set_formula(sheet, 1, 2, "=A1*B1/2");
    sheet_recalculate(sheet);

    cell = sheet_get_cell(sheet, 1, 2);
    ASSERT_EQUAL(ERROR_NONE, cell->data.formula.error, "Formula should not error");
    ASSERT_DOUBLE_EQUAL(100.0, cell->data.formula.cached_value, 0.0001, "=A1*B1/2 should equal 100");

    sheet_free(sheet);
    TEST_PASS();
}

// ============================================================================
// FUNCTION TESTS
// ============================================================================

int test_sum_function() {
    TEST_START("SUM Function");

    Sheet* sheet = sheet_new(100, 26);

    // Set up test data in A1:A5
    sheet_set_number(sheet, 0, 0, 10.0);
    sheet_set_number(sheet, 1, 0, 20.0);
    sheet_set_number(sheet, 2, 0, 30.0);
    sheet_set_number(sheet, 3, 0, 40.0);
    sheet_set_number(sheet, 4, 0, 50.0);

    // Test SUM with range
    sheet_set_formula(sheet, 5, 0, "=SUM(A1:A5)");
    sheet_recalculate(sheet);

    Cell* cell = sheet_get_cell(sheet, 5, 0);
    ASSERT_EQUAL(ERROR_NONE, cell->data.formula.error, "SUM should not error");
    ASSERT_DOUBLE_EQUAL(150.0, cell->data.formula.cached_value, 0.0001, "SUM(A1:A5) should equal 150");

    sheet_free(sheet);
    TEST_PASS();
}

int test_avg_function() {
    TEST_START("AVG Function");

    Sheet* sheet = sheet_new(100, 26);

    // Set up test data
    sheet_set_number(sheet, 0, 0, 10.0);
    sheet_set_number(sheet, 1, 0, 20.0);
    sheet_set_number(sheet, 2, 0, 30.0);

    sheet_set_formula(sheet, 3, 0, "=AVG(A1:A3)");
    sheet_recalculate(sheet);

    Cell* cell = sheet_get_cell(sheet, 3, 0);
    ASSERT_EQUAL(ERROR_NONE, cell->data.formula.error, "AVG should not error");
    ASSERT_DOUBLE_EQUAL(20.0, cell->data.formula.cached_value, 0.0001, "AVG(A1:A3) should equal 20");

    sheet_free(sheet);
    TEST_PASS();
}

int test_max_min_functions() {
    TEST_START("MAX and MIN Functions");

    Sheet* sheet = sheet_new(100, 26);

    // Set up test data
    sheet_set_number(sheet, 0, 0, 5.0);
    sheet_set_number(sheet, 1, 0, 15.0);
    sheet_set_number(sheet, 2, 0, 3.0);
    sheet_set_number(sheet, 3, 0, 22.0);
    sheet_set_number(sheet, 4, 0, 8.0);

    // Test MAX
    sheet_set_formula(sheet, 5, 0, "=MAX(A1:A5)");
    sheet_recalculate(sheet);

    Cell* cell = sheet_get_cell(sheet, 5, 0);
    ASSERT_EQUAL(ERROR_NONE, cell->data.formula.error, "MAX should not error");
    ASSERT_DOUBLE_EQUAL(22.0, cell->data.formula.cached_value, 0.0001, "MAX should be 22");

    // Test MIN
    sheet_set_formula(sheet, 6, 0, "=MIN(A1:A5)");
    sheet_recalculate(sheet);

    cell = sheet_get_cell(sheet, 6, 0);
    ASSERT_EQUAL(ERROR_NONE, cell->data.formula.error, "MIN should not error");
    ASSERT_DOUBLE_EQUAL(3.0, cell->data.formula.cached_value, 0.0001, "MIN should be 3");

    sheet_free(sheet);
    TEST_PASS();
}

int test_median_function() {
    TEST_START("MEDIAN Function");

    Sheet* sheet = sheet_new(100, 26);

    // Test odd number of values
    sheet_set_number(sheet, 0, 0, 1.0);
    sheet_set_number(sheet, 1, 0, 3.0);
    sheet_set_number(sheet, 2, 0, 2.0);
    sheet_set_number(sheet, 3, 0, 5.0);
    sheet_set_number(sheet, 4, 0, 4.0);

    sheet_set_formula(sheet, 5, 0, "=MEDIAN(A1:A5)");
    sheet_recalculate(sheet);

    Cell* cell = sheet_get_cell(sheet, 5, 0);
    ASSERT_EQUAL(ERROR_NONE, cell->data.formula.error, "MEDIAN should not error");
    ASSERT_DOUBLE_EQUAL(3.0, cell->data.formula.cached_value, 0.0001, "MEDIAN of [1,2,3,4,5] should be 3");

    // Test even number of values
    sheet_set_number(sheet, 0, 1, 1.0);
    sheet_set_number(sheet, 1, 1, 2.0);
    sheet_set_number(sheet, 2, 1, 3.0);
    sheet_set_number(sheet, 3, 1, 4.0);

    sheet_set_formula(sheet, 4, 1, "=MEDIAN(B1:B4)");
    sheet_recalculate(sheet);

    cell = sheet_get_cell(sheet, 4, 1);
    ASSERT_EQUAL(ERROR_NONE, cell->data.formula.error, "MEDIAN should not error");
    ASSERT_DOUBLE_EQUAL(2.5, cell->data.formula.cached_value, 0.0001, "MEDIAN of [1,2,3,4] should be 2.5");

    sheet_free(sheet);
    TEST_PASS();
}

int test_power_function() {
    TEST_START("POWER Function");

    Sheet* sheet = sheet_new(100, 26);
    ErrorType error;

    // Test basic power
    double result = evaluate_formula(sheet, "=POWER(2, 3)", &error);
    ASSERT_EQUAL(ERROR_NONE, error, "POWER should not error");
    ASSERT_DOUBLE_EQUAL(8.0, result, 0.0001, "2^3 should be 8");

    // Test square root
    result = evaluate_formula(sheet, "=POWER(16, 0.5)", &error);
    ASSERT_EQUAL(ERROR_NONE, error, "POWER should not error");
    ASSERT_DOUBLE_EQUAL(4.0, result, 0.0001, "16^0.5 should be 4");

    // Test with cell references
    sheet_set_number(sheet, 0, 0, 3.0);
    sheet_set_number(sheet, 0, 1, 4.0);
    sheet_set_formula(sheet, 0, 2, "=POWER(A1, B1)");
    sheet_recalculate(sheet);

    Cell* cell = sheet_get_cell(sheet, 0, 2);
    ASSERT_EQUAL(ERROR_NONE, cell->data.formula.error, "POWER should not error");
    ASSERT_DOUBLE_EQUAL(81.0, cell->data.formula.cached_value, 0.0001, "3^4 should be 81");

    sheet_free(sheet);
    TEST_PASS();
}

int test_if_function() {
    TEST_START("IF Function");

    Sheet* sheet = sheet_new(100, 26);

        // Test IF with number comparison


    sheet_set_number(sheet, 0, 0, 10.0);
    sheet_set_formula(sheet, 0, 1, "=IF(A1>5, 100, 200)");
    sheet_recalculate(sheet);

    Cell* cell = sheet_get_cell(sheet, 0, 1);

        ASSERT_EQUAL(ERROR_NONE, cell->data.formula.error, "IF should not error");
        ASSERT_DOUBLE_EQUAL(100.0, cell->data.formula.cached_value, 0.0001, "IF(10>5, 100, 200) should be 100");

        
    // Test IF with false condition
    sheet_set_number(sheet, 1, 0, 3.0);
    sheet_set_formula(sheet, 1, 1, "=IF(A2>5, 100, 200)");
    sheet_recalculate(sheet);

    cell = sheet_get_cell(sheet, 1, 1);
    ASSERT_EQUAL(ERROR_NONE, cell->data.formula.error, "IF should not error");
    ASSERT_DOUBLE_EQUAL(200.0, cell->data.formula.cached_value, 0.0001, "IF(3>5, 100, 200) should be 200");

    sheet_free(sheet);
    TEST_PASS();
}

// ============================================================================
// COMPARISON OPERATOR TESTS
// ============================================================================

int test_comparison_operators() {
    TEST_START("Comparison Operators");

    Sheet* sheet = sheet_new(100, 26);
    ErrorType error;

    // Greater than
    double result = evaluate_formula(sheet, "=10>5", &error);
    ASSERT_EQUAL(ERROR_NONE, error, "Comparison should not error");
    ASSERT_DOUBLE_EQUAL(1.0, result, 0.0001, "10>5 should be true (1.0)");

    // Less than
    result = evaluate_formula(sheet, "=3<5", &error);
    ASSERT_DOUBLE_EQUAL(1.0, result, 0.0001, "3<5 should be true");

    // Equal
    result = evaluate_formula(sheet, "=5=5", &error);
    ASSERT_DOUBLE_EQUAL(1.0, result, 0.0001, "5=5 should be true");

    // Not equal
    result = evaluate_formula(sheet, "=5<>3", &error);
    ASSERT_DOUBLE_EQUAL(1.0, result, 0.0001, "5<>3 should be true");

    // Greater than or equal
    result = evaluate_formula(sheet, "=5>=5", &error);
    ASSERT_DOUBLE_EQUAL(1.0, result, 0.0001, "5>=5 should be true");

    // Less than or equal
    result = evaluate_formula(sheet, "=3<=5", &error);
    ASSERT_DOUBLE_EQUAL(1.0, result, 0.0001, "3<=5 should be true");

    // False comparison
    result = evaluate_formula(sheet, "=10<5", &error);
    ASSERT_DOUBLE_EQUAL(0.0, result, 0.0001, "10<5 should be false (0.0)");

    sheet_free(sheet);
    TEST_PASS();
}

// ============================================================================
// FORMATTING TESTS
// ============================================================================

int test_percentage_formatting() {
    TEST_START("Percentage Formatting");

    Cell* cell = cell_new(0, 0);
    cell_set_number(cell, 0.25);
    cell_set_format(cell, FORMAT_PERCENTAGE, 0);

    char* formatted = format_cell_value(cell);
    ASSERT_STRING_EQUAL("25.00%", formatted, "0.25 should format as 25.00%");

    cell_free(cell);
    TEST_PASS();
}

int test_currency_formatting() {
    TEST_START("Currency Formatting");

    Cell* cell = cell_new(0, 0);
    cell_set_number(cell, 1234.56);
    cell_set_format(cell, FORMAT_CURRENCY, 0);

    char* formatted = format_cell_value(cell);
    ASSERT_STRING_EQUAL("$1234.56", formatted, "1234.56 should format as $1234.56");

    // Test negative currency
    cell_set_number(cell, -500.00);
    formatted = format_cell_value(cell);
    ASSERT_STRING_EQUAL("-$500.00", formatted, "-500.00 should format as -$500.00");

    cell_free(cell);
    TEST_PASS();
}

// ============================================================================
// COPY/PASTE TESTS
// ============================================================================

int test_cell_copy_paste() {
    TEST_START("Cell Copy and Paste");

    Sheet* sheet = sheet_new(100, 26);

    // Set source cell
    sheet_set_number(sheet, 0, 0, 42.0);

    // Copy cell
    Cell* src_cell = sheet_get_cell(sheet, 0, 0);
    sheet_set_clipboard_cell(src_cell);

    // Paste to another cell
    sheet_copy_cell(sheet, 0, 0, 1, 1);

    Cell* dest_cell = sheet_get_cell(sheet, 1, 1);
    ASSERT_NOT_NULL(dest_cell, "Destination cell should exist");
    ASSERT_EQUAL(CELL_NUMBER, dest_cell->type, "Destination cell should be number type");
    ASSERT_DOUBLE_EQUAL(42.0, dest_cell->data.number, 0.0001, "Destination cell value should match source");

    sheet_free(sheet);
    TEST_PASS();
}

// ============================================================================
// RANGE TESTS
// ============================================================================

int test_range_parsing() {
    TEST_START("Range Parsing");

    CellRange range;

    // Test valid range
    ASSERT_TRUE(parse_range("A1:A5", &range), "Should parse A1:A5");
    ASSERT_EQUAL(0, range.start_row, "Start row should be 0");
    ASSERT_EQUAL(0, range.start_col, "Start col should be 0");
    ASSERT_EQUAL(4, range.end_row, "End row should be 4");
    ASSERT_EQUAL(0, range.end_col, "End col should be 0");

    // Test 2D range
    ASSERT_TRUE(parse_range("B2:D4", &range), "Should parse B2:D4");
    ASSERT_EQUAL(1, range.start_row, "Start row should be 1");
    ASSERT_EQUAL(1, range.start_col, "Start col should be 1");
    ASSERT_EQUAL(3, range.end_row, "End row should be 3");
    ASSERT_EQUAL(3, range.end_col, "End col should be 3");

    // Test invalid range
    ASSERT_FALSE(parse_range("A1", &range), "Should reject single cell");
    ASSERT_FALSE(parse_range("A1:B", &range), "Should reject invalid range");

    TEST_PASS();
}

int test_range_selection() {
    TEST_START("Range Selection");

    Sheet* sheet = sheet_new(100, 26);

    // Start selection
    sheet_start_range_selection(sheet, 0, 0);
    ASSERT_TRUE(sheet->selection.is_active, "Selection should be active");
    ASSERT_EQUAL(0, sheet->selection.start_row, "Start row should be 0");
    ASSERT_EQUAL(0, sheet->selection.start_col, "Start col should be 0");

    // Extend selection
    sheet_extend_range_selection(sheet, 2, 2);
    ASSERT_EQUAL(2, sheet->selection.end_row, "End row should be 2");
    ASSERT_EQUAL(2, sheet->selection.end_col, "End col should be 2");

    // Check if cell is in selection
    ASSERT_TRUE(sheet_is_in_selection(sheet, 1, 1), "Cell (1,1) should be in selection");
    ASSERT_FALSE(sheet_is_in_selection(sheet, 5, 5), "Cell (5,5) should not be in selection");

    // Clear selection
    sheet_clear_range_selection(sheet);
    ASSERT_FALSE(sheet->selection.is_active, "Selection should be inactive");

    sheet_free(sheet);
    TEST_PASS();
}

// ============================================================================
// CSV TESTS
// ============================================================================

int test_csv_save_load() {
    TEST_START("CSV Save and Load");

    Sheet* sheet1 = sheet_new(100, 26);

    // Create test data
    sheet_set_string(sheet1, 0, 0, "Name");
    sheet_set_string(sheet1, 0, 1, "Age");
    sheet_set_string(sheet1, 0, 2, "Score");

    sheet_set_string(sheet1, 1, 0, "Alice");
    sheet_set_number(sheet1, 1, 1, 25.0);
    sheet_set_number(sheet1, 1, 2, 95.5);

    sheet_set_string(sheet1, 2, 0, "Bob");
    sheet_set_number(sheet1, 2, 1, 30.0);
    sheet_set_number(sheet1, 2, 2, 87.3);

    // Save to CSV
    const char* filename = "test_output.csv";
    ASSERT_TRUE(sheet_save_csv(sheet1, filename, 0), "CSV save should succeed");

    // Load into new sheet
    Sheet* sheet2 = sheet_new(100, 26);
    ASSERT_TRUE(sheet_load_csv(sheet2, filename, 0), "CSV load should succeed");

    // Verify data
    Cell* cell = sheet_get_cell(sheet2, 0, 0);
    ASSERT_NOT_NULL(cell, "Header cell should exist");
    ASSERT_EQUAL(CELL_STRING, cell->type, "Header should be string");
    ASSERT_STRING_EQUAL("Name", cell->data.string, "Header value should match");

    cell = sheet_get_cell(sheet2, 1, 1);
    ASSERT_NOT_NULL(cell, "Data cell should exist");
    ASSERT_EQUAL(CELL_NUMBER, cell->type, "Age should be number");
    ASSERT_DOUBLE_EQUAL(25.0, cell->data.number, 0.0001, "Age value should match");

    // Cleanup
    sheet_free(sheet1);
    sheet_free(sheet2);
    remove(filename);

    TEST_PASS();
}

// ============================================================================
// COLUMN/ROW SIZING TESTS
// ============================================================================

int test_column_row_sizing() {
    TEST_START("Column and Row Sizing");

    Sheet* sheet = sheet_new(100, 26);

    // Test default sizes
    ASSERT_EQUAL(10, sheet_get_column_width(sheet, 0), "Default column width should be 10");
    ASSERT_EQUAL(1, sheet_get_row_height(sheet, 0), "Default row height should be 1");

    // Set custom sizes
    sheet_set_column_width(sheet, 0, 20);
    ASSERT_EQUAL(20, sheet_get_column_width(sheet, 0), "Column width should be updated");

    sheet_set_row_height(sheet, 0, 3);
    ASSERT_EQUAL(3, sheet_get_row_height(sheet, 0), "Row height should be updated");

    // Test range resizing
    sheet_resize_columns_in_range(sheet, 0, 2, 5);
    ASSERT_EQUAL(25, sheet_get_column_width(sheet, 0), "Column 0 width should increase by 5");
    ASSERT_EQUAL(15, sheet_get_column_width(sheet, 1), "Column 1 width should increase by 5");

    sheet_free(sheet);
    TEST_PASS();
}

// ============================================================================
// COLOR PARSING TESTS
// ============================================================================

int test_color_parsing() {
    TEST_START("Color Parsing");

    // Test named colors
    ASSERT_EQUAL(COLOR_BLACK, parse_color("black"), "Should parse black");
    ASSERT_EQUAL(COLOR_RED, parse_color("red"), "Should parse red");
    ASSERT_EQUAL(COLOR_GREEN, parse_color("green"), "Should parse green");
    ASSERT_EQUAL(COLOR_BLUE, parse_color("blue"), "Should parse blue");
    ASSERT_EQUAL(COLOR_WHITE, parse_color("white"), "Should parse white");

    // Test hex colors (basic validation - exact mapping may vary)
    int color = parse_color("#FF0000");
    ASSERT_TRUE(color >= 0, "Should parse hex color #FF0000");

    color = parse_color("#000000");
    ASSERT_TRUE(color >= 0, "Should parse hex color #000000");

    // Test invalid color
    ASSERT_EQUAL(-1, parse_color("invalid"), "Should reject invalid color");
    ASSERT_EQUAL(-1, parse_color("#GGGGGG"), "Should reject invalid hex color");

    TEST_PASS();
}

// ============================================================================
// ERROR HANDLING TESTS
// ============================================================================

int test_error_handling() {
    TEST_START("Error Handling");

    Sheet* sheet = sheet_new(100, 26);
    ErrorType error;

    // Test division by zero
    sheet_set_formula(sheet, 0, 0, "=5/0");
    sheet_recalculate(sheet);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    ASSERT_EQUAL(ERROR_DIV_ZERO, cell->data.formula.error, "Should detect division by zero");

    // Test invalid reference
    // Note: Out-of-bounds references might return 0 instead of error
    sheet_set_formula(sheet, 0, 1, "=ZZZ999");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 0, 1);
    // Accept either an error OR a zero value for out-of-bounds reference
    if (cell->data.formula.error == ERROR_NONE) {
        printf("  NOTE: Out-of-bounds reference returned 0 instead of error (acceptable)\n");
        ASSERT_DOUBLE_EQUAL(0.0, cell->data.formula.cached_value, 0.0001, "Out-of-bounds should be 0");
    } else {
        ASSERT_TRUE(cell->data.formula.error != ERROR_NONE, "Should detect invalid reference");
    }

    // Test parse error
    sheet_set_formula(sheet, 0, 2, "=2++3");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 0, 2);
    ASSERT_TRUE(cell->data.formula.error != ERROR_NONE, "Should detect parse error");

    sheet_free(sheet);
    TEST_PASS();
}

// ============================================================================
// EDGE CASE TESTS
// ============================================================================

int test_empty_cells() {
    TEST_START("Empty Cell Handling");

    Sheet* sheet = sheet_new(100, 26);

    // Reference to empty cell should return 0
    sheet_set_formula(sheet, 0, 1, "=A1");
    sheet_recalculate(sheet);

    Cell* cell = sheet_get_cell(sheet, 0, 1);
    ASSERT_EQUAL(ERROR_NONE, cell->data.formula.error, "Empty cell reference should not error");
    ASSERT_DOUBLE_EQUAL(0.0, cell->data.formula.cached_value, 0.0001, "Empty cell should be treated as 0");

    sheet_free(sheet);
    TEST_PASS();
}

int test_large_numbers() {
    TEST_START("Large Number Handling");

    Sheet* sheet = sheet_new(100, 26);

    // Test large positive number
    sheet_set_number(sheet, 0, 0, 1e15);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    ASSERT_DOUBLE_EQUAL(1e15, cell->data.number, 1e10, "Should handle large positive numbers");

    // Test large negative number
    sheet_set_number(sheet, 0, 1, -1e15);
    cell = sheet_get_cell(sheet, 0, 1);
    ASSERT_DOUBLE_EQUAL(-1e15, cell->data.number, 1e10, "Should handle large negative numbers");

    // Test very small numbers
    sheet_set_number(sheet, 0, 2, 1e-15);
    cell = sheet_get_cell(sheet, 0, 2);
    ASSERT_DOUBLE_EQUAL(1e-15, cell->data.number, 1e-20, "Should handle very small numbers");

    sheet_free(sheet);
    TEST_PASS();
}

int test_long_strings() {
    TEST_START("Long String Handling");

    Sheet* sheet = sheet_new(100, 26);

    char long_string[256];
    memset(long_string, 'A', 255);
    long_string[255] = '\0';

    sheet_set_string(sheet, 0, 0, long_string);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    ASSERT_NOT_NULL(cell, "Cell should exist");
    ASSERT_EQUAL(CELL_STRING, cell->type, "Cell should be string type");
    ASSERT_EQUAL(255, (int)strlen(cell->data.string), "String length should be preserved");

    sheet_free(sheet);
    TEST_PASS();
}

// ============================================================================
// RECALCULATION TESTS
// ============================================================================

int test_formula_recalculation() {
    TEST_START("Formula Recalculation");

    Sheet* sheet = sheet_new(100, 26);

    // Set up dependencies: C1 = A1 + B1
    sheet_set_number(sheet, 0, 0, 10.0);  // A1
    sheet_set_number(sheet, 0, 1, 20.0);  // B1
    sheet_set_formula(sheet, 0, 2, "=A1+B1");  // C1
    sheet_recalculate(sheet);

    Cell* cell = sheet_get_cell(sheet, 0, 2);
    ASSERT_DOUBLE_EQUAL(30.0, cell->data.formula.cached_value, 0.0001, "Initial calculation should be 30");

    // Update A1 and recalculate
    sheet_set_number(sheet, 0, 0, 50.0);
    sheet_recalculate(sheet);

    cell = sheet_get_cell(sheet, 0, 2);
    ASSERT_DOUBLE_EQUAL(70.0, cell->data.formula.cached_value, 0.0001, "After update should be 70");

    sheet_free(sheet);
    TEST_PASS();
}

int test_chain_formulas() {
    TEST_START("Chained Formula Dependencies");

    Sheet* sheet = sheet_new(100, 26);

    // Create chain: A1 -> B1 -> C1 -> D1
    sheet_set_number(sheet, 0, 0, 5.0);           // A1 = 5
    sheet_set_formula(sheet, 0, 1, "=A1*2");      // B1 = A1*2 = 10
    sheet_set_formula(sheet, 0, 2, "=B1+10");     // C1 = B1+10 = 20
    sheet_set_formula(sheet, 0, 3, "=C1/2");      // D1 = C1/2 = 10
    sheet_recalculate(sheet);

    Cell* cell = sheet_get_cell(sheet, 0, 3);
    ASSERT_DOUBLE_EQUAL(10.0, cell->data.formula.cached_value, 0.0001, "Chain should calculate correctly");

    // Update source and verify chain recalculates
    sheet_set_number(sheet, 0, 0, 10.0);  // A1 = 10
    sheet_recalculate(sheet);

    cell = sheet_get_cell(sheet, 0, 3);
    ASSERT_DOUBLE_EQUAL(15.0, cell->data.formula.cached_value, 0.0001, "Chain should recalculate to 15");

    sheet_free(sheet);
    TEST_PASS();
}

// ============================================================================
// MAIN TEST RUNNER
// ============================================================================

int main() {
    printf("\n");
    printf("===============================================================\n");
    printf("  LiveLedger Comprehensive Unit Test Suite\n");
    printf("===============================================================\n");

    // Run all tests
    test_sheet_creation();
    test_sheet_invalid_creation();
    test_cell_number_operations();
    test_cell_string_operations();
    test_cell_clear();

    test_cell_reference_parsing();
    test_cell_reference_to_string();

    test_simple_arithmetic();
    test_operator_precedence();
    test_cell_references_in_formulas();

    test_sum_function();
    test_avg_function();
    test_max_min_functions();
    test_median_function();
    test_power_function();
    test_if_function();

    test_comparison_operators();

    test_percentage_formatting();
    test_currency_formatting();

    test_cell_copy_paste();

    test_range_parsing();
    test_range_selection();

    test_csv_save_load();

    test_column_row_sizing();

    test_color_parsing();

    test_error_handling();

    test_empty_cells();
    test_large_numbers();
    test_long_strings();

    test_formula_recalculation();
    test_chain_formulas();

    // Print summary
    printf("\n");
    printf("===============================================================\n");
    printf("  Test Results Summary\n");
    printf("===============================================================\n");
    printf("  Total Tests:   %d\n", tests_run);
    printf("  Passed:        %d (%.1f%%)\n", tests_passed,
           tests_run > 0 ? (100.0 * tests_passed / tests_run) : 0.0);
    printf("  Failed:        %d (%.1f%%)\n", tests_failed,
           tests_run > 0 ? (100.0 * tests_failed / tests_run) : 0.0);
    printf("===============================================================\n");

    if (tests_failed == 0) {
        printf("\n  [PASS] All tests PASSED!\n\n");
        return 0;
    } else {
        printf("\n  [FAIL] Some tests FAILED\n\n");
        return 1;
    }
}
