// test_liveledger.c - Comprehensive Unit Tests for LiveLedger
// Compile with: cl /O2 /W3 /TC test_liveledger.c sheet.c console.c charts.c /Fe:test_liveledger.exe /link user32.lib

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <assert.h>
#include <windows.h>

// Include the headers
#include "sheet.h"
#include "console.h"
#include "constants.h"

// Test framework macros
#define TEST_ASSERT(condition, msg) do { \
    tests_run++; \
    if (!(condition)) { \
        printf("  FAIL: %s (line %d)\n", msg, __LINE__); \
        tests_failed++; \
    } else { \
        tests_passed++; \
    } \
} while(0)

#define TEST_ASSERT_EQ_DOUBLE(expected, actual, tolerance, msg) do { \
    tests_run++; \
    if (fabs((expected) - (actual)) > (tolerance)) { \
        printf("  FAIL: %s - expected %f, got %f (line %d)\n", msg, (expected), (actual), __LINE__); \
        tests_failed++; \
    } else { \
        tests_passed++; \
    } \
} while(0)

#define TEST_ASSERT_EQ_INT(expected, actual, msg) do { \
    tests_run++; \
    if ((expected) != (actual)) { \
        printf("  FAIL: %s - expected %d, got %d (line %d)\n", msg, (expected), (actual), __LINE__); \
        tests_failed++; \
    } else { \
        tests_passed++; \
    } \
} while(0)

#define TEST_ASSERT_EQ_STR(expected, actual, msg) do { \
    tests_run++; \
    if ((expected) == NULL && (actual) == NULL) { \
        tests_passed++; \
    } else if ((expected) == NULL || (actual) == NULL || strcmp((expected), (actual)) != 0) { \
        printf("  FAIL: %s - expected '%s', got '%s' (line %d)\n", msg, (expected) ? (expected) : "NULL", (actual) ? (actual) : "NULL", __LINE__); \
        tests_failed++; \
    } else { \
        tests_passed++; \
    } \
} while(0)

#define TEST_SECTION(name) printf("\n=== Testing %s ===\n", name)

// Global test counters
static int tests_run = 0;
static int tests_passed = 0;
static int tests_failed = 0;

// ============================================================================
// SHEET OPERATIONS TESTS
// ============================================================================

void test_sheet_creation(void) {
    TEST_SECTION("Sheet Creation");
    
    // Test basic sheet creation
    Sheet* sheet = sheet_new(100, 26);
    TEST_ASSERT(sheet != NULL, "sheet_new should return non-NULL");
    TEST_ASSERT_EQ_INT(100, sheet->rows, "Sheet should have 100 rows");
    TEST_ASSERT_EQ_INT(26, sheet->cols, "Sheet should have 26 columns");
    TEST_ASSERT(sheet->cells != NULL, "Cells array should be allocated");
    TEST_ASSERT(sheet->col_widths != NULL, "Column widths array should be allocated");
    TEST_ASSERT(sheet->row_heights != NULL, "Row heights array should be allocated");
    TEST_ASSERT(sheet->name != NULL, "Sheet name should be set");
    TEST_ASSERT_EQ_STR("Sheet1", sheet->name, "Default sheet name should be 'Sheet1'");
    
    // Test default column widths
    for (int i = 0; i < 26; i++) {
        TEST_ASSERT_EQ_INT(DEFAULT_COLUMN_WIDTH, sheet->col_widths[i], "Default column width should be DEFAULT_COLUMN_WIDTH");
    }
    
    // Test default row heights
    for (int i = 0; i < 100; i++) {
        TEST_ASSERT_EQ_INT(1, sheet->row_heights[i], "Default row height should be 1");
    }
    
    sheet_free(sheet);
    
    // Test larger sheet creation
    Sheet* large_sheet = sheet_new(1000, 100);
    TEST_ASSERT(large_sheet != NULL, "Large sheet should be created");
    TEST_ASSERT_EQ_INT(1000, large_sheet->rows, "Large sheet should have 1000 rows");
    TEST_ASSERT_EQ_INT(100, large_sheet->cols, "Large sheet should have 100 columns");
    sheet_free(large_sheet);
    
    // Test minimum sheet
    Sheet* small_sheet = sheet_new(1, 1);
    TEST_ASSERT(small_sheet != NULL, "Minimum 1x1 sheet should be created");
    sheet_free(small_sheet);
}

void test_sheet_get_cell(void) {
    TEST_SECTION("Sheet Get Cell");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Test getting cell that doesn't exist yet
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    TEST_ASSERT(cell == NULL, "Cell should be NULL before creation");
    
    // Test out-of-bounds access
    cell = sheet_get_cell(sheet, -1, 0);
    TEST_ASSERT(cell == NULL, "Negative row should return NULL");
    
    cell = sheet_get_cell(sheet, 0, -1);
    TEST_ASSERT(cell == NULL, "Negative column should return NULL");
    
    cell = sheet_get_cell(sheet, 100, 0);
    TEST_ASSERT(cell == NULL, "Row >= rows should return NULL");
    
    cell = sheet_get_cell(sheet, 0, 26);
    TEST_ASSERT(cell == NULL, "Column >= cols should return NULL");
    
    sheet_free(sheet);
}

void test_sheet_get_or_create_cell(void) {
    TEST_SECTION("Sheet Get or Create Cell");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Test creating cell
    Cell* cell = sheet_get_or_create_cell(sheet, 5, 5);
    TEST_ASSERT(cell != NULL, "Cell should be created");
    TEST_ASSERT_EQ_INT(5, cell->row, "Cell row should be 5");
    TEST_ASSERT_EQ_INT(5, cell->col, "Cell column should be 5");
    TEST_ASSERT_EQ_INT(CELL_EMPTY, cell->type, "New cell should be CELL_EMPTY");
    
    // Test getting same cell again
    Cell* same_cell = sheet_get_or_create_cell(sheet, 5, 5);
    TEST_ASSERT(same_cell == cell, "Should return same cell pointer");
    
    // Test out-of-bounds
    Cell* oob_cell = sheet_get_or_create_cell(sheet, 100, 0);
    TEST_ASSERT(oob_cell == NULL, "Out-of-bounds should return NULL");
    
    sheet_free(sheet);
}

// ============================================================================
// CELL OPERATIONS TESTS
// ============================================================================

void test_cell_creation(void) {
    TEST_SECTION("Cell Creation");
    
    Cell* cell = cell_new(10, 20);
    TEST_ASSERT(cell != NULL, "cell_new should return non-NULL");
    TEST_ASSERT_EQ_INT(10, cell->row, "Cell row should be 10");
    TEST_ASSERT_EQ_INT(20, cell->col, "Cell column should be 20");
    TEST_ASSERT_EQ_INT(CELL_EMPTY, cell->type, "New cell should be CELL_EMPTY");
    TEST_ASSERT_EQ_INT(10, cell->width, "Default cell width should be 10");
    TEST_ASSERT_EQ_INT(2, cell->precision, "Default precision should be 2");
    TEST_ASSERT_EQ_INT(2, cell->align, "Default alignment should be 2 (right)");
    TEST_ASSERT_EQ_INT(FORMAT_GENERAL, cell->format, "Default format should be FORMAT_GENERAL");
    TEST_ASSERT_EQ_INT(-1, cell->text_color, "Default text color should be -1");
    TEST_ASSERT_EQ_INT(-1, cell->background_color, "Default background color should be -1");
    
    cell_free(cell);
}

void test_cell_set_number(void) {
    TEST_SECTION("Cell Set Number");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Test setting a number
    sheet_set_number(sheet, 0, 0, 42.5);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    TEST_ASSERT(cell != NULL, "Cell should exist after setting number");
    TEST_ASSERT_EQ_INT(CELL_NUMBER, cell->type, "Cell type should be CELL_NUMBER");
    TEST_ASSERT_EQ_DOUBLE(42.5, cell->data.number, 0.0001, "Cell value should be 42.5");
    
    // Test setting zero
    sheet_set_number(sheet, 1, 0, 0.0);
    cell = sheet_get_cell(sheet, 1, 0);
    TEST_ASSERT_EQ_INT(CELL_NUMBER, cell->type, "Cell type should be CELL_NUMBER for zero");
    TEST_ASSERT_EQ_DOUBLE(0.0, cell->data.number, 0.0001, "Cell value should be 0.0");
    
    // Test setting negative number
    sheet_set_number(sheet, 2, 0, -123.456);
    cell = sheet_get_cell(sheet, 2, 0);
    TEST_ASSERT_EQ_DOUBLE(-123.456, cell->data.number, 0.0001, "Cell value should be -123.456");
    
    // Test setting very large number
    sheet_set_number(sheet, 3, 0, 1e15);
    cell = sheet_get_cell(sheet, 3, 0);
    TEST_ASSERT_EQ_DOUBLE(1e15, cell->data.number, 1e9, "Cell value should be 1e15");
    
    // Test setting very small number
    sheet_set_number(sheet, 4, 0, 1e-10);
    cell = sheet_get_cell(sheet, 4, 0);
    TEST_ASSERT_EQ_DOUBLE(1e-10, cell->data.number, 1e-15, "Cell value should be 1e-10");
    
    sheet_free(sheet);
}

void test_cell_set_string(void) {
    TEST_SECTION("Cell Set String");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Test setting a string
    sheet_set_string(sheet, 0, 0, "Hello World");
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    TEST_ASSERT(cell != NULL, "Cell should exist after setting string");
    TEST_ASSERT_EQ_INT(CELL_STRING, cell->type, "Cell type should be CELL_STRING");
    TEST_ASSERT_EQ_STR("Hello World", cell->data.string, "Cell string should be 'Hello World'");
    TEST_ASSERT_EQ_INT(0, cell->align, "String cells should be left-aligned");
    
    // Test setting empty string
    sheet_set_string(sheet, 1, 0, "");
    cell = sheet_get_cell(sheet, 1, 0);
    TEST_ASSERT_EQ_INT(CELL_STRING, cell->type, "Cell type should be CELL_STRING for empty");
    TEST_ASSERT_EQ_STR("", cell->data.string, "Cell string should be empty");
    
    // Test setting string with special characters
    sheet_set_string(sheet, 2, 0, "Test, with \"quotes\" and \nnewline");
    cell = sheet_get_cell(sheet, 2, 0);
    TEST_ASSERT_EQ_STR("Test, with \"quotes\" and \nnewline", cell->data.string, "String with special chars");
    
    // Test overwriting number with string
    sheet_set_number(sheet, 3, 0, 100.0);
    sheet_set_string(sheet, 3, 0, "Now a string");
    cell = sheet_get_cell(sheet, 3, 0);
    TEST_ASSERT_EQ_INT(CELL_STRING, cell->type, "Cell type should change to CELL_STRING");
    TEST_ASSERT_EQ_STR("Now a string", cell->data.string, "Cell should now be a string");
    
    sheet_free(sheet);
}

void test_cell_set_formula(void) {
    TEST_SECTION("Cell Set Formula");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Set up some values for formulas to reference
    sheet_set_number(sheet, 0, 0, 10.0);  // A1 = 10
    sheet_set_number(sheet, 0, 1, 20.0);  // B1 = 20
    
    // Test setting a formula
    sheet_set_formula(sheet, 0, 2, "=A1+B1");
    Cell* cell = sheet_get_cell(sheet, 0, 2);
    TEST_ASSERT(cell != NULL, "Cell should exist after setting formula");
    TEST_ASSERT_EQ_INT(CELL_FORMULA, cell->type, "Cell type should be CELL_FORMULA");
    TEST_ASSERT_EQ_STR("=A1+B1", cell->data.formula.expression, "Formula expression should be stored");
    
    // Recalculate and check result
    sheet_recalculate(sheet);
    TEST_ASSERT_EQ_INT(ERROR_NONE, cell->data.formula.error, "Formula should have no error");
    TEST_ASSERT_EQ_DOUBLE(30.0, cell->data.formula.cached_value, 0.0001, "Formula result should be 30");
    
    sheet_free(sheet);
}

void test_cell_clear(void) {
    TEST_SECTION("Cell Clear");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Set a number and clear it
    sheet_set_number(sheet, 0, 0, 42.0);
    sheet_clear_cell(sheet, 0, 0);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    TEST_ASSERT_EQ_INT(CELL_EMPTY, cell->type, "Cell should be CELL_EMPTY after clearing");
    
    // Set a string and clear it
    sheet_set_string(sheet, 1, 0, "Test");
    sheet_clear_cell(sheet, 1, 0);
    cell = sheet_get_cell(sheet, 1, 0);
    TEST_ASSERT_EQ_INT(CELL_EMPTY, cell->type, "String cell should be CELL_EMPTY after clearing");
    
    // Set a formula and clear it
    sheet_set_formula(sheet, 2, 0, "=1+1");
    sheet_clear_cell(sheet, 2, 0);
    cell = sheet_get_cell(sheet, 2, 0);
    TEST_ASSERT_EQ_INT(CELL_EMPTY, cell->type, "Formula cell should be CELL_EMPTY after clearing");
    
    // Clear non-existent cell (should not crash)
    sheet_clear_cell(sheet, 50, 50);
    
    sheet_free(sheet);
}

// ============================================================================
// CELL REFERENCE TESTS
// ============================================================================

void test_cell_reference_parsing(void) {
    TEST_SECTION("Cell Reference Parsing");
    
    int row, col;
    
    // Test simple references
    TEST_ASSERT(parse_cell_reference("A1", &row, &col), "Should parse A1");
    TEST_ASSERT_EQ_INT(0, row, "A1 row should be 0");
    TEST_ASSERT_EQ_INT(0, col, "A1 col should be 0");
    
    TEST_ASSERT(parse_cell_reference("B1", &row, &col), "Should parse B1");
    TEST_ASSERT_EQ_INT(0, row, "B1 row should be 0");
    TEST_ASSERT_EQ_INT(1, col, "B1 col should be 1");
    
    TEST_ASSERT(parse_cell_reference("Z1", &row, &col), "Should parse Z1");
    TEST_ASSERT_EQ_INT(0, row, "Z1 row should be 0");
    TEST_ASSERT_EQ_INT(25, col, "Z1 col should be 25");
    
    // Test two-letter columns
    TEST_ASSERT(parse_cell_reference("AA1", &row, &col), "Should parse AA1");
    TEST_ASSERT_EQ_INT(0, row, "AA1 row should be 0");
    TEST_ASSERT_EQ_INT(26, col, "AA1 col should be 26");
    
    TEST_ASSERT(parse_cell_reference("AB1", &row, &col), "Should parse AB1");
    TEST_ASSERT_EQ_INT(0, row, "AB1 row should be 0");
    TEST_ASSERT_EQ_INT(27, col, "AB1 col should be 27");
    
    // Test larger row numbers
    TEST_ASSERT(parse_cell_reference("A100", &row, &col), "Should parse A100");
    TEST_ASSERT_EQ_INT(99, row, "A100 row should be 99");
    TEST_ASSERT_EQ_INT(0, col, "A100 col should be 0");
    
    TEST_ASSERT(parse_cell_reference("C999", &row, &col), "Should parse C999");
    TEST_ASSERT_EQ_INT(998, row, "C999 row should be 998");
    TEST_ASSERT_EQ_INT(2, col, "C999 col should be 2");
    
    // Test with whitespace
    TEST_ASSERT(parse_cell_reference("  A1", &row, &col), "Should parse with leading whitespace");
    TEST_ASSERT_EQ_INT(0, row, "Row should be 0");
    TEST_ASSERT_EQ_INT(0, col, "Col should be 0");
    
    // Test invalid references
    TEST_ASSERT(!parse_cell_reference("1A", &row, &col), "Should not parse 1A");
    TEST_ASSERT(!parse_cell_reference("A", &row, &col), "Should not parse A (no number)");
    TEST_ASSERT(!parse_cell_reference("1", &row, &col), "Should not parse 1 (no letter)");
    TEST_ASSERT(!parse_cell_reference("", &row, &col), "Should not parse empty string");
    TEST_ASSERT(!parse_cell_reference("A1B", &row, &col), "Should not parse A1B (trailing garbage)");
}

void test_cell_reference_to_string(void) {
    TEST_SECTION("Cell Reference to String");
    
    char buffer[16];
    
    cell_reference_to_string(0, 0, buffer, sizeof(buffer));
    TEST_ASSERT_EQ_STR("A1", buffer, "Row 0, Col 0 should be A1");
    
    cell_reference_to_string(0, 1, buffer, sizeof(buffer));
    TEST_ASSERT_EQ_STR("B1", buffer, "Row 0, Col 1 should be B1");
    
    cell_reference_to_string(0, 25, buffer, sizeof(buffer));
    TEST_ASSERT_EQ_STR("Z1", buffer, "Row 0, Col 25 should be Z1");
    
    cell_reference_to_string(0, 26, buffer, sizeof(buffer));
    TEST_ASSERT_EQ_STR("AA1", buffer, "Row 0, Col 26 should be AA1");
    
    cell_reference_to_string(0, 27, buffer, sizeof(buffer));
    TEST_ASSERT_EQ_STR("AB1", buffer, "Row 0, Col 27 should be AB1");
    
    cell_reference_to_string(99, 0, buffer, sizeof(buffer));
    TEST_ASSERT_EQ_STR("A100", buffer, "Row 99, Col 0 should be A100");
    
    cell_reference_to_string(999, 51, buffer, sizeof(buffer));
    TEST_ASSERT_EQ_STR("AZ1000", buffer, "Row 999, Col 51 should be AZ1000");
}

// ============================================================================
// FORMULA EVALUATION TESTS
// ============================================================================

void test_basic_arithmetic(void) {
    TEST_SECTION("Basic Arithmetic");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Addition
    sheet_set_formula(sheet, 0, 0, "=1+2");
    sheet_recalculate(sheet);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    TEST_ASSERT_EQ_DOUBLE(3.0, cell->data.formula.cached_value, 0.0001, "1+2 should be 3");
    
    // Subtraction
    sheet_set_formula(sheet, 1, 0, "=10-7");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 1, 0);
    TEST_ASSERT_EQ_DOUBLE(3.0, cell->data.formula.cached_value, 0.0001, "10-7 should be 3");
    
    // Multiplication
    sheet_set_formula(sheet, 2, 0, "=3*4");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 2, 0);
    TEST_ASSERT_EQ_DOUBLE(12.0, cell->data.formula.cached_value, 0.0001, "3*4 should be 12");
    
    // Division
    sheet_set_formula(sheet, 3, 0, "=20/4");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 3, 0);
    TEST_ASSERT_EQ_DOUBLE(5.0, cell->data.formula.cached_value, 0.0001, "20/4 should be 5");
    
    // Complex expression
    sheet_set_formula(sheet, 4, 0, "=1+2*3");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 4, 0);
    TEST_ASSERT_EQ_DOUBLE(7.0, cell->data.formula.cached_value, 0.0001, "1+2*3 should be 7");
    
    // Parentheses
    sheet_set_formula(sheet, 5, 0, "=(1+2)*3");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 5, 0);
    TEST_ASSERT_EQ_DOUBLE(9.0, cell->data.formula.cached_value, 0.0001, "(1+2)*3 should be 9");
    
    // Negative numbers
    sheet_set_formula(sheet, 6, 0, "=5+-3");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 6, 0);
    TEST_ASSERT_EQ_DOUBLE(2.0, cell->data.formula.cached_value, 0.0001, "5+-3 should be 2");
    
    // Decimals
    sheet_set_formula(sheet, 7, 0, "=1.5+2.5");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 7, 0);
    TEST_ASSERT_EQ_DOUBLE(4.0, cell->data.formula.cached_value, 0.0001, "1.5+2.5 should be 4");
    
    sheet_free(sheet);
}

void test_cell_references_in_formulas(void) {
    TEST_SECTION("Cell References in Formulas");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Set up values
    sheet_set_number(sheet, 0, 0, 10.0);  // A1 = 10
    sheet_set_number(sheet, 0, 1, 20.0);  // B1 = 20
    sheet_set_number(sheet, 1, 0, 5.0);   // A2 = 5
    
    // Simple reference
    sheet_set_formula(sheet, 2, 0, "=A1");
    sheet_recalculate(sheet);
    Cell* cell = sheet_get_cell(sheet, 2, 0);
    TEST_ASSERT_EQ_DOUBLE(10.0, cell->data.formula.cached_value, 0.0001, "=A1 should be 10");
    
    // Reference in addition
    sheet_set_formula(sheet, 2, 1, "=A1+B1");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 2, 1);
    TEST_ASSERT_EQ_DOUBLE(30.0, cell->data.formula.cached_value, 0.0001, "=A1+B1 should be 30");
    
    // Reference with constant
    sheet_set_formula(sheet, 2, 2, "=A1*2");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 2, 2);
    TEST_ASSERT_EQ_DOUBLE(20.0, cell->data.formula.cached_value, 0.0001, "=A1*2 should be 20");
    
    // Multiple references
    sheet_set_formula(sheet, 2, 3, "=A1+B1+A2");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 2, 3);
    TEST_ASSERT_EQ_DOUBLE(35.0, cell->data.formula.cached_value, 0.0001, "=A1+B1+A2 should be 35");
    
    // Reference to empty cell (should be 0)
    sheet_set_formula(sheet, 3, 0, "=Z99");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 3, 0);
    TEST_ASSERT_EQ_DOUBLE(0.0, cell->data.formula.cached_value, 0.0001, "Reference to empty cell should be 0");
    
    // Chain of references
    sheet_set_formula(sheet, 4, 0, "=A3");  // A5 = A3 = A1 = 10
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 4, 0);
    TEST_ASSERT_EQ_DOUBLE(10.0, cell->data.formula.cached_value, 0.0001, "Chain reference should work");
    
    sheet_free(sheet);
}

void test_division_by_zero(void) {
    TEST_SECTION("Division by Zero");
    
    Sheet* sheet = sheet_new(100, 26);
    
    sheet_set_formula(sheet, 0, 0, "=1/0");
    sheet_recalculate(sheet);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    TEST_ASSERT_EQ_INT(ERROR_DIV_ZERO, cell->data.formula.error, "Division by zero should set ERROR_DIV_ZERO");
    
    // Division by cell containing zero
    sheet_set_number(sheet, 1, 0, 0.0);
    sheet_set_formula(sheet, 1, 1, "=10/A2");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 1, 1);
    TEST_ASSERT_EQ_INT(ERROR_DIV_ZERO, cell->data.formula.error, "Division by zero cell should set ERROR_DIV_ZERO");
    
    sheet_free(sheet);
}

void test_sum_function(void) {
    TEST_SECTION("SUM Function");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Set up data
    sheet_set_number(sheet, 0, 0, 1.0);   // A1
    sheet_set_number(sheet, 1, 0, 2.0);   // A2
    sheet_set_number(sheet, 2, 0, 3.0);   // A3
    sheet_set_number(sheet, 3, 0, 4.0);   // A4
    sheet_set_number(sheet, 4, 0, 5.0);   // A5
    
    // SUM of vertical range
    sheet_set_formula(sheet, 5, 0, "=SUM(A1:A5)");
    sheet_recalculate(sheet);
    Cell* cell = sheet_get_cell(sheet, 5, 0);
    TEST_ASSERT_EQ_INT(ERROR_NONE, cell->data.formula.error, "SUM should have no error");
    TEST_ASSERT_EQ_DOUBLE(15.0, cell->data.formula.cached_value, 0.0001, "SUM(A1:A5) should be 15");
    
    // SUM of horizontal range
    sheet_set_number(sheet, 0, 1, 10.0);  // B1
    sheet_set_number(sheet, 0, 2, 20.0);  // C1
    sheet_set_number(sheet, 0, 3, 30.0);  // D1
    
    sheet_set_formula(sheet, 0, 4, "=SUM(B1:D1)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 0, 4);
    TEST_ASSERT_EQ_DOUBLE(60.0, cell->data.formula.cached_value, 0.0001, "SUM(B1:D1) should be 60");
    
    // SUM with negative numbers
    sheet_set_number(sheet, 6, 0, -5.0);
    sheet_set_number(sheet, 7, 0, 15.0);
    sheet_set_formula(sheet, 8, 0, "=SUM(A7:A8)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 8, 0);
    TEST_ASSERT_EQ_DOUBLE(10.0, cell->data.formula.cached_value, 0.0001, "SUM with negative should be 10");
    
    // SUM of single cell
    sheet_set_formula(sheet, 9, 0, "=SUM(A1:A1)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 9, 0);
    TEST_ASSERT_EQ_DOUBLE(1.0, cell->data.formula.cached_value, 0.0001, "SUM of single cell should be 1");
    
    // SUM with empty cells (should treat as 0)
    sheet_set_formula(sheet, 10, 0, "=SUM(Z1:Z5)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 10, 0);
    TEST_ASSERT_EQ_DOUBLE(0.0, cell->data.formula.cached_value, 0.0001, "SUM of empty cells should be 0");
    
    sheet_free(sheet);
}

void test_avg_function(void) {
    TEST_SECTION("AVG Function");
    
    Sheet* sheet = sheet_new(100, 26);
    
    sheet_set_number(sheet, 0, 0, 10.0);
    sheet_set_number(sheet, 1, 0, 20.0);
    sheet_set_number(sheet, 2, 0, 30.0);
    
    sheet_set_formula(sheet, 3, 0, "=AVG(A1:A3)");
    sheet_recalculate(sheet);
    Cell* cell = sheet_get_cell(sheet, 3, 0);
    TEST_ASSERT_EQ_DOUBLE(20.0, cell->data.formula.cached_value, 0.0001, "AVG(A1:A3) should be 20");
    
    // Single value
    sheet_set_formula(sheet, 4, 0, "=AVG(A1:A1)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 4, 0);
    TEST_ASSERT_EQ_DOUBLE(10.0, cell->data.formula.cached_value, 0.0001, "AVG of single value should be that value");
    
    sheet_free(sheet);
}

void test_max_min_functions(void) {
    TEST_SECTION("MAX and MIN Functions");
    
    Sheet* sheet = sheet_new(100, 26);
    
    sheet_set_number(sheet, 0, 0, 5.0);
    sheet_set_number(sheet, 1, 0, 10.0);
    sheet_set_number(sheet, 2, 0, 3.0);
    sheet_set_number(sheet, 3, 0, 8.0);
    
    // MAX
    sheet_set_formula(sheet, 4, 0, "=MAX(A1:A4)");
    sheet_recalculate(sheet);
    Cell* cell = sheet_get_cell(sheet, 4, 0);
    TEST_ASSERT_EQ_DOUBLE(10.0, cell->data.formula.cached_value, 0.0001, "MAX(A1:A4) should be 10");
    
    // MIN
    sheet_set_formula(sheet, 5, 0, "=MIN(A1:A4)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 5, 0);
    TEST_ASSERT_EQ_DOUBLE(3.0, cell->data.formula.cached_value, 0.0001, "MIN(A1:A4) should be 3");
    
    // With negative numbers
    sheet_set_number(sheet, 6, 0, -5.0);
    sheet_set_number(sheet, 7, 0, -2.0);
    
    sheet_set_formula(sheet, 8, 0, "=MAX(A7:A8)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 8, 0);
    TEST_ASSERT_EQ_DOUBLE(-2.0, cell->data.formula.cached_value, 0.0001, "MAX of negatives should be -2");
    
    sheet_set_formula(sheet, 9, 0, "=MIN(A7:A8)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 9, 0);
    TEST_ASSERT_EQ_DOUBLE(-5.0, cell->data.formula.cached_value, 0.0001, "MIN of negatives should be -5");
    
    sheet_free(sheet);
}

void test_median_function(void) {
    TEST_SECTION("MEDIAN Function");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Odd number of values
    sheet_set_number(sheet, 0, 0, 1.0);
    sheet_set_number(sheet, 1, 0, 2.0);
    sheet_set_number(sheet, 2, 0, 3.0);
    sheet_set_number(sheet, 3, 0, 4.0);
    sheet_set_number(sheet, 4, 0, 5.0);
    
    sheet_set_formula(sheet, 5, 0, "=MEDIAN(A1:A5)");
    sheet_recalculate(sheet);
    Cell* cell = sheet_get_cell(sheet, 5, 0);
    TEST_ASSERT_EQ_DOUBLE(3.0, cell->data.formula.cached_value, 0.0001, "MEDIAN of 1,2,3,4,5 should be 3");
    
    // Even number of values
    sheet_set_number(sheet, 0, 1, 1.0);
    sheet_set_number(sheet, 1, 1, 2.0);
    sheet_set_number(sheet, 2, 1, 3.0);
    sheet_set_number(sheet, 3, 1, 4.0);
    
    sheet_set_formula(sheet, 5, 1, "=MEDIAN(B1:B4)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 5, 1);
    TEST_ASSERT_EQ_DOUBLE(2.5, cell->data.formula.cached_value, 0.0001, "MEDIAN of 1,2,3,4 should be 2.5");
    
    // Unsorted values
    sheet_set_number(sheet, 0, 2, 5.0);
    sheet_set_number(sheet, 1, 2, 1.0);
    sheet_set_number(sheet, 2, 2, 3.0);
    
    sheet_set_formula(sheet, 5, 2, "=MEDIAN(C1:C3)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 5, 2);
    TEST_ASSERT_EQ_DOUBLE(3.0, cell->data.formula.cached_value, 0.0001, "MEDIAN of unsorted 5,1,3 should be 3");
    
    sheet_free(sheet);
}

void test_mode_function(void) {
    TEST_SECTION("MODE Function");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Values with clear mode
    sheet_set_number(sheet, 0, 0, 1.0);
    sheet_set_number(sheet, 1, 0, 2.0);
    sheet_set_number(sheet, 2, 0, 2.0);
    sheet_set_number(sheet, 3, 0, 3.0);
    sheet_set_number(sheet, 4, 0, 2.0);
    
    sheet_set_formula(sheet, 5, 0, "=MODE(A1:A5)");
    sheet_recalculate(sheet);
    Cell* cell = sheet_get_cell(sheet, 5, 0);
    TEST_ASSERT_EQ_DOUBLE(2.0, cell->data.formula.cached_value, 0.0001, "MODE should be 2");
    
    sheet_free(sheet);
}

void test_if_function(void) {
    TEST_SECTION("IF Function");
    
    Sheet* sheet = sheet_new(100, 26);
    
    sheet_set_number(sheet, 0, 0, 10.0);  // A1 = 10
    sheet_set_number(sheet, 0, 1, 5.0);   // B1 = 5
    
    // Numeric IF - true condition
    sheet_set_formula(sheet, 1, 0, "=IF(A1>B1, 100, 200)");
    sheet_recalculate(sheet);
    Cell* cell = sheet_get_cell(sheet, 1, 0);
    TEST_ASSERT_EQ_DOUBLE(100.0, cell->data.formula.cached_value, 0.0001, "IF(10>5) should return true value");
    
    // Numeric IF - false condition
    sheet_set_formula(sheet, 1, 1, "=IF(A1<B1, 100, 200)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 1, 1);
    TEST_ASSERT_EQ_DOUBLE(200.0, cell->data.formula.cached_value, 0.0001, "IF(10<5) should return false value");
    
    // IF with equals
    sheet_set_formula(sheet, 2, 0, "=IF(A1=10, 1, 0)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 2, 0);
    TEST_ASSERT_EQ_DOUBLE(1.0, cell->data.formula.cached_value, 0.0001, "IF(A1=10) should be true");
    
    // IF with not equals
    sheet_set_formula(sheet, 2, 1, "=IF(A1<>10, 1, 0)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 2, 1);
    TEST_ASSERT_EQ_DOUBLE(0.0, cell->data.formula.cached_value, 0.0001, "IF(A1<>10) should be false");
    
    // IF with >= and <=
    sheet_set_formula(sheet, 3, 0, "=IF(A1>=10, 1, 0)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 3, 0);
    TEST_ASSERT_EQ_DOUBLE(1.0, cell->data.formula.cached_value, 0.0001, "IF(A1>=10) should be true");
    
    sheet_set_formula(sheet, 3, 1, "=IF(A1<=10, 1, 0)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 3, 1);
    TEST_ASSERT_EQ_DOUBLE(1.0, cell->data.formula.cached_value, 0.0001, "IF(A1<=10) should be true");
    
    // String IF results
    sheet_set_formula(sheet, 4, 0, "=IF(A1>5, \"High\", \"Low\")");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 4, 0);
    TEST_ASSERT(cell->data.formula.is_string_result, "IF with string result should set is_string_result");
    TEST_ASSERT_EQ_STR("High", cell->data.formula.cached_string, "IF should return 'High'");
    
    sheet_free(sheet);
}

void test_power_function(void) {
    TEST_SECTION("POWER Function");
    
    Sheet* sheet = sheet_new(100, 26);
    
    sheet_set_formula(sheet, 0, 0, "=POWER(2, 3)");
    sheet_recalculate(sheet);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    TEST_ASSERT_EQ_DOUBLE(8.0, cell->data.formula.cached_value, 0.0001, "POWER(2,3) should be 8");
    
    sheet_set_formula(sheet, 1, 0, "=POWER(10, 2)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 1, 0);
    TEST_ASSERT_EQ_DOUBLE(100.0, cell->data.formula.cached_value, 0.0001, "POWER(10,2) should be 100");
    
    // Square root
    sheet_set_formula(sheet, 2, 0, "=POWER(16, 0.5)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 2, 0);
    TEST_ASSERT_EQ_DOUBLE(4.0, cell->data.formula.cached_value, 0.0001, "POWER(16,0.5) should be 4");
    
    // Power of 0
    sheet_set_formula(sheet, 3, 0, "=POWER(5, 0)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 3, 0);
    TEST_ASSERT_EQ_DOUBLE(1.0, cell->data.formula.cached_value, 0.0001, "POWER(5,0) should be 1");
    
    // Cell reference
    sheet_set_number(sheet, 4, 0, 3.0);
    sheet_set_number(sheet, 4, 1, 4.0);
    sheet_set_formula(sheet, 4, 2, "=POWER(A5, B5)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 4, 2);
    TEST_ASSERT_EQ_DOUBLE(81.0, cell->data.formula.cached_value, 0.0001, "POWER(3,4) should be 81");
    
    sheet_free(sheet);
}

void test_xlookup_function(void) {
    TEST_SECTION("XLOOKUP Function");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Set up lookup table
    sheet_set_string(sheet, 0, 0, "Apple");   // A1
    sheet_set_number(sheet, 0, 1, 0.5);       // B1
    sheet_set_number(sheet, 0, 2, 100);       // C1
    
    sheet_set_string(sheet, 1, 0, "Orange");  // A2
    sheet_set_number(sheet, 1, 1, 0.75);      // B2
    sheet_set_number(sheet, 1, 2, 85);        // C2
    
    sheet_set_string(sheet, 2, 0, "Banana");  // A3
    sheet_set_number(sheet, 2, 1, 0.3);       // B3
    sheet_set_number(sheet, 2, 2, 120);       // C3
    
    // String lookup - find price
    sheet_set_formula(sheet, 5, 0, "=XLOOKUP(\"Orange\", A1:A3, B1:B3, 0)");
    sheet_recalculate(sheet);
    Cell* cell = sheet_get_cell(sheet, 5, 0);
    TEST_ASSERT_EQ_INT(ERROR_NONE, cell->data.formula.error, "XLOOKUP should succeed");
    TEST_ASSERT_EQ_DOUBLE(0.75, cell->data.formula.cached_value, 0.0001, "XLOOKUP for Orange price should be 0.75");
    
    // String lookup - find stock
    sheet_set_formula(sheet, 5, 1, "=XLOOKUP(\"Banana\", A1:A3, C1:C3, 0)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 5, 1);
    TEST_ASSERT_EQ_DOUBLE(120.0, cell->data.formula.cached_value, 0.0001, "XLOOKUP for Banana stock should be 120");
    
    // Lookup not found
    sheet_set_formula(sheet, 5, 2, "=XLOOKUP(\"Grape\", A1:A3, B1:B3, 0)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 5, 2);
    TEST_ASSERT_EQ_INT(ERROR_NA, cell->data.formula.error, "XLOOKUP for missing item should return #N/A");
    
    // Numeric lookup
    sheet_set_number(sheet, 10, 0, 1);   // A11
    sheet_set_number(sheet, 10, 1, 100); // B11
    sheet_set_number(sheet, 11, 0, 2);   // A12
    sheet_set_number(sheet, 11, 1, 200); // B12
    sheet_set_number(sheet, 12, 0, 3);   // A13
    sheet_set_number(sheet, 12, 1, 300); // B13
    
    sheet_set_formula(sheet, 13, 0, "=XLOOKUP(2, A11:A13, B11:B13, 0)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 13, 0);
    TEST_ASSERT_EQ_DOUBLE(200.0, cell->data.formula.cached_value, 0.0001, "Numeric XLOOKUP should find 200");
    
    sheet_free(sheet);
}

void test_nested_functions(void) {
    TEST_SECTION("Nested Functions");
    
    Sheet* sheet = sheet_new(100, 26);
    
    sheet_set_number(sheet, 0, 0, 1.0);
    sheet_set_number(sheet, 1, 0, 2.0);
    sheet_set_number(sheet, 2, 0, 3.0);
    
    // SUM inside expression
    sheet_set_formula(sheet, 3, 0, "=SUM(A1:A3)*2");
    sheet_recalculate(sheet);
    Cell* cell = sheet_get_cell(sheet, 3, 0);
    TEST_ASSERT_EQ_DOUBLE(12.0, cell->data.formula.cached_value, 0.0001, "SUM*2 should be 12");
    
    // Division of functions
    sheet_set_formula(sheet, 4, 0, "=SUM(A1:A3)/MAX(A1:A3)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 4, 0);
    TEST_ASSERT_EQ_DOUBLE(2.0, cell->data.formula.cached_value, 0.0001, "SUM/MAX should be 2");
    
    sheet_free(sheet);
}

// ============================================================================
// FORMATTING TESTS
// ============================================================================

void test_percentage_format(void) {
    TEST_SECTION("Percentage Format");
    
    Sheet* sheet = sheet_new(100, 26);
    
    sheet_set_number(sheet, 0, 0, 0.5);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    cell_set_format(cell, FORMAT_PERCENTAGE, 0);
    
    char* display = format_cell_value(cell);
    TEST_ASSERT(strstr(display, "50") != NULL, "0.5 as percentage should contain '50'");
    TEST_ASSERT(strstr(display, "%") != NULL, "Percentage should contain '%' symbol");
    
    sheet_set_number(sheet, 1, 0, 0.1234);
    cell = sheet_get_cell(sheet, 1, 0);
    cell_set_format(cell, FORMAT_PERCENTAGE, 0);
    display = format_cell_value(cell);
    TEST_ASSERT(strstr(display, "12") != NULL, "0.1234 as percentage should contain '12'");
    
    sheet_set_number(sheet, 2, 0, 1.5);
    cell = sheet_get_cell(sheet, 2, 0);
    cell_set_format(cell, FORMAT_PERCENTAGE, 0);
    display = format_cell_value(cell);
    TEST_ASSERT(strstr(display, "150") != NULL, "1.5 as percentage should contain '150'");
    
    sheet_free(sheet);
}

void test_currency_format(void) {
    TEST_SECTION("Currency Format");
    
    Sheet* sheet = sheet_new(100, 26);
    
    sheet_set_number(sheet, 0, 0, 1234.56);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    cell_set_format(cell, FORMAT_CURRENCY, 0);
    
    char* display = format_cell_value(cell);
    TEST_ASSERT(strstr(display, "$") != NULL, "Currency should contain '$'");
    TEST_ASSERT(strstr(display, "1234") != NULL, "Currency should contain the value");
    
    // Negative currency
    sheet_set_number(sheet, 1, 0, -500.00);
    cell = sheet_get_cell(sheet, 1, 0);
    cell_set_format(cell, FORMAT_CURRENCY, 0);
    display = format_cell_value(cell);
    TEST_ASSERT(strstr(display, "-") != NULL, "Negative currency should show minus");
    TEST_ASSERT(strstr(display, "500") != NULL, "Negative currency should show value");
    
    sheet_free(sheet);
}

void test_date_formats(void) {
    TEST_SECTION("Date Formats");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Excel serial date for a known date (e.g., 44927 = 12/25/2022)
    // Note: Excel dates start from 1/1/1900
    double test_date = 44927.0;  // This should be around Dec 25, 2022
    
    sheet_set_number(sheet, 0, 0, test_date);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    
    // Test MM/DD/YYYY format
    cell_set_format(cell, FORMAT_DATE, DATE_STYLE_MM_DD_YYYY);
    char* display = format_cell_value(cell);
    TEST_ASSERT(display != NULL, "Date display should not be NULL");
    TEST_ASSERT(strlen(display) > 0, "Date display should not be empty");
    TEST_ASSERT(strchr(display, '/') != NULL || strchr(display, '-') != NULL, "Date should contain separator");
    
    // Test DD/MM/YYYY format
    cell_set_format(cell, FORMAT_DATE, DATE_STYLE_DD_MM_YYYY);
    display = format_cell_value(cell);
    TEST_ASSERT(display != NULL, "DD/MM/YYYY format should work");
    
    // Test YYYY-MM-DD format
    cell_set_format(cell, FORMAT_DATE, DATE_STYLE_YYYY_MM_DD);
    display = format_cell_value(cell);
    TEST_ASSERT(display != NULL, "YYYY-MM-DD format should work");
    
    // Test long format
    cell_set_format(cell, FORMAT_DATE, DATE_STYLE_MON_DD_YYYY);
    display = format_cell_value(cell);
    TEST_ASSERT(display != NULL, "Long date format should work");
    
    sheet_free(sheet);
}

void test_time_formats(void) {
    TEST_SECTION("Time Formats");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Time fraction: 0.5 = 12:00 PM (noon)
    sheet_set_number(sheet, 0, 0, 0.5);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    
    // Test 12-hour format
    cell_set_format(cell, FORMAT_TIME, TIME_STYLE_12HR);
    char* display = format_cell_value(cell);
    TEST_ASSERT(display != NULL, "Time display should not be NULL");
    TEST_ASSERT(strstr(display, "12") != NULL || strstr(display, ":") != NULL, "Time should show hours");
    
    // Test 24-hour format
    cell_set_format(cell, FORMAT_TIME, TIME_STYLE_24HR);
    display = format_cell_value(cell);
    TEST_ASSERT(display != NULL, "24-hour format should work");
    
    // Test with seconds
    cell_set_format(cell, FORMAT_TIME, TIME_STYLE_SECONDS);
    display = format_cell_value(cell);
    TEST_ASSERT(display != NULL, "Time with seconds should work");
    
    // Test 6:00 PM (0.75)
    sheet_set_number(sheet, 1, 0, 0.75);
    cell = sheet_get_cell(sheet, 1, 0);
    cell_set_format(cell, FORMAT_TIME, TIME_STYLE_12HR);
    display = format_cell_value(cell);
    TEST_ASSERT(strstr(display, "6") != NULL || strstr(display, "PM") != NULL, "0.75 should be 6 PM");
    
    sheet_free(sheet);
}

void test_cell_colors(void) {
    TEST_SECTION("Cell Colors");
    
    Sheet* sheet = sheet_new(100, 26);
    
    Cell* cell = sheet_get_or_create_cell(sheet, 0, 0);
    
    // Test default colors
    TEST_ASSERT_EQ_INT(-1, cell->text_color, "Default text color should be -1");
    TEST_ASSERT_EQ_INT(-1, cell->background_color, "Default background color should be -1");
    
    // Set text color
    cell_set_text_color(cell, COLOR_RED);
    TEST_ASSERT_EQ_INT(COLOR_RED, cell->text_color, "Text color should be red");
    
    // Set background color
    cell_set_background_color(cell, COLOR_YELLOW);
    TEST_ASSERT_EQ_INT(COLOR_YELLOW, cell->background_color, "Background color should be yellow");
    
    sheet_free(sheet);
}

void test_color_parsing(void) {
    TEST_SECTION("Color Parsing");
    
    // Test named colors
    TEST_ASSERT_EQ_INT(COLOR_BLACK, parse_color("black"), "Should parse 'black'");
    TEST_ASSERT_EQ_INT(COLOR_BLUE, parse_color("blue"), "Should parse 'blue'");
    TEST_ASSERT_EQ_INT(COLOR_GREEN, parse_color("green"), "Should parse 'green'");
    TEST_ASSERT_EQ_INT(COLOR_CYAN, parse_color("cyan"), "Should parse 'cyan'");
    TEST_ASSERT_EQ_INT(COLOR_RED, parse_color("red"), "Should parse 'red'");
    TEST_ASSERT_EQ_INT(COLOR_MAGENTA, parse_color("magenta"), "Should parse 'magenta'");
    TEST_ASSERT_EQ_INT(COLOR_YELLOW, parse_color("yellow"), "Should parse 'yellow'");
    TEST_ASSERT_EQ_INT(COLOR_WHITE, parse_color("white"), "Should parse 'white'");
    
    // Test hex colors (these map to nearest console color)
    int red_hex = parse_color("#FF0000");
    TEST_ASSERT(red_hex >= 0, "Should parse #FF0000");
    
    int blue_hex = parse_color("#0000FF");
    TEST_ASSERT(blue_hex >= 0, "Should parse #0000FF");
    
    // Test invalid colors
    TEST_ASSERT_EQ_INT(-1, parse_color("invalid"), "Should return -1 for invalid color");
    TEST_ASSERT_EQ_INT(-1, parse_color("#GGG"), "Should return -1 for invalid hex");
    TEST_ASSERT_EQ_INT(-1, parse_color(""), "Should return -1 for empty string");
}

// ============================================================================
// RANGE OPERATIONS TESTS
// ============================================================================

void test_range_selection(void) {
    TEST_SECTION("Range Selection");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Test initial state
    TEST_ASSERT_EQ_INT(0, sheet->selection.is_active, "Selection should not be active initially");
    
    // Start selection
    sheet_start_range_selection(sheet, 5, 5);
    TEST_ASSERT_EQ_INT(1, sheet->selection.is_active, "Selection should be active after start");
    TEST_ASSERT_EQ_INT(5, sheet->selection.start_row, "Start row should be 5");
    TEST_ASSERT_EQ_INT(5, sheet->selection.start_col, "Start col should be 5");
    TEST_ASSERT_EQ_INT(5, sheet->selection.end_row, "End row should initially be 5");
    TEST_ASSERT_EQ_INT(5, sheet->selection.end_col, "End col should initially be 5");
    
    // Extend selection
    sheet_extend_range_selection(sheet, 10, 8);
    TEST_ASSERT_EQ_INT(10, sheet->selection.end_row, "End row should be 10");
    TEST_ASSERT_EQ_INT(8, sheet->selection.end_col, "End col should be 8");
    
    // Test is_in_selection
    TEST_ASSERT(sheet_is_in_selection(sheet, 5, 5), "5,5 should be in selection");
    TEST_ASSERT(sheet_is_in_selection(sheet, 7, 6), "7,6 should be in selection");
    TEST_ASSERT(sheet_is_in_selection(sheet, 10, 8), "10,8 should be in selection");
    TEST_ASSERT(!sheet_is_in_selection(sheet, 4, 5), "4,5 should not be in selection");
    TEST_ASSERT(!sheet_is_in_selection(sheet, 5, 4), "5,4 should not be in selection");
    TEST_ASSERT(!sheet_is_in_selection(sheet, 11, 8), "11,8 should not be in selection");
    
    // Clear selection
    sheet_clear_range_selection(sheet);
    TEST_ASSERT_EQ_INT(0, sheet->selection.is_active, "Selection should not be active after clear");
    TEST_ASSERT(!sheet_is_in_selection(sheet, 5, 5), "Nothing should be in selection after clear");
    
    sheet_free(sheet);
}

void test_range_copy_paste(void) {
    TEST_SECTION("Range Copy/Paste");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Set up source data
    sheet_set_number(sheet, 0, 0, 1.0);
    sheet_set_number(sheet, 0, 1, 2.0);
    sheet_set_number(sheet, 1, 0, 3.0);
    sheet_set_number(sheet, 1, 1, 4.0);
    
    // Select and copy the range A1:B2
    sheet_start_range_selection(sheet, 0, 0);
    sheet_extend_range_selection(sheet, 1, 1);
    sheet_copy_range(sheet);
    
    TEST_ASSERT_EQ_INT(1, sheet->range_clipboard.is_active, "Clipboard should be active after copy");
    TEST_ASSERT_EQ_INT(2, sheet->range_clipboard.rows, "Clipboard should have 2 rows");
    TEST_ASSERT_EQ_INT(2, sheet->range_clipboard.cols, "Clipboard should have 2 cols");
    
    // Paste at a different location
    sheet_paste_range(sheet, 5, 5);
    
    // Verify paste
    Cell* cell = sheet_get_cell(sheet, 5, 5);
    TEST_ASSERT(cell != NULL, "Pasted cell F6 should exist");
    TEST_ASSERT_EQ_INT(CELL_NUMBER, cell->type, "Pasted cell should be number");
    TEST_ASSERT_EQ_DOUBLE(1.0, cell->data.number, 0.0001, "F6 should be 1");
    
    cell = sheet_get_cell(sheet, 5, 6);
    TEST_ASSERT_EQ_DOUBLE(2.0, cell->data.number, 0.0001, "G6 should be 2");
    
    cell = sheet_get_cell(sheet, 6, 5);
    TEST_ASSERT_EQ_DOUBLE(3.0, cell->data.number, 0.0001, "F7 should be 3");
    
    cell = sheet_get_cell(sheet, 6, 6);
    TEST_ASSERT_EQ_DOUBLE(4.0, cell->data.number, 0.0001, "G7 should be 4");
    
    sheet_free(sheet);
}

void test_clipboard_cell(void) {
    TEST_SECTION("Cell Clipboard");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Set up a cell
    sheet_set_number(sheet, 0, 0, 42.0);
    Cell* source = sheet_get_cell(sheet, 0, 0);
    
    // Copy to clipboard
    sheet_set_clipboard_cell(source);
    Cell* clipboard = sheet_get_clipboard_cell();
    TEST_ASSERT(clipboard != NULL, "Clipboard should not be NULL after copy");
    TEST_ASSERT_EQ_INT(CELL_NUMBER, clipboard->type, "Clipboard cell should be number");
    TEST_ASSERT_EQ_DOUBLE(42.0, clipboard->data.number, 0.0001, "Clipboard value should be 42");
    
    // Copy cell to destination
    sheet_copy_cell(sheet, 0, 0, 5, 5);
    Cell* dest = sheet_get_cell(sheet, 5, 5);
    TEST_ASSERT(dest != NULL, "Destination should exist");
    TEST_ASSERT_EQ_DOUBLE(42.0, dest->data.number, 0.0001, "Destination should have value 42");
    
    // Test copying string
    sheet_set_string(sheet, 1, 0, "Test String");
    source = sheet_get_cell(sheet, 1, 0);
    sheet_set_clipboard_cell(source);
    clipboard = sheet_get_clipboard_cell();
    TEST_ASSERT_EQ_INT(CELL_STRING, clipboard->type, "Clipboard should be string");
    TEST_ASSERT_EQ_STR("Test String", clipboard->data.string, "Clipboard string should match");
    
    // Test copying formula
    sheet_set_formula(sheet, 2, 0, "=1+1");
    sheet_recalculate(sheet);
    source = sheet_get_cell(sheet, 2, 0);
    sheet_set_clipboard_cell(source);
    clipboard = sheet_get_clipboard_cell();
    TEST_ASSERT_EQ_INT(CELL_FORMULA, clipboard->type, "Clipboard should be formula");
    TEST_ASSERT_EQ_STR("=1+1", clipboard->data.formula.expression, "Clipboard formula should match");
    
    sheet_free(sheet);
}

// ============================================================================
// COLUMN AND ROW SIZING TESTS
// ============================================================================

void test_column_width(void) {
    TEST_SECTION("Column Width");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Test default width
    TEST_ASSERT_EQ_INT(DEFAULT_COLUMN_WIDTH, sheet_get_column_width(sheet, 0), "Default width should be DEFAULT_COLUMN_WIDTH");
    
    // Set custom width
    sheet_set_column_width(sheet, 0, 20);
    TEST_ASSERT_EQ_INT(20, sheet_get_column_width(sheet, 0), "Width should be 20 after set");
    
    // Test min width constraint
    sheet_set_column_width(sheet, 1, 0);  // Below minimum
    TEST_ASSERT(sheet_get_column_width(sheet, 1) >= MIN_COLUMN_WIDTH, "Width should be at least MIN_COLUMN_WIDTH");
    
    // Test resize in range
    sheet_set_column_width(sheet, 5, 10);
    sheet_set_column_width(sheet, 6, 10);
    sheet_set_column_width(sheet, 7, 10);
    
    sheet_resize_columns_in_range(sheet, 5, 7, 5);  // Increase by 5
    TEST_ASSERT_EQ_INT(15, sheet_get_column_width(sheet, 5), "Width should increase to 15");
    TEST_ASSERT_EQ_INT(15, sheet_get_column_width(sheet, 6), "Width should increase to 15");
    TEST_ASSERT_EQ_INT(15, sheet_get_column_width(sheet, 7), "Width should increase to 15");
    
    // Test decrease
    sheet_resize_columns_in_range(sheet, 5, 7, -3);
    TEST_ASSERT_EQ_INT(12, sheet_get_column_width(sheet, 5), "Width should decrease to 12");
    
    // Test out of bounds
    TEST_ASSERT_EQ_INT(DEFAULT_COLUMN_WIDTH, sheet_get_column_width(sheet, -1), "Negative index should return default");
    TEST_ASSERT_EQ_INT(DEFAULT_COLUMN_WIDTH, sheet_get_column_width(sheet, 100), "Out of bounds should return default");
    
    sheet_free(sheet);
}

void test_row_height(void) {
    TEST_SECTION("Row Height");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Test default height
    TEST_ASSERT_EQ_INT(1, sheet_get_row_height(sheet, 0), "Default height should be 1");
    
    // Set custom height
    sheet_set_row_height(sheet, 0, 3);
    TEST_ASSERT_EQ_INT(3, sheet_get_row_height(sheet, 0), "Height should be 3 after set");
    
    // Test min height constraint
    sheet_set_row_height(sheet, 1, 0);  // Below minimum
    TEST_ASSERT(sheet_get_row_height(sheet, 1) >= MIN_ROW_HEIGHT, "Height should be at least MIN_ROW_HEIGHT");
    
    // Test resize in range
    sheet_set_row_height(sheet, 5, 2);
    sheet_set_row_height(sheet, 6, 2);
    
    sheet_resize_rows_in_range(sheet, 5, 6, 2);  // Increase by 2
    TEST_ASSERT_EQ_INT(4, sheet_get_row_height(sheet, 5), "Height should increase to 4");
    TEST_ASSERT_EQ_INT(4, sheet_get_row_height(sheet, 6), "Height should increase to 4");
    
    sheet_free(sheet);
}

// ============================================================================
// INSERT/DELETE ROW/COLUMN TESTS
// ============================================================================

void test_insert_row(void) {
    TEST_SECTION("Insert Row");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Set up data
    sheet_set_number(sheet, 0, 0, 1.0);  // A1
    sheet_set_number(sheet, 1, 0, 2.0);  // A2
    sheet_set_number(sheet, 2, 0, 3.0);  // A3
    
    // Insert row at position 1
    sheet_insert_row(sheet, 1);
    
    // Verify data shifted down
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    TEST_ASSERT_EQ_DOUBLE(1.0, cell->data.number, 0.0001, "A1 should still be 1");
    
    cell = sheet_get_cell(sheet, 1, 0);
    TEST_ASSERT(cell == NULL || cell->type == CELL_EMPTY, "A2 (new row) should be empty");
    
    cell = sheet_get_cell(sheet, 2, 0);
    TEST_ASSERT(cell != NULL, "A3 should exist");
    TEST_ASSERT_EQ_DOUBLE(2.0, cell->data.number, 0.0001, "A3 should now be 2 (shifted from A2)");
    
    cell = sheet_get_cell(sheet, 3, 0);
    TEST_ASSERT_EQ_DOUBLE(3.0, cell->data.number, 0.0001, "A4 should now be 3 (shifted from A3)");
    
    sheet_free(sheet);
}

void test_delete_row(void) {
    TEST_SECTION("Delete Row");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Set up data
    sheet_set_number(sheet, 0, 0, 1.0);  // A1
    sheet_set_number(sheet, 1, 0, 2.0);  // A2
    sheet_set_number(sheet, 2, 0, 3.0);  // A3
    sheet_set_number(sheet, 3, 0, 4.0);  // A4
    
    // Delete row 1 (second row)
    sheet_delete_row(sheet, 1);
    
    // Verify data shifted up
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    TEST_ASSERT_EQ_DOUBLE(1.0, cell->data.number, 0.0001, "A1 should still be 1");
    
    cell = sheet_get_cell(sheet, 1, 0);
    TEST_ASSERT_EQ_DOUBLE(3.0, cell->data.number, 0.0001, "A2 should now be 3 (shifted from A3)");
    
    cell = sheet_get_cell(sheet, 2, 0);
    TEST_ASSERT_EQ_DOUBLE(4.0, cell->data.number, 0.0001, "A3 should now be 4 (shifted from A4)");
    
    sheet_free(sheet);
}

void test_insert_column(void) {
    TEST_SECTION("Insert Column");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Set up data
    sheet_set_number(sheet, 0, 0, 1.0);  // A1
    sheet_set_number(sheet, 0, 1, 2.0);  // B1
    sheet_set_number(sheet, 0, 2, 3.0);  // C1
    
    // Insert column at position 1
    sheet_insert_column(sheet, 1);
    
    // Verify data shifted right
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    TEST_ASSERT_EQ_DOUBLE(1.0, cell->data.number, 0.0001, "A1 should still be 1");
    
    cell = sheet_get_cell(sheet, 0, 1);
    TEST_ASSERT(cell == NULL || cell->type == CELL_EMPTY, "B1 (new column) should be empty");
    
    cell = sheet_get_cell(sheet, 0, 2);
    TEST_ASSERT_EQ_DOUBLE(2.0, cell->data.number, 0.0001, "C1 should now be 2 (shifted from B1)");
    
    cell = sheet_get_cell(sheet, 0, 3);
    TEST_ASSERT_EQ_DOUBLE(3.0, cell->data.number, 0.0001, "D1 should now be 3 (shifted from C1)");
    
    sheet_free(sheet);
}

void test_delete_column(void) {
    TEST_SECTION("Delete Column");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Set up data
    sheet_set_number(sheet, 0, 0, 1.0);  // A1
    sheet_set_number(sheet, 0, 1, 2.0);  // B1
    sheet_set_number(sheet, 0, 2, 3.0);  // C1
    sheet_set_number(sheet, 0, 3, 4.0);  // D1
    
    // Delete column 1 (B)
    sheet_delete_column(sheet, 1);
    
    // Verify data shifted left
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    TEST_ASSERT_EQ_DOUBLE(1.0, cell->data.number, 0.0001, "A1 should still be 1");
    
    cell = sheet_get_cell(sheet, 0, 1);
    TEST_ASSERT_EQ_DOUBLE(3.0, cell->data.number, 0.0001, "B1 should now be 3 (shifted from C1)");
    
    cell = sheet_get_cell(sheet, 0, 2);
    TEST_ASSERT_EQ_DOUBLE(4.0, cell->data.number, 0.0001, "C1 should now be 4 (shifted from D1)");
    
    sheet_free(sheet);
}

// ============================================================================
// CSV OPERATIONS TESTS
// ============================================================================

void test_csv_save_load_flatten(void) {
    TEST_SECTION("CSV Save/Load (Flatten Mode)");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Set up test data
    sheet_set_string(sheet, 0, 0, "Name");
    sheet_set_string(sheet, 0, 1, "Value");
    sheet_set_string(sheet, 1, 0, "Apple");
    sheet_set_number(sheet, 1, 1, 10.0);
    sheet_set_string(sheet, 2, 0, "Orange");
    sheet_set_number(sheet, 2, 1, 20.0);
    sheet_set_formula(sheet, 3, 1, "=B2+B3");  // Sum formula
    sheet_recalculate(sheet);
    
    // Save to CSV (flatten mode)
    const char* filename = "test_flatten.csv";
    int result = sheet_save_csv(sheet, filename, 0);  // 0 = flatten
    TEST_ASSERT(result, "CSV save should succeed");
    
    // Create a new sheet and load
    Sheet* loaded = sheet_new(100, 26);
    result = sheet_load_csv(loaded, filename, 0);
    TEST_ASSERT(result, "CSV load should succeed");
    
    // Verify loaded data
    char* value = sheet_get_display_value(loaded, 0, 0);
    TEST_ASSERT_EQ_STR("Name", value, "Loaded A1 should be 'Name'");
    
    value = sheet_get_display_value(loaded, 1, 0);
    TEST_ASSERT_EQ_STR("Apple", value, "Loaded A2 should be 'Apple'");
    
    Cell* cell = sheet_get_cell(loaded, 1, 1);
    TEST_ASSERT(cell != NULL, "Loaded B2 should exist");
    TEST_ASSERT_EQ_INT(CELL_NUMBER, cell->type, "Loaded B2 should be number");
    TEST_ASSERT_EQ_DOUBLE(10.0, cell->data.number, 0.0001, "Loaded B2 should be 10");
    
    // Formula should be loaded as value (flattened)
    cell = sheet_get_cell(loaded, 3, 1);
    TEST_ASSERT(cell != NULL, "Loaded B4 should exist");
    TEST_ASSERT_EQ_DOUBLE(30.0, cell->data.number, 0.0001, "Flattened formula should be 30");
    
    sheet_free(sheet);
    sheet_free(loaded);
    
    // Clean up test file
    remove(filename);
}

void test_csv_save_load_preserve(void) {
    TEST_SECTION("CSV Save/Load (Preserve Mode)");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Set up test data with formula
    sheet_set_number(sheet, 0, 0, 5.0);
    sheet_set_number(sheet, 1, 0, 10.0);
    sheet_set_formula(sheet, 2, 0, "=A1+A2");
    sheet_recalculate(sheet);
    
    // Save to CSV (preserve mode)
    const char* filename = "test_preserve.csv";
    int result = sheet_save_csv(sheet, filename, 1);  // 1 = preserve
    TEST_ASSERT(result, "CSV save with preserve should succeed");
    
    // Create a new sheet and load
    Sheet* loaded = sheet_new(100, 26);
    result = sheet_load_csv(loaded, filename, 1);
    TEST_ASSERT(result, "CSV load with preserve should succeed");
    
    // Verify formula is preserved
    Cell* cell = sheet_get_cell(loaded, 2, 0);
    TEST_ASSERT(cell != NULL, "Loaded A3 should exist");
    TEST_ASSERT_EQ_INT(CELL_FORMULA, cell->type, "Loaded A3 should be formula");
    
    // Recalculate and verify
    sheet_recalculate(loaded);
    TEST_ASSERT_EQ_DOUBLE(15.0, cell->data.formula.cached_value, 0.0001, "Preserved formula should evaluate to 15");
    
    sheet_free(sheet);
    sheet_free(loaded);
    
    // Clean up test file
    remove(filename);
}

void test_csv_special_characters(void) {
    TEST_SECTION("CSV Special Characters");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Test string with comma
    sheet_set_string(sheet, 0, 0, "Hello, World");
    
    // Test string with quote
    sheet_set_string(sheet, 1, 0, "He said \"Hello\"");
    
    // Test string with newline
    sheet_set_string(sheet, 2, 0, "Line1\nLine2");
    
    // Save and reload
    const char* filename = "test_special.csv";
    sheet_save_csv(sheet, filename, 0);
    
    Sheet* loaded = sheet_new(100, 26);
    sheet_load_csv(loaded, filename, 0);
    
    char* value = sheet_get_display_value(loaded, 0, 0);
    TEST_ASSERT_EQ_STR("Hello, World", value, "Comma should be preserved");
    
    value = sheet_get_display_value(loaded, 1, 0);
    TEST_ASSERT_EQ_STR("He said \"Hello\"", value, "Quotes should be preserved");
    
    sheet_free(sheet);
    sheet_free(loaded);
    remove(filename);
}

// ============================================================================
// DISPLAY VALUE TESTS
// ============================================================================

void test_display_values(void) {
    TEST_SECTION("Display Values");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Empty cell
    char* display = sheet_get_display_value(sheet, 0, 0);
    TEST_ASSERT_EQ_STR("", display, "Empty cell should display empty string");
    
    // Number
    sheet_set_number(sheet, 1, 0, 123.45);
    display = sheet_get_display_value(sheet, 1, 0);
    TEST_ASSERT(strstr(display, "123") != NULL, "Number should display");
    
    // String
    sheet_set_string(sheet, 2, 0, "Test");
    display = sheet_get_display_value(sheet, 2, 0);
    TEST_ASSERT_EQ_STR("Test", display, "String should display exactly");
    
    // Formula with result
    sheet_set_formula(sheet, 3, 0, "=1+1");
    sheet_recalculate(sheet);
    display = sheet_get_display_value(sheet, 3, 0);
    TEST_ASSERT(strstr(display, "2") != NULL, "Formula result should display");
    
    // Error display
    sheet_set_formula(sheet, 4, 0, "=1/0");
    sheet_recalculate(sheet);
    display = sheet_get_display_value(sheet, 4, 0);
    TEST_ASSERT_EQ_STR("#DIV/0!", display, "Division by zero should show #DIV/0!");
    
    sheet_free(sheet);
}

// ============================================================================
// ERROR HANDLING TESTS
// ============================================================================

void test_error_handling(void) {
    TEST_SECTION("Error Handling");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Division by zero
    sheet_set_formula(sheet, 0, 0, "=10/0");
    sheet_recalculate(sheet);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    TEST_ASSERT_EQ_INT(ERROR_DIV_ZERO, cell->data.formula.error, "Should be ERROR_DIV_ZERO");
    
    // Invalid formula syntax
    sheet_set_formula(sheet, 1, 0, "=1+");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 1, 0);
    TEST_ASSERT_EQ_INT(ERROR_PARSE, cell->data.formula.error, "Should be ERROR_PARSE");
    
    // XLOOKUP not found
    sheet_set_string(sheet, 2, 0, "Apple");
    sheet_set_number(sheet, 2, 1, 1.0);
    sheet_set_formula(sheet, 3, 0, "=XLOOKUP(\"Orange\", A3:A3, B3:B3, 0)");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 3, 0);
    TEST_ASSERT_EQ_INT(ERROR_NA, cell->data.formula.error, "XLOOKUP not found should be ERROR_NA");
    
    sheet_free(sheet);
}

// ============================================================================
// RECALCULATION TESTS
// ============================================================================

void test_recalculation(void) {
    TEST_SECTION("Recalculation");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Set up chain of formulas
    sheet_set_number(sheet, 0, 0, 10.0);    // A1 = 10
    sheet_set_formula(sheet, 1, 0, "=A1*2"); // A2 = A1 * 2
    sheet_set_formula(sheet, 2, 0, "=A2+5"); // A3 = A2 + 5
    
    sheet_recalculate(sheet);
    
    Cell* cell = sheet_get_cell(sheet, 1, 0);
    TEST_ASSERT_EQ_DOUBLE(20.0, cell->data.formula.cached_value, 0.0001, "A2 should be 20");
    
    cell = sheet_get_cell(sheet, 2, 0);
    TEST_ASSERT_EQ_DOUBLE(25.0, cell->data.formula.cached_value, 0.0001, "A3 should be 25");
    
    // Change source value
    sheet_set_number(sheet, 0, 0, 5.0);
    sheet_recalculate(sheet);
    
    cell = sheet_get_cell(sheet, 1, 0);
    TEST_ASSERT_EQ_DOUBLE(10.0, cell->data.formula.cached_value, 0.0001, "A2 should update to 10");
    
    cell = sheet_get_cell(sheet, 2, 0);
    TEST_ASSERT_EQ_DOUBLE(15.0, cell->data.formula.cached_value, 0.0001, "A3 should update to 15");
    
    sheet_free(sheet);
}

void test_recalc_after_clear(void) {
    TEST_SECTION("Recalculation After Clear");
    
    Sheet* sheet = sheet_new(100, 26);
    
    sheet_set_number(sheet, 0, 0, 100.0);
    sheet_set_formula(sheet, 1, 0, "=A1");
    sheet_recalculate(sheet);
    
    Cell* cell = sheet_get_cell(sheet, 1, 0);
    TEST_ASSERT_EQ_DOUBLE(100.0, cell->data.formula.cached_value, 0.0001, "A2 should be 100");
    
    // Clear source cell
    sheet_clear_cell(sheet, 0, 0);
    sheet_recalculate(sheet);
    
    cell = sheet_get_cell(sheet, 1, 0);
    TEST_ASSERT_EQ_DOUBLE(0.0, cell->data.formula.cached_value, 0.0001, "A2 should be 0 after clearing A1");
    
    sheet_free(sheet);
}

// ============================================================================
// EDGE CASE TESTS
// ============================================================================

void test_edge_cases(void) {
    TEST_SECTION("Edge Cases");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // Very long string
    char long_string[1024];
    memset(long_string, 'A', 1000);
    long_string[1000] = '\0';
    sheet_set_string(sheet, 0, 0, long_string);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    TEST_ASSERT(cell != NULL, "Cell with long string should exist");
    TEST_ASSERT_EQ_INT(CELL_STRING, cell->type, "Cell should be string type");
    
    // Very large number
    sheet_set_number(sheet, 1, 0, 1e308);
    cell = sheet_get_cell(sheet, 1, 0);
    TEST_ASSERT(cell->data.number > 1e307, "Very large number should be stored");
    
    // Very small number
    sheet_set_number(sheet, 2, 0, 1e-308);
    cell = sheet_get_cell(sheet, 2, 0);
    TEST_ASSERT(cell->data.number < 1e-307, "Very small number should be stored");
    
    // Negative zero
    sheet_set_number(sheet, 3, 0, -0.0);
    cell = sheet_get_cell(sheet, 3, 0);
    TEST_ASSERT_EQ_DOUBLE(0.0, cell->data.number, 0.0001, "Negative zero should be 0");
    
    // Empty formula
    sheet_set_formula(sheet, 4, 0, "=");
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 4, 0);
    // Should not crash, error is acceptable
    
    // Formula with only whitespace
    sheet_set_formula(sheet, 5, 0, "=   ");
    sheet_recalculate(sheet);
    // Should not crash
    
    sheet_free(sheet);
}

void test_boundary_cells(void) {
    TEST_SECTION("Boundary Cells");
    
    Sheet* sheet = sheet_new(100, 26);
    
    // First cell
    sheet_set_number(sheet, 0, 0, 1.0);
    Cell* cell = sheet_get_cell(sheet, 0, 0);
    TEST_ASSERT(cell != NULL, "First cell (0,0) should work");
    
    // Last cell in small sheet
    sheet_set_number(sheet, 99, 25, 999.0);
    cell = sheet_get_cell(sheet, 99, 25);
    TEST_ASSERT(cell != NULL, "Last cell (99,25) should work");
    TEST_ASSERT_EQ_DOUBLE(999.0, cell->data.number, 0.0001, "Last cell value should be correct");
    
    // Reference to boundary cell
    sheet_set_formula(sheet, 0, 1, "=A100");  // Reference to row 100 (index 99)
    sheet_recalculate(sheet);
    cell = sheet_get_cell(sheet, 0, 1);
    TEST_ASSERT_EQ_DOUBLE(0.0, cell->data.formula.cached_value, 0.0001, "Reference to empty boundary cell should be 0");
    
    sheet_free(sheet);
}

// ============================================================================
// PERFORMANCE TESTS (Simple)
// ============================================================================

void test_performance(void) {
    TEST_SECTION("Performance (Simple)");
    
    Sheet* sheet = sheet_new(1000, 100);
    
    // Fill many cells
    DWORD start = GetTickCount();
    for (int row = 0; row < 100; row++) {
        for (int col = 0; col < 26; col++) {
            sheet_set_number(sheet, row, col, (double)(row * col));
        }
    }
    DWORD elapsed = GetTickCount() - start;
    TEST_ASSERT(elapsed < 5000, "Setting 2600 cells should take < 5 seconds");
    printf("  INFO: Setting 2600 cells took %lu ms\n", elapsed);
    
    // Add many formulas
    start = GetTickCount();
    for (int row = 0; row < 50; row++) {
        sheet_set_formula(sheet, row, 26, "=SUM(A1:Z1)");
    }
    sheet_recalculate(sheet);
    elapsed = GetTickCount() - start;
    TEST_ASSERT(elapsed < 5000, "Adding 50 SUM formulas should take < 5 seconds");
    printf("  INFO: Adding and calculating 50 SUM formulas took %lu ms\n", elapsed);
    
    sheet_free(sheet);
}

// ============================================================================
// MAIN TEST RUNNER
// ============================================================================

int main(void) {
    printf("LiveLedger Unit Tests\n");
    printf("=====================\n");
    
    // Sheet Operations
    test_sheet_creation();
    test_sheet_get_cell();
    test_sheet_get_or_create_cell();
    
    // Cell Operations
    test_cell_creation();
    test_cell_set_number();
    test_cell_set_string();
    test_cell_set_formula();
    test_cell_clear();
    
    // Cell Reference Parsing
    test_cell_reference_parsing();
    test_cell_reference_to_string();
    
    // Formula Evaluation
    test_basic_arithmetic();
    test_cell_references_in_formulas();
    test_division_by_zero();
    
    // Built-in Functions
    test_sum_function();
    test_avg_function();
    test_max_min_functions();
    test_median_function();
    test_mode_function();
    test_if_function();
    test_power_function();
    test_xlookup_function();
    test_nested_functions();
    
    // Formatting
    test_percentage_format();
    test_currency_format();
    test_date_formats();
    test_time_formats();
    test_cell_colors();
    test_color_parsing();
    
    // Range Operations
    test_range_selection();
    test_range_copy_paste();
    test_clipboard_cell();
    
    // Column/Row Sizing
    test_column_width();
    test_row_height();
    
    // Insert/Delete Operations
    test_insert_row();
    test_delete_row();
    test_insert_column();
    test_delete_column();
    
    // CSV Operations
    test_csv_save_load_flatten();
    test_csv_save_load_preserve();
    test_csv_special_characters();
    
    // Display Values
    test_display_values();
    
    // Error Handling
    test_error_handling();
    
    // Recalculation
    test_recalculation();
    test_recalc_after_clear();
    
    // Edge Cases
    test_edge_cases();
    test_boundary_cells();
    
    // Performance
    test_performance();
    
    // Summary
    printf("\n=====================\n");
    printf("Test Results Summary\n");
    printf("=====================\n");
    printf("Total Tests:  %d\n", tests_run);
    printf("Passed:       %d\n", tests_passed);
    printf("Failed:       %d\n", tests_failed);
    printf("Pass Rate:    %.1f%%\n", (tests_run > 0) ? (100.0 * tests_passed / tests_run) : 0.0);
    
    if (tests_failed == 0) {
        printf("\n*** ALL TESTS PASSED ***\n");
    } else {
        printf("\n*** SOME TESTS FAILED ***\n");
    }
    
    return tests_failed;
}
