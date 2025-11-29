// charts.h - ASCII chart generation for LiveLedger terminal spreadsheet
#ifndef CHARTS_H
#define CHARTS_H

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <float.h>
#include "constants.h"
#include "sheet.h"
#include "console.h"

// Chart types
typedef enum {
    CHART_LINE,
    CHART_BAR,
    CHART_PIE,
    CHART_SCATTER
} ChartType;

// Chart configuration
typedef struct {
    ChartType type;
    char x_label[64];
    char y_label[64];
    char title[128];
    int width;
    int height;
    int show_grid;
    int show_legend;
} ChartConfig;

// Chart data point
typedef struct {
    double x;
    double y;
    char label[64];
} ChartPoint;

// Chart data series
typedef struct {
    ChartPoint* points;
    int count;
    char name[64];
    char symbol;  // Symbol to use for this series
} ChartSeries;

// Chart structure
typedef struct {
    ChartConfig config;
    ChartSeries* series;
    int series_count;
    double x_min, x_max;
    double y_min, y_max;
    char** canvas;  // ASCII canvas for drawing
    int canvas_width;
    int canvas_height;
} Chart;

// Function prototypes
Chart* chart_create(ChartType type, const char* x_label, const char* y_label);
Chart* chart_create_sized(ChartType type, const char* x_label, const char* y_label, int width, int height);
void chart_free(Chart* chart);
int chart_add_data_from_range(Chart* chart, Sheet* sheet, RangeSelection* range);
void chart_render(Chart* chart);
void chart_display(Chart* chart, Console* console, int x, int y);
char** chart_get_output(Chart* chart, int* line_count);

// Helper functions
void chart_draw_line(Chart* chart, int x1, int y1, int x2, int y2, char symbol);
void chart_draw_axes(Chart* chart);
void chart_plot_line_chart(Chart* chart);
void chart_plot_bar_chart(Chart* chart);
void chart_plot_pie_chart(Chart* chart);
void chart_plot_scatter_chart(Chart* chart);
void chart_set_pixel(Chart* chart, int x, int y, char c);
int chart_scale_x(Chart* chart, double value);
int chart_scale_y(Chart* chart, double value);

#endif // CHARTS_H
