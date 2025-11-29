// console.h - Windows Console API wrapper for terminal spreadsheet
#ifndef CONSOLE_H
#define CONSOLE_H

#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

// Include debug logging
extern void debug_log(const char* format, ...);
extern FILE* debug_file;

// Console color attributes
#define COLOR_BLACK         0
#define COLOR_BLUE          1
#define COLOR_GREEN         2
#define COLOR_CYAN          3
#define COLOR_RED           4
#define COLOR_MAGENTA       5
#define COLOR_YELLOW        6
#define COLOR_WHITE         7
#define COLOR_BRIGHT        8

// Combine foreground and background colors
#define MAKE_COLOR(fg, bg)  ((bg << 4) | fg)

// Special keys
#define KEY_UP              0x48
#define KEY_DOWN            0x50
#define KEY_LEFT            0x4B
#define KEY_RIGHT           0x4D
#define KEY_PGUP            0x49
#define KEY_PGDN            0x51
#define KEY_HOME            0x47
#define KEY_END             0x4F
#define KEY_DELETE          0x53
#define KEY_F1              0x3B
#define KEY_ESC             0x1B
#define KEY_ENTER           0x0D
#define KEY_BACKSPACE       0x08
#define KEY_TAB             0x09

typedef struct {
    HANDLE hOut;
    HANDLE hIn;
    CONSOLE_SCREEN_BUFFER_INFO originalInfo;
    CHAR_INFO* backBuffer;
    CHAR_INFO* frontBuffer;
    SHORT width;
    SHORT height;
    DWORD originalMode;
} Console;

typedef struct {
    int type;  // 0 = char, 1 = special key
    union {
        char ch;
        int special;
    } key;
    BOOL ctrl;
    BOOL alt;
    BOOL shift;
} KeyEvent;

// Function prototypes
Console* console_init(void);
void console_cleanup(Console* con);
void console_clear(Console* con);
void console_set_cursor(Console* con, SHORT x, SHORT y);
void console_hide_cursor(Console* con);
void console_show_cursor(Console* con);
void console_write_char(Console* con, SHORT x, SHORT y, char ch, WORD attr);
void console_write_string(Console* con, SHORT x, SHORT y, const char* str, WORD attr);
void console_draw_box(Console* con, SHORT x, SHORT y, SHORT w, SHORT h, WORD attr);
void console_flip(Console* con);
BOOL console_get_key(Console* con, KeyEvent* key);
void console_get_size(Console* con, SHORT* width, SHORT* height);

#endif // CONSOLE_H
