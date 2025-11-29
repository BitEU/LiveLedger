// console.c - Windows Console API implementation
#include "console.h"


Console* console_init(void) {
    Console* con = (Console*)malloc(sizeof(Console));
    if (!con) return NULL;
    
    // Get console handles
    con->hOut = GetStdHandle(STD_OUTPUT_HANDLE);
    con->hIn = GetStdHandle(STD_INPUT_HANDLE);
    
    // Check if handles are valid
    if (con->hOut == INVALID_HANDLE_VALUE || con->hIn == INVALID_HANDLE_VALUE) {
        free(con);
        return NULL;
    }
    
    // Save original console state
    if (!GetConsoleScreenBufferInfo(con->hOut, &con->originalInfo)) {
        free(con);
        return NULL;
    }
    
    if (!GetConsoleMode(con->hIn, &con->originalMode)) {
        free(con);
        return NULL;
    }
    
    // Set input mode for reading individual keys
    SetConsoleMode(con->hIn, ENABLE_WINDOW_INPUT | ENABLE_MOUSE_INPUT);
    
    // Get console size
    CONSOLE_SCREEN_BUFFER_INFO csbi;
    GetConsoleScreenBufferInfo(con->hOut, &csbi);
    con->width = csbi.srWindow.Right - csbi.srWindow.Left + 1;
    con->height = csbi.srWindow.Bottom - csbi.srWindow.Top + 1;
    
    // Validate console size
    if (con->width <= 0 || con->height <= 0) {
        free(con);
        return NULL;
    }
    
    // Allocate double buffer
    int bufferSize = con->width * con->height;
    con->backBuffer = (CHAR_INFO*)malloc(bufferSize * sizeof(CHAR_INFO));
    con->frontBuffer = (CHAR_INFO*)malloc(bufferSize * sizeof(CHAR_INFO));
    
    // Check if allocation succeeded
    if (!con->backBuffer || !con->frontBuffer) {
        if (con->backBuffer) free(con->backBuffer);
        if (con->frontBuffer) free(con->frontBuffer);
        free(con);
        return NULL;
    }
    
    // Initialize buffers
    for (int i = 0; i < bufferSize; i++) {
        con->backBuffer[i].Char.AsciiChar = ' ';
        con->backBuffer[i].Attributes = MAKE_COLOR(COLOR_WHITE, COLOR_BLACK);
        con->frontBuffer[i] = con->backBuffer[i];
    }
    
    // Clear screen
    console_clear(con);
    
    return con;
}

void console_cleanup(Console* con) {
    if (!con) return;
    
    // Restore original console state
    SetConsoleTextAttribute(con->hOut, con->originalInfo.wAttributes);
    SetConsoleMode(con->hIn, con->originalMode);
    
    // Free buffers
    if (con->backBuffer) {
        free(con->backBuffer);
        con->backBuffer = NULL;
    }
    if (con->frontBuffer) {
        free(con->frontBuffer);
        con->frontBuffer = NULL;
    }
    free(con);
}

void console_clear(Console* con) {
    COORD topLeft = {0, 0};
    DWORD written;
    FillConsoleOutputCharacter(con->hOut, ' ', con->width * con->height, topLeft, &written);
    FillConsoleOutputAttribute(con->hOut, MAKE_COLOR(COLOR_WHITE, COLOR_BLACK), 
                               con->width * con->height, topLeft, &written);
    console_set_cursor(con, 0, 0);
}

void console_set_cursor(Console* con, SHORT x, SHORT y) {
    COORD pos = {x, y};
    SetConsoleCursorPosition(con->hOut, pos);
}

void console_hide_cursor(Console* con) {
    CONSOLE_CURSOR_INFO cursorInfo;
    GetConsoleCursorInfo(con->hOut, &cursorInfo);
    cursorInfo.bVisible = FALSE;
    SetConsoleCursorInfo(con->hOut, &cursorInfo);
}

void console_show_cursor(Console* con) {
    CONSOLE_CURSOR_INFO cursorInfo;
    GetConsoleCursorInfo(con->hOut, &cursorInfo);
    cursorInfo.bVisible = TRUE;
    SetConsoleCursorInfo(con->hOut, &cursorInfo);
}

void console_write_char(Console* con, SHORT x, SHORT y, char ch, WORD attr) {
    if (x >= 0 && x < con->width && y >= 0 && y < con->height) {
        int index = y * con->width + x;
        con->backBuffer[index].Char.AsciiChar = ch;
        con->backBuffer[index].Attributes = attr;
    }
}

void console_write_string(Console* con, SHORT x, SHORT y, const char* str, WORD attr) {
    int len = (int)strlen(str);
    for (int i = 0; i < len && x + i < con->width; i++) {
        console_write_char(con, x + i, y, str[i], attr);
    }
}

void console_draw_box(Console* con, SHORT x, SHORT y, SHORT w, SHORT h, WORD attr) {
    // Draw corners
    console_write_char(con, x, y, '+', attr);
    console_write_char(con, x + w - 1, y, '+', attr);
    console_write_char(con, x, y + h - 1, '+', attr);
    console_write_char(con, x + w - 1, y + h - 1, '+', attr);
    
    // Draw horizontal lines
    for (SHORT i = 1; i < w - 1; i++) {
        console_write_char(con, x + i, y, '-', attr);
        console_write_char(con, x + i, y + h - 1, '-', attr);
    }
    
    // Draw vertical lines
    for (SHORT i = 1; i < h - 1; i++) {
        console_write_char(con, x, y + i, '|', attr);
        console_write_char(con, x + w - 1, y + i, '|', attr);
    }
}

void console_flip(Console* con) {
    if (debug_file) debug_log("console_flip called");
    
    if (!con) {
        if (debug_file) debug_log("ERROR: console_flip - con is NULL");
        return;
    }
    
    if (!con->backBuffer) {
        if (debug_file) debug_log("ERROR: console_flip - backBuffer is NULL");
        return;
    }
    
    if (debug_file) debug_log("console_flip - calling WriteConsoleOutput, size: %dx%d", con->width, con->height);
    
    // Only update changed characters (optimization)
    COORD bufferSize = {con->width, con->height};
    COORD bufferCoord = {0, 0};
    SMALL_RECT writeRegion = {0, 0, con->width - 1, con->height - 1};
    
    BOOL result = WriteConsoleOutput(con->hOut, con->backBuffer, bufferSize, bufferCoord, &writeRegion);
    if (!result) {
        DWORD error = GetLastError();
        if (debug_file) debug_log("ERROR: WriteConsoleOutput failed with error: %lu", error);
        return;
    }
    
    if (debug_file) debug_log("WriteConsoleOutput succeeded, copying buffers");
    // Copy back buffer to front buffer
    memcpy(con->frontBuffer, con->backBuffer, con->width * con->height * sizeof(CHAR_INFO));
    
    if (debug_file) debug_log("console_flip completed successfully");
}

BOOL console_get_key(Console* con, KeyEvent* key) {
    INPUT_RECORD inputRecord;
    DWORD eventsRead;
    
    // Check if input is available
    DWORD numEvents;
    GetNumberOfConsoleInputEvents(con->hIn, &numEvents);
    if (numEvents == 0) return FALSE;
    
    // Read input
    if (ReadConsoleInput(con->hIn, &inputRecord, 1, &eventsRead)) {
        if (inputRecord.EventType == KEY_EVENT && inputRecord.Event.KeyEvent.bKeyDown) {
            KEY_EVENT_RECORD* keyEvent = &inputRecord.Event.KeyEvent;
            
            // Set modifiers
            key->ctrl = (keyEvent->dwControlKeyState & (LEFT_CTRL_PRESSED | RIGHT_CTRL_PRESSED)) != 0;
            key->alt = (keyEvent->dwControlKeyState & (LEFT_ALT_PRESSED | RIGHT_ALT_PRESSED)) != 0;
            key->shift = (keyEvent->dwControlKeyState & SHIFT_PRESSED) != 0;
            
            // Check for special keys
            if (keyEvent->wVirtualKeyCode == VK_UP) {
                key->type = 1;
                key->key.special = KEY_UP;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode == VK_DOWN) {
                key->type = 1;
                key->key.special = KEY_DOWN;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode == VK_LEFT) {
                key->type = 1;
                key->key.special = KEY_LEFT;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode == VK_RIGHT) {
                key->type = 1;
                key->key.special = KEY_RIGHT;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode == VK_PRIOR) {
                key->type = 1;
                key->key.special = KEY_PGUP;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode == VK_NEXT) {
                key->type = 1;
                key->key.special = KEY_PGDN;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode == VK_HOME) {
                key->type = 1;
                key->key.special = KEY_HOME;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode == VK_END) {
                key->type = 1;
                key->key.special = KEY_END;
                return TRUE;            } else if (keyEvent->wVirtualKeyCode == VK_F1) {
                key->type = 1;
                key->key.special = KEY_F1;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode >= 'A' && keyEvent->wVirtualKeyCode <= 'Z' && key->ctrl) {
                // Handle Ctrl+letter combinations (with or without Shift)
                key->type = 0;
                key->key.ch = (char)keyEvent->wVirtualKeyCode + ('a' - 'A'); // Convert to lowercase
                return TRUE;
            } else if (key->ctrl && key->shift) {
                // Handle Ctrl+Shift+number combinations for formatting
                switch (keyEvent->wVirtualKeyCode) {
                    case '1':
                        key->type = 0;
                        key->key.ch = '1';
                        return TRUE;
                    case '3':
                        key->type = 0;
                        key->key.ch = '3';
                        return TRUE;
                    case '4':
                        key->type = 0;
                        key->key.ch = '4';
                        return TRUE;
                    case '5':
                        key->type = 0;
                        key->key.ch = '5';
                        return TRUE;
                }
            } else if (key->ctrl) {
                // Handle Ctrl+symbol combinations
                switch (keyEvent->wVirtualKeyCode) {
                    case '5':  // Ctrl+5 for %
                        if (key->shift) {
                            key->type = 0;
                            key->key.ch = '%';
                            return TRUE;
                        }
                        break;
                    case '4':  // Ctrl+4 for $
                        if (key->shift) {
                            key->type = 0;
                            key->key.ch = '$';
                            return TRUE;
                        }
                        break;
                    case '3':  // Ctrl+3 for #
                        if (key->shift) {
                            key->type = 0;
                            key->key.ch = '#';
                            return TRUE;
                        }
                        break;
                }
                // Also handle when the character is already the symbol
                if (keyEvent->uChar.AsciiChar != 0) {
                    key->type = 0;
                    key->key.ch = keyEvent->uChar.AsciiChar;
                    return TRUE;
                }
            } else if (keyEvent->uChar.AsciiChar != 0) {
                // Regular character
                key->type = 0;
                key->key.ch = keyEvent->uChar.AsciiChar;
                return TRUE;
            }
        }
    }
    
    return FALSE;
}

void console_get_size(Console* con, SHORT* width, SHORT* height) {
    CONSOLE_SCREEN_BUFFER_INFO csbi;
    GetConsoleScreenBufferInfo(con->hOut, &csbi);
    *width = csbi.srWindow.Right - csbi.srWindow.Left + 1;
    *height = csbi.srWindow.Bottom - csbi.srWindow.Top + 1;
}