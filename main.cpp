#define UNICODE
#define _UNICODE
#include <locale>
#include <codecvt>
#include <richedit.h>
#include <ctime>
#include <sstream>
#include <iomanip>
#include <fstream>
#include <Richedit.h>
#include <windows.h>
#include <commctrl.h>
#include <string>
#include <vector>
#include <thread>
#include <mutex>
#include <atomic>

// Global variables
HWND hWndMain;
HWND hWndEdit;
HWND hWndButton;
HWND hWndCombo;
HWND hWndStopBtn;
HWND hWndExportBtn;
HWND hWndClearBtn;
HWND hWndPauseBtn;
HWND hWndTooltip;
HWND hWndColorBtns[10];
HWND hWndBaudCombo;
HWND hWndStatusBar;
HICON hPlayIcon;
HICON hStopIcon;
DWORD currentBaud = 115200;
std::vector<HANDLE> hSerialPorts;
std::vector<std::thread> readThreads;
std::vector<int> portIndices;
std::atomic<bool> running(true);
bool isListening = true;
bool paused = false;
std::mutex textMutex;
std::vector<COLORREF> portColors;

// Function declarations
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
DWORD CALLBACK EditStreamCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb);
void CreateMainWindow(HINSTANCE hInstance);
void SetupSerialPorts(bool startThreads = true);
void ReadFromPort(HANDLE hPort, int portIndex);
void AppendTextToEdit(const std::wstring& text, int portIndex = 0);
void RefreshPorts();

DWORD CALLBACK EditStreamCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb) {
    std::ofstream* file = (std::ofstream*)dwCookie;
    file->write((char*)pbBuff, cb);
    *pcb = cb;
    return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    // Register window class
    WNDCLASSEX wc = {0};
    wc.cbSize = sizeof(WNDCLASSEX);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = L"SerialMonitorClass";
    wc.hIcon = LoadIcon(GetModuleHandle(NULL), MAKEINTRESOURCE(2000));
    wc.hIconSm = LoadIcon(GetModuleHandle(NULL), MAKEINTRESOURCE(2000));

    if (!RegisterClassEx(&wc)) {
        MessageBox(NULL, L"Window Registration Failed!", L"Error!", MB_ICONEXCLAMATION | MB_OK);
        return 0;
    }

    // Initialize common controls
    InitCommonControls();

    // Initialize port colors
    portColors = {
        RGB(255, 255, 255), // 0: white
        RGB(150, 255, 150), // 1: light green
        RGB(255, 200, 100), // 2: light orange
        RGB(150, 150, 255), // 3: light blue
        RGB(255, 255, 150), // 4: light yellow
        RGB(200, 150, 255), // 5: light purple
        RGB(150, 255, 200), // 6: mint green
        RGB(255, 255, 255), // 7: white
        RGB(255, 220, 220), // 8: very light red
        RGB(255, 180, 180), // 9: light red
        RGB(255, 140, 140), // 10: red
        RGB(255, 100, 100), // 11: dark red
        RGB(255, 230, 180), // 12: very light orange
        RGB(255, 200, 140), // 13: light orange
        RGB(255, 170, 100), // 14: orange
        RGB(255, 140, 60),  // 15: dark orange
        RGB(255, 255, 200), // 16: very light yellow
        RGB(255, 255, 150), // 17: light yellow
        RGB(255, 255, 100), // 18: yellow
        RGB(255, 255, 50),  // 19: dark yellow
        RGB(255, 220, 230)  // 20: very light pink
    };
    // Extend to 256 with the last color
    portColors.resize(256, portColors.back());

    // Create main window
    CreateMainWindow(hInstance);

    // Setup serial ports but don't start listening
    SetupSerialPorts(false);
    AppendTextToEdit(L"Multi Serial Monitor started. Ready to monitor serial ports at 115200 baud.\r\n", 0);

    // Set initial state to not listening
    isListening = false;
    SendMessage(hWndStopBtn, BM_SETIMAGE, IMAGE_ICON, (LPARAM)hPlayIcon);
    // Update status bar
    std::wstring statusText = L"Ready - " + std::to_wstring(hSerialPorts.size()) + L" ports available";
    SendMessage(hWndStatusBar, SB_SETTEXT, 0, (LPARAM)statusText.c_str());
    // Baud combo is enabled since not listening

    // Show window
    ShowWindow(hWndMain, nCmdShow);
    UpdateWindow(hWndMain);

    // Message loop
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // Cleanup
    running = false;
    for (auto& t : readThreads) {
        if (t.joinable()) t.join();
    }
    for (auto h : hSerialPorts) {
        if (h != INVALID_HANDLE_VALUE) CloseHandle(h);
    }

    return msg.wParam;
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_CREATE:
            {
                // Create rich edit control
                LoadLibrary(L"Msftedit.dll");
                hWndEdit = CreateWindowEx(WS_EX_CLIENTEDGE, MSFTEDIT_CLASS, L"",
                    WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
                    0, 40, 800, 570, hwnd, NULL, GetModuleHandle(NULL), NULL);
                // Set background color
                SendMessage(hWndEdit, EM_SETBKGNDCOLOR, 0, RGB(64, 64, 64));
                // Create button
                hWndButton = CreateWindowEx(0, L"BUTTON", L"",
                    WS_CHILD | WS_VISIBLE | BS_ICON,
                    0, 0, 40, 40, hwnd, (HMENU)1, GetModuleHandle(NULL), NULL);
                // Set refresh icon
                HICON hRefreshIcon = LoadIcon(GetModuleHandle(NULL), MAKEINTRESOURCE(2006));
                SendMessage(hWndButton, BM_SETIMAGE, IMAGE_ICON, (LPARAM)hRefreshIcon);
                // Create combo box
                hWndCombo = CreateWindowEx(0, L"COMBOBOX", NULL,
                    WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
                    45, 0, 100, 200, hwnd, (HMENU)2, GetModuleHandle(NULL), NULL);
                // Create baud rate combo box
                hWndBaudCombo = CreateWindowEx(0, L"COMBOBOX", NULL,
                    WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
                    750, 0, 80, 200, hwnd, (HMENU)3, GetModuleHandle(NULL), NULL);
                // Add baud rates
                SendMessage(hWndBaudCombo, CB_ADDSTRING, 0, (LPARAM)L"1200");
                SendMessage(hWndBaudCombo, CB_SETITEMDATA, 0, 1200);
                SendMessage(hWndBaudCombo, CB_ADDSTRING, 0, (LPARAM)L"2400");
                SendMessage(hWndBaudCombo, CB_SETITEMDATA, 1, 2400);
                SendMessage(hWndBaudCombo, CB_ADDSTRING, 0, (LPARAM)L"4800");
                SendMessage(hWndBaudCombo, CB_SETITEMDATA, 2, 4800);
                SendMessage(hWndBaudCombo, CB_ADDSTRING, 0, (LPARAM)L"9600");
                SendMessage(hWndBaudCombo, CB_SETITEMDATA, 3, 9600);
                SendMessage(hWndBaudCombo, CB_ADDSTRING, 0, (LPARAM)L"14400");
                SendMessage(hWndBaudCombo, CB_SETITEMDATA, 4, 14400);
                SendMessage(hWndBaudCombo, CB_ADDSTRING, 0, (LPARAM)L"19200");
                SendMessage(hWndBaudCombo, CB_SETITEMDATA, 5, 19200);
                SendMessage(hWndBaudCombo, CB_ADDSTRING, 0, (LPARAM)L"28800");
                SendMessage(hWndBaudCombo, CB_SETITEMDATA, 6, 28800);
                SendMessage(hWndBaudCombo, CB_ADDSTRING, 0, (LPARAM)L"38400");
                SendMessage(hWndBaudCombo, CB_SETITEMDATA, 7, 38400);
                SendMessage(hWndBaudCombo, CB_ADDSTRING, 0, (LPARAM)L"57600");
                SendMessage(hWndBaudCombo, CB_SETITEMDATA, 8, 57600);
                SendMessage(hWndBaudCombo, CB_ADDSTRING, 0, (LPARAM)L"115200");
                SendMessage(hWndBaudCombo, CB_SETITEMDATA, 9, 115200);
                SendMessage(hWndBaudCombo, CB_ADDSTRING, 0, (LPARAM)L"230400");
                SendMessage(hWndBaudCombo, CB_SETITEMDATA, 10, 230400);
                SendMessage(hWndBaudCombo, CB_ADDSTRING, 0, (LPARAM)L"460800");
                SendMessage(hWndBaudCombo, CB_SETITEMDATA, 11, 460800);
                SendMessage(hWndBaudCombo, CB_ADDSTRING, 0, (LPARAM)L"921600");
                SendMessage(hWndBaudCombo, CB_SETITEMDATA, 12, 921600);
                // Set default to 115200 (index 9)
                SendMessage(hWndBaudCombo, CB_SETCURSEL, 9, 0);
                // Create color palette buttons
                for(int i = 0; i < 10; i++) {
                    hWndColorBtns[i] = CreateWindowEx(0, L"BUTTON", L"",
                        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
                        150 + i * 42, 0, 40, 40, hwnd, reinterpret_cast<HMENU>(10 + i), GetModuleHandle(NULL), NULL);
                }
                // Load play and stop icons
                hPlayIcon = LoadIcon(GetModuleHandle(NULL), MAKEINTRESOURCE(2004));
                hStopIcon = LoadIcon(GetModuleHandle(NULL), MAKEINTRESOURCE(2005));
                // Create stop/start button
                hWndStopBtn = CreateWindowEx(0, L"BUTTON", L"",
                    WS_CHILD | WS_VISIBLE | BS_ICON,
                    570, 0, 40, 40, hwnd, (HMENU)4, GetModuleHandle(NULL), NULL);
                // Set stop icon initially
                SendMessage(hWndStopBtn, BM_SETIMAGE, IMAGE_ICON, (LPARAM)hStopIcon);
                // Create export button
                hWndExportBtn = CreateWindowEx(0, L"BUTTON", L"",
                    WS_CHILD | WS_VISIBLE | BS_ICON,
                    615, 0, 40, 40, hwnd, (HMENU)5, GetModuleHandle(NULL), NULL);
                // Set export icon
                HICON hExportIcon = LoadIcon(GetModuleHandle(NULL), MAKEINTRESOURCE(2003));
                SendMessage(hWndExportBtn, BM_SETIMAGE, IMAGE_ICON, (LPARAM)hExportIcon);
                // Create clear button
                hWndClearBtn = CreateWindowEx(0, L"BUTTON", L"",
                    WS_CHILD | WS_VISIBLE | BS_ICON,
                    660, 0, 40, 40, hwnd, (HMENU)7, GetModuleHandle(NULL), NULL);
                // Set clear icon
                HICON hClearIcon = LoadIcon(GetModuleHandle(NULL), MAKEINTRESOURCE(2002));
                SendMessage(hWndClearBtn, BM_SETIMAGE, IMAGE_ICON, (LPARAM)hClearIcon);
                // Create pause button
                hWndPauseBtn = CreateWindowEx(0, L"BUTTON", L"",
                    WS_CHILD | WS_VISIBLE | BS_ICON,
                    705, 0, 40, 40, hwnd, (HMENU)6, GetModuleHandle(NULL), NULL);
                // Set pause icon
                HICON hPauseIcon = LoadIcon(GetModuleHandle(NULL), MAKEINTRESOURCE(2001));
                SendMessage(hWndPauseBtn, BM_SETIMAGE, IMAGE_ICON, (LPARAM)hPauseIcon);
                // Create tooltip
                hWndTooltip = CreateWindowEx(0, TOOLTIPS_CLASS, NULL,
                    WS_POPUP | TTS_ALWAYSTIP,
                    CW_USEDEFAULT, CW_USEDEFAULT,
                    CW_USEDEFAULT, CW_USEDEFAULT,
                    hwnd, NULL, GetModuleHandle(NULL), NULL);
                // Add tools
                TOOLINFO ti = {0};
                ti.cbSize = sizeof(TOOLINFO);
                ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
                ti.hwnd = hwnd;
                ti.uId = (UINT_PTR)hWndButton;
                ti.lpszText = const_cast<LPWSTR>(L"Refresh Ports");
                SendMessage(hWndTooltip, TTM_ADDTOOL, 0, (LPARAM)&ti);
                ti.uId = (UINT_PTR)hWndStopBtn;
                ti.lpszText = const_cast<LPWSTR>(L"Stop/Start Listening");
                SendMessage(hWndTooltip, TTM_ADDTOOL, 0, (LPARAM)&ti);
                ti.uId = (UINT_PTR)hWndExportBtn;
                ti.lpszText = const_cast<LPWSTR>(L"Export to HTML");
                SendMessage(hWndTooltip, TTM_ADDTOOL, 0, (LPARAM)&ti);
                ti.uId = (UINT_PTR)hWndClearBtn;
                ti.lpszText = const_cast<LPWSTR>(L"Clear Display");
                SendMessage(hWndTooltip, TTM_ADDTOOL, 0, (LPARAM)&ti);
                ti.uId = (UINT_PTR)hWndPauseBtn;
                ti.lpszText = const_cast<LPWSTR>(L"Pause/Resume Scroll");
                SendMessage(hWndTooltip, TTM_ADDTOOL, 0, (LPARAM)&ti);
                ti.uId = (UINT_PTR)hWndBaudCombo;
                ti.lpszText = const_cast<LPWSTR>(L"Select Baud Rate");
                SendMessage(hWndTooltip, TTM_ADDTOOL, 0, (LPARAM)&ti);
                ti.uId = (UINT_PTR)hWndCombo;
                ti.lpszText = const_cast<LPWSTR>(L"Select Serial Port");
                SendMessage(hWndTooltip, TTM_ADDTOOL, 0, (LPARAM)&ti);
                // Create status bar
                hWndStatusBar = CreateWindowEx(0, STATUSCLASSNAME, NULL,
                    WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP, 0, 0, 0, 0,
                    hwnd, (HMENU)100, GetModuleHandle(NULL), NULL);
                // Add tooltips for color buttons
                for(int i = 0; i < 10; i++) {
                    ti.uId = (UINT_PTR)hWndColorBtns[i];
                    std::wstring colorText = L"Color " + std::to_wstring(i + 1);
                    ti.lpszText = const_cast<LPWSTR>(colorText.c_str());
                    SendMessage(hWndTooltip, TTM_ADDTOOL, 0, (LPARAM)&ti);
                }
            }
            break;
        case WM_SIZE:
            {
                RECT rcClient;
                GetClientRect(hwnd, &rcClient);
                int statusHeight = 25;
                // Resize edit control and buttons
                MoveWindow(hWndEdit, 0, 40, LOWORD(lParam), HIWORD(lParam) - 40 - statusHeight, TRUE);
                MoveWindow(hWndButton, 0, 0, 40, 40, TRUE);
                MoveWindow(hWndCombo, 45, 0, 100, 200, TRUE);
                MoveWindow(hWndBaudCombo, 750, 0, 80, 200, TRUE);
                MoveWindow(hWndStopBtn, 570, 0, 40, 40, TRUE);
                MoveWindow(hWndExportBtn, 615, 0, 40, 40, TRUE);
                MoveWindow(hWndClearBtn, 660, 0, 40, 40, TRUE);
                MoveWindow(hWndPauseBtn, 705, 0, 40, 40, TRUE);
                // Position status bar
                MoveWindow(hWndStatusBar, 0, HIWORD(lParam) - statusHeight, LOWORD(lParam), statusHeight, TRUE);
            }
            break;
        case WM_DRAWITEM:
            {
                DRAWITEMSTRUCT* dis = (DRAWITEMSTRUCT*)lParam;
                if (dis->CtlID >= 10 && dis->CtlID < 20) {
                    int colorIndex = dis->CtlID - 9; // 1 to 10
                    HBRUSH hBrush = CreateSolidBrush(portColors[colorIndex]);
                    FillRect(dis->hDC, &dis->rcItem, hBrush);
                    DeleteObject(hBrush);
                }
            }
            break;
        case WM_NOTIFY:
            {
                LPNMHDR pnmh = (LPNMHDR)lParam;
                if (pnmh->code == TTN_SHOW) {
                    LPNMTTDISPINFO pInfo = (LPNMTTDISPINFO)lParam;
                    SendMessage(hWndStatusBar, SB_SETTEXT, 0, (LPARAM)pInfo->lpszText);
                } else if (pnmh->code == TTN_POP) {
                    std::wstring statusText;
                    if (isListening) {
                        statusText = L"Monitoring " + std::to_wstring(hSerialPorts.size()) + L" ports";
                    } else {
                        statusText = L"Ready - " + std::to_wstring(hSerialPorts.size()) + L" ports available";
                    }
                    SendMessage(hWndStatusBar, SB_SETTEXT, 0, (LPARAM)statusText.c_str());
                }
            }
            break;
        case WM_COMMAND:
            if (LOWORD(wParam) == 1) {
                // Refresh button clicked
                // Disable all controls
                EnableWindow(hWndButton, FALSE);
                EnableWindow(hWndStopBtn, FALSE);
                EnableWindow(hWndExportBtn, FALSE);
                EnableWindow(hWndClearBtn, FALSE);
                EnableWindow(hWndPauseBtn, FALSE);
                for(int i = 0; i < 10; i++) EnableWindow(hWndColorBtns[i], FALSE);
                EnableWindow(hWndCombo, FALSE);
                EnableWindow(hWndBaudCombo, FALSE);
                // Update status
                SendMessage(hWndStatusBar, SB_SETTEXT, 0, (LPARAM)L"Refreshing ports...");
                // Create progress bar
                RECT rcStatus;
                GetWindowRect(hWndStatusBar, &rcStatus);
                ScreenToClient(hwnd, (LPPOINT)&rcStatus.left);
                ScreenToClient(hwnd, (LPPOINT)&rcStatus.right);
                int progressX = rcStatus.left + 200;
                int progressWidth = rcStatus.right - rcStatus.left - 200;
                int progressHeight = rcStatus.bottom - rcStatus.top - 4;
                HWND hProgress = CreateWindowEx(0, PROGRESS_CLASS, NULL, WS_CHILD | WS_VISIBLE, progressX, rcStatus.top + 2, progressWidth, progressHeight, hwnd, NULL, GetModuleHandle(NULL), NULL);
                SendMessage(hProgress, PBM_SETRANGE, 0, MAKELPARAM(0, 100));
                SendMessage(hProgress, PBM_SETPOS, 0, 0);
                // Step 1: stop running
                SendMessage(hProgress, PBM_SETPOS, 10, 0);
                running = false;
                for (auto& t : readThreads) {
                    if (t.joinable()) t.join();
                }
                // Step 2: close handles
                SendMessage(hProgress, PBM_SETPOS, 30, 0);
                for (auto h : hSerialPorts) {
                    if (h != INVALID_HANDLE_VALUE) CloseHandle(h);
                }
                // Step 3: clear
                SendMessage(hProgress, PBM_SETPOS, 50, 0);
                hSerialPorts.clear();
                readThreads.clear();
                portIndices.clear();
                SendMessage(hWndCombo, CB_RESETCONTENT, 0, 0);
                // Step 4: restart
                SendMessage(hProgress, PBM_SETPOS, 70, 0);
                running = true;
                SetupSerialPorts(true);
                // Step 5: done
                SendMessage(hProgress, PBM_SETPOS, 100, 0);
                Sleep(200); // brief pause to show 100%
                // Destroy progress bar
                DestroyWindow(hProgress);
                // Enable all controls
                EnableWindow(hWndButton, TRUE);
                EnableWindow(hWndStopBtn, TRUE);
                EnableWindow(hWndExportBtn, TRUE);
                EnableWindow(hWndClearBtn, TRUE);
                EnableWindow(hWndPauseBtn, TRUE);
                for(int i = 0; i < 10; i++) EnableWindow(hWndColorBtns[i], TRUE);
                EnableWindow(hWndCombo, TRUE);
                EnableWindow(hWndBaudCombo, TRUE);
                // Update status
                std::wstring statusText = isListening ? L"Monitoring " + std::to_wstring(hSerialPorts.size()) + L" ports" : L"Ready - " + std::to_wstring(hSerialPorts.size()) + L" ports available";
                SendMessage(hWndStatusBar, SB_SETTEXT, 0, (LPARAM)statusText.c_str());
            } else if (LOWORD(wParam) >= 10 && LOWORD(wParam) < 20) {
                // Color palette button clicked
                int sel = SendMessage(hWndCombo, CB_GETCURSEL, 0, 0);
                if (sel != CB_ERR) {
                    int portIndex = SendMessage(hWndCombo, CB_GETITEMDATA, sel, 0);
                    int colorIndex = LOWORD(wParam) - 9;
                    portColors[portIndex] = portColors[colorIndex];
                }
            } else if (LOWORD(wParam) == 4) {
                // Stop/Start listening button clicked
                if (isListening) {
                    // Disable all controls
                    EnableWindow(hWndButton, FALSE);
                    EnableWindow(hWndStopBtn, FALSE);
                    EnableWindow(hWndExportBtn, FALSE);
                    EnableWindow(hWndClearBtn, FALSE);
                    EnableWindow(hWndPauseBtn, FALSE);
                    for(int i = 0; i < 10; i++) EnableWindow(hWndColorBtns[i], FALSE);
                    EnableWindow(hWndCombo, FALSE);
                    EnableWindow(hWndBaudCombo, FALSE);
                    // Update status
                    SendMessage(hWndStatusBar, SB_SETTEXT, 0, (LPARAM)L"Releasing resources...");
                    // Create progress bar
                    RECT rcStatus;
                    GetWindowRect(hWndStatusBar, &rcStatus);
                    ScreenToClient(hwnd, (LPPOINT)&rcStatus.left);
                    ScreenToClient(hwnd, (LPPOINT)&rcStatus.right);
                    int progressX = rcStatus.left + 200;
                    int progressWidth = rcStatus.right - rcStatus.left - 200;
                    int progressHeight = rcStatus.bottom - rcStatus.top - 4;
                    HWND hProgress = CreateWindowEx(0, PROGRESS_CLASS, NULL, WS_CHILD | WS_VISIBLE, progressX, rcStatus.top + 2, progressWidth, progressHeight, hwnd, NULL, GetModuleHandle(NULL), NULL);
                    SendMessage(hProgress, PBM_SETRANGE, 0, MAKELPARAM(0, 100));
                    SendMessage(hProgress, PBM_SETPOS, 0, 0);
                    // Step 1: stop running
                    SendMessage(hProgress, PBM_SETPOS, 20, 0);
                    running = false;
                    isListening = false;
                    // Step 2: join threads
                    SendMessage(hProgress, PBM_SETPOS, 40, 0);
                    for (auto& t : readThreads) {
                        if (t.joinable()) t.join();
                    }
                    // Step 3: close handles
                    SendMessage(hProgress, PBM_SETPOS, 60, 0);
                    for (auto h : hSerialPorts) {
                        if (h != INVALID_HANDLE_VALUE) CloseHandle(h);
                    }
                    // Step 4: clear vectors
                    SendMessage(hProgress, PBM_SETPOS, 80, 0);
                    hSerialPorts.clear();
                    readThreads.clear();
                    portIndices.clear();
                    // Step 5: done
                    SendMessage(hProgress, PBM_SETPOS, 100, 0);
                    Sleep(200); // brief pause to show 100%
                    // Destroy progress bar
                    DestroyWindow(hProgress);
                    // Enable all controls
                    EnableWindow(hWndButton, TRUE);
                    EnableWindow(hWndStopBtn, TRUE);
                    EnableWindow(hWndExportBtn, TRUE);
                    EnableWindow(hWndClearBtn, TRUE);
                    EnableWindow(hWndPauseBtn, TRUE);
                    for(int i = 0; i < 10; i++) EnableWindow(hWndColorBtns[i], TRUE);
                    EnableWindow(hWndCombo, TRUE);
                    EnableWindow(hWndBaudCombo, TRUE);
                    // Set stop icon to play
                    SendMessage(hWndStopBtn, BM_SETIMAGE, IMAGE_ICON, (LPARAM)hPlayIcon);
                    RedrawWindow(hWndStopBtn, NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
                    // Update status bar
                    std::wstring statusText = L"Stopped - " + std::to_wstring(hSerialPorts.size()) + L" ports available";
                    SendMessage(hWndStatusBar, SB_SETTEXT, 0, (LPARAM)statusText.c_str());
                } else {
                    running = true;
                    readThreads.resize(hSerialPorts.size());
                    for (size_t i = 0; i < hSerialPorts.size(); ++i) {
                        if (hSerialPorts[i] != INVALID_HANDLE_VALUE && !readThreads[i].joinable()) {
                            readThreads[i] = std::thread(ReadFromPort, hSerialPorts[i], portIndices[i]);
                        }
                    }
                    isListening = true;
                    EnableWindow(hWndBaudCombo, FALSE);
                    SendMessage(hWndStopBtn, BM_SETIMAGE, IMAGE_ICON, (LPARAM)hStopIcon);
                    RedrawWindow(hWndStopBtn, NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
                    // Update status bar
                    std::wstring statusText = L"Monitoring " + std::to_wstring(hSerialPorts.size()) + L" ports";
                    SendMessage(hWndStatusBar, SB_SETTEXT, 0, (LPARAM)statusText.c_str());
                }
            } else if (LOWORD(wParam) == 5) {
                // Export button clicked
                GETTEXTEX gt = {0};
                gt.cb = 1024 * 100;
                gt.flags = GT_USECRLF;
                gt.codepage = CP_UTF8;
                std::vector<char> buffer(gt.cb);
                gt.lpDefaultChar = NULL;
                gt.lpUsedDefChar = NULL;
                LRESULT len = SendMessage(hWndEdit, EM_GETTEXTEX, (WPARAM)&gt, (LPARAM)buffer.data());
                if (len > 0) {
                    std::string text(buffer.data(), len);
                    std::ofstream file("output.html");
                    file << "<html><head><style>body { background-color: #404040; font-family: monospace; }</style></head><body>\n";
                    std::stringstream ss(text);
                    std::string line;
                    int currentPort = 0;
                    while (std::getline(ss, line, '\r')) {
                        if (!line.empty() && line != "\n") {
                            size_t pos = line.find(" COM");
                            if (pos != std::string::npos) {
                                size_t end = line.find(":", pos);
                                if (end != std::string::npos) {
                                    try {
                                        int port = std::stoi(line.substr(pos + 4, end - pos - 4));
                                        currentPort = port;
                                        COLORREF c = portColors[port];
                                        file << "<span style=\"color: rgb(" << (int)GetRValue(c) << "," << (int)GetGValue(c) << "," << (int)GetBValue(c) << ")\">" << line << "</span><br>\n";
                                    } catch (...) {
                                        COLORREF c = portColors[currentPort];
                                        file << "<span style=\"color: rgb(" << (int)GetRValue(c) << "," << (int)GetGValue(c) << "," << (int)GetBValue(c) << ")\">" << line << "</span><br>\n";
                                    }
                                } else {
                                    COLORREF c = portColors[currentPort];
                                    file << "<span style=\"color: rgb(" << (int)GetRValue(c) << "," << (int)GetGValue(c) << "," << (int)GetBValue(c) << ")\">" << line << "</span><br>\n";
                                }
                            } else {
                                COLORREF c = portColors[currentPort];
                                file << "<span style=\"color: rgb(" << (int)GetRValue(c) << "," << (int)GetGValue(c) << "," << (int)GetBValue(c) << ")\">" << line << "</span><br>\n";
                            }
                        }
                    }
                    file << "</body></html>";
                    file.close();
                    MessageBox(NULL, L"Exported to output.html", L"Export", MB_OK);
                }
            } else if (LOWORD(wParam) == 6) {
                // Pause/Resume scroll button clicked
                if (paused) {
                    paused = false;
                    SetWindowText(hWndPauseBtn, L"Pause Scroll");
                } else {
                    paused = true;
                    SetWindowText(hWndPauseBtn, L"Resume Scroll");
                }
            } else if (LOWORD(wParam) == 7) {
                // Clear button clicked
                SendMessage(hWndEdit, WM_SETTEXT, 0, (LPARAM)L"");
            } else if (LOWORD(wParam) == 3 && HIWORD(wParam) == CBN_SELCHANGE) {
                // Baud combo selection changed
                int sel = SendMessage(hWndBaudCombo, CB_GETCURSEL, 0, 0);
                if (sel != CB_ERR) {
                    currentBaud = (DWORD)SendMessage(hWndBaudCombo, CB_GETITEMDATA, sel, 0);
                    // Stop listening if active
                    if (isListening) {
                        running = false;
                        isListening = false;
                        for (auto& t : readThreads) {
                            if (t.joinable()) t.join();
                        }
                        SendMessage(hWndStopBtn, BM_SETIMAGE, IMAGE_ICON, (LPARAM)hPlayIcon);
                        RedrawWindow(hWndStopBtn, NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
                        EnableWindow(hWndBaudCombo, TRUE);
                    }
                    // Close existing handles
                    for (auto h : hSerialPorts) {
                        if (h != INVALID_HANDLE_VALUE) CloseHandle(h);
                    }
                    hSerialPorts.clear();
                    readThreads.clear();
                    portIndices.clear();
                    // Close dropdown and reset combo
                    SendMessage(hWndCombo, CB_SHOWDROPDOWN, FALSE, 0);
                    SendMessage(hWndCombo, CB_RESETCONTENT, 0, 0);
                    // Setup ports with new baud but don't start listening
                    SetupSerialPorts(false);
                }
            }
            break;
        case WM_CLOSE:
            DestroyWindow(hwnd);
            break;
        case WM_DESTROY:
            PostQuitMessage(0);
            break;
        default:
            return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

void CreateMainWindow(HINSTANCE hInstance) {
    hWndMain = CreateWindowEx(WS_EX_CLIENTEDGE, L"SerialMonitorClass", L"Multi Serial Monitor",
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 860, 640,
        NULL, NULL, hInstance, NULL);

    if (hWndMain == NULL) {
        MessageBox(NULL, L"Window Creation Failed!", L"Error!", MB_ICONEXCLAMATION | MB_OK);
    }
}

void SetupSerialPorts(bool startThreads) {
    for (int i = 1; i <= 255; ++i) {
        WCHAR portName[10];
        swprintf(portName, L"COM%d", i);
        WCHAR path[MAX_PATH];
        if (QueryDosDevice(portName, path, MAX_PATH)) {
            // port exists
            std::wstring fullName = L"\\\\.\\" + std::wstring(portName);
            HANDLE hPort = CreateFile(fullName.c_str(), GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, NULL);
            if (hPort != INVALID_HANDLE_VALUE) {
                // Configure serial port
                DCB dcb = {0};
                dcb.DCBlength = sizeof(DCB);
                GetCommState(hPort, &dcb);
                dcb.BaudRate = currentBaud;
                dcb.ByteSize = 8;
                dcb.StopBits = ONESTOPBIT;
                dcb.Parity = NOPARITY;
                SetCommState(hPort, &dcb);
                // Set timeouts
                COMMTIMEOUTS timeouts = {0};
                timeouts.ReadIntervalTimeout = 50;
                timeouts.ReadTotalTimeoutConstant = 50;
                timeouts.ReadTotalTimeoutMultiplier = 10;
                SetCommTimeouts(hPort, &timeouts);
                hSerialPorts.push_back(hPort);
                portIndices.push_back(i);
                if (startThreads) {
                    readThreads.emplace_back(ReadFromPort, hPort, i);
                    std::wstring msg = L"Opened " + std::wstring(portName) + L"\r\n";
                    AppendTextToEdit(msg, i);
                }
                // Add to combo
                int cbIndex = SendMessage(hWndCombo, CB_ADDSTRING, 0, (LPARAM)portName);
                SendMessage(hWndCombo, CB_SETITEMDATA, cbIndex, i);
            }
        }
    }
    // Select the first item if any and starting threads
    if (startThreads && SendMessage(hWndCombo, CB_GETCOUNT, 0, 0) > 0) {
        SendMessage(hWndCombo, CB_SETCURSEL, 0, 0);
    }
}

void ReadFromPort(HANDLE hPort, int portIndex) {
    char buffer[1024];
    DWORD bytesRead;
    OVERLAPPED ov = {0};
    ov.hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);

    while (running) {
        if (ReadFile(hPort, buffer, sizeof(buffer) - 1, &bytesRead, &ov)) {
            if (bytesRead > 0) {
                buffer[bytesRead] = '\0';
                int wlen = MultiByteToWideChar(CP_UTF8, 0, buffer, bytesRead, NULL, 0);
                std::vector<wchar_t> wbuf(wlen + 1);
                MultiByteToWideChar(CP_UTF8, 0, buffer, bytesRead, wbuf.data(), wlen);
                wbuf[wlen] = L'\0';
                std::time_t t = std::time(nullptr);
                std::tm* tm = std::localtime(&t);
                std::wstringstream ss;
                ss << std::put_time(tm, L"%H:%M:%S") << L" COM" << portIndex << L": " << wbuf.data() << L"\r\n";
                std::wstring text = ss.str();
                AppendTextToEdit(text, portIndex);
            }
        } else if (GetLastError() == ERROR_IO_PENDING) {
            WaitForSingleObject(ov.hEvent, INFINITE);
            GetOverlappedResult(hPort, &ov, &bytesRead, FALSE);
            if (bytesRead > 0) {
                buffer[bytesRead] = '\0';
                int wlen = MultiByteToWideChar(CP_UTF8, 0, buffer, bytesRead, NULL, 0);
                std::vector<wchar_t> wbuf(wlen + 1);
                MultiByteToWideChar(CP_UTF8, 0, buffer, bytesRead, wbuf.data(), wlen);
                wbuf[wlen] = L'\0';
                std::time_t t = std::time(nullptr);
                std::tm* tm = std::localtime(&t);
                std::wstringstream ss;
                ss << std::put_time(tm, L"%H:%M:%S") << L" COM" << portIndex << L": " << wbuf.data() << L"\r\n";
                std::wstring text = ss.str();
                AppendTextToEdit(text, portIndex);
            }
        }
        ResetEvent(ov.hEvent);
    }

    CloseHandle(ov.hEvent);
}

void AppendTextToEdit(const std::wstring& text, int portIndex) {
    if (paused) return;
    std::lock_guard<std::mutex> lock(textMutex);
    CHARFORMAT cf = {0};
    cf.cbSize = sizeof(cf);
    cf.dwMask = CFM_COLOR;
    cf.crTextColor = portColors[portIndex];
    SendMessage(hWndEdit, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
    SendMessage(hWndEdit, EM_SETSEL, -1, -1);
    SendMessage(hWndEdit, EM_REPLACESEL, FALSE, (LPARAM)text.c_str());
}

void RefreshPorts() {
    // Stop running
    running = false;
    // Join threads
    for (auto& t : readThreads) {
        if (t.joinable()) t.join();
    }
    // Close handles
    for (auto h : hSerialPorts) {
        if (h != INVALID_HANDLE_VALUE) CloseHandle(h);
    }
    // Clear vectors
    hSerialPorts.clear();
    readThreads.clear();
    portIndices.clear();
    // Clear combo
    SendMessage(hWndCombo, CB_RESETCONTENT, 0, 0);
    // Restart
    running = true;
    SetupSerialPorts(true);
}
