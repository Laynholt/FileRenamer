#include <windows.h>
#include <filesystem>
#include <fstream>
#include <functional>
#include <iostream>
#include <stdexcept>
#include <string>
#include <vector>

namespace fs = std::filesystem;
bool captureOnDesktop = false;
void Require(bool ok, const char* message) { if (!ok) throw std::runtime_error(message); }
void Wait(const std::function<bool()>& ready, const char* message) {
    for (int i = 0; i < 150; ++i) { if (ready()) return; Sleep(20); }
    Require(false, message);
}
std::wstring Text(HWND window) {
    const auto length = SendMessageW(window, WM_GETTEXTLENGTH, 0, 0);
    std::wstring text(static_cast<size_t>(length) + 1, L'\0');
    SendMessageW(window, WM_GETTEXT, text.size(), reinterpret_cast<LPARAM>(text.data()));
    text.resize(wcslen(text.c_str()));
    return text;
}
HWND FindControl(HWND parent, int id) {
    struct Search { int id; HWND found = nullptr; } search {id};
    EnumChildWindows(parent, [](HWND child, LPARAM data) -> BOOL {
        auto& s = *reinterpret_cast<Search*>(data);
        if (GetDlgCtrlID(child) == s.id) { s.found = child; return FALSE; }
        return TRUE;
    }, reinterpret_cast<LPARAM>(&search));
    return search.found;
}
int ClientWidth(HWND window) { RECT r = {}; GetClientRect(window, &r); return r.right; }
int ContentTop(HWND content) {
    RECT r = {};
    GetWindowRect(content, &r);
    MapWindowPoints(nullptr, GetParent(content), reinterpret_cast<POINT*>(&r), 2);
    return r.top;
}
struct WindowSearch { DWORD processId; const wchar_t* className; HWND result = nullptr; };
HWND FindProcessWindow(DWORD processId, const wchar_t* className) {
    WindowSearch search {processId, className};
    EnumWindows([](HWND window, LPARAM data) -> BOOL {
        auto& s = *reinterpret_cast<WindowSearch*>(data);
        DWORD pid = 0;
        GetWindowThreadProcessId(window, &pid);
        wchar_t name[128] = {};
        GetClassNameW(window, name, 128);
        if (pid == s.processId && wcscmp(name, s.className) == 0) { s.result = window; return FALSE; }
        return TRUE;
    }, reinterpret_cast<LPARAM>(&search));
    return search.result;
}
void Capture(HWND window, const fs::path& path) {
    RECT rect = {};
    GetWindowRect(window, &rect);
    const int width = rect.right - rect.left, height = rect.bottom - rect.top;
    HDC screen = GetDC(nullptr), memory = CreateCompatibleDC(screen);
    BITMAPINFO info = {};
    info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    info.bmiHeader.biWidth = width;
    info.bmiHeader.biHeight = -height;
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    void* pixels = nullptr;
    HBITMAP bitmap = CreateDIBSection(screen, &info, DIB_RGB_COLORS, &pixels, nullptr, 0);
    HGDIOBJ previous = SelectObject(memory, bitmap);
    if (captureOnDesktop) {
        RedrawWindow(window, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
        HDC source = GetWindowDC(window);
        BitBlt(memory, 0, 0, width, height, source, 0, 0, SRCCOPY);
        ReleaseDC(window, source);
    } else {
        PrintWindow(window, memory, 2);
    }
    BITMAPFILEHEADER header = {};
    header.bfType = 0x4d42;
    header.bfOffBits = sizeof(header) + sizeof(BITMAPINFOHEADER);
    header.bfSize = header.bfOffBits + width * height * 4;
    std::ofstream out(path, std::ios::binary);
    out.write(reinterpret_cast<const char*>(&header), sizeof(header));
    out.write(reinterpret_cast<const char*>(&info.bmiHeader), sizeof(BITMAPINFOHEADER));
    out.write(static_cast<const char*>(pixels), static_cast<std::streamsize>(width) * height * 4);
    SelectObject(memory, previous);
    DeleteObject(bitmap);
    DeleteDC(memory);
    ReleaseDC(nullptr, screen);
}

int wmain(int argc, wchar_t** argv) {
    PROCESS_INFORMATION process = {};
    const auto fixture = fs::temp_directory_path() / (L"FileRenamer-ui-" + std::to_wstring(GetCurrentProcessId()));
    HWND main = nullptr;
    bool fixtureCreated = false;
    const HDESK originalDesktop = GetThreadDesktop(GetCurrentThreadId());
    HDESK testDesktop = nullptr;
    int outcome = 0;
    try {
        captureOnDesktop = argc == 3 && std::wstring(argv[2]) == L"--capture-on-desktop";
        Require(argc == 2 || captureOnDesktop, "application path required");
        fixtureCreated = fs::create_directory(fixture);
        Require(fixtureCreated, "create unique test folder");
        std::ofstream(fixture / L"Show.S02E03.1080p.mkv") << "test";
        std::ofstream(fixture / L"Show S02E04 720p.mkv") << "test";
        std::ofstream(fixture / L"notes.txt") << "test";
        // Keep test windows and focus independent of the user's active desktop.
        std::wstring desktopName = L"FileRenamerTest-" + std::to_wstring(GetCurrentProcessId());
        if (!captureOnDesktop) {
            testDesktop = CreateDesktopW(desktopName.c_str(), nullptr, nullptr, 0, GENERIC_ALL, nullptr);
            Require(testDesktop && SetThreadDesktop(testDesktop), "create isolated test desktop");
        }
        STARTUPINFOW startup = {};
        startup.cb = sizeof(startup);
        startup.dwFlags = STARTF_USESHOWWINDOW;
        startup.wShowWindow = SW_SHOWNOACTIVATE;
        startup.lpDesktop = captureOnDesktop ? nullptr : desktopName.data();
        std::wstring command = L"\"" + std::wstring(argv[1]) + L"\"";
        Require(CreateProcessW(nullptr, command.data(), nullptr, nullptr, FALSE, 0, nullptr, nullptr, &startup, &process),
                "launch application");
        WaitForInputIdle(process.hProcess, 5000);
        Wait([&] { main = FindProcessWindow(process.dwProcessId, L"FileRenamerWinApiClass"); return main != nullptr; }, "main window exists");
        auto control = [&](int id) { return GetDlgItem(main, id); };
        SendMessageW(control(1001), WM_SETTEXT, 0, reinterpret_cast<LPARAM>(fixture.c_str()));
        SendMessageW(main, WM_COMMAND, MAKEWPARAM(1005, BN_CLICKED), reinterpret_cast<LPARAM>(control(1005)));
        HWND popup = nullptr;
        Wait([&] { popup = FindProcessWindow(process.dwProcessId, L"FileRenamerSuggestions"); return popup && IsWindowVisible(popup); }, "regex opens suggestions");
        HWND list = FindControl(popup, 2301), details = FindControl(popup, 2302);
        Require(list && details, "suggestion controls exist");
        Require((GetWindowLongPtrW(details, GWL_STYLE) & WS_VSCROLL) == 0,
                "suggestion details must not show a native light scrollbar");
        Require((GetWindowLongPtrW(list, GWL_STYLE) & WS_VSCROLL) == 0,
                "suggestion list must not show a native light scrollbar");
        HWND detailsScroll = FindControl(popup, 2304);
        HWND detailsContent = GetParent(details);
        HWND detailsViewport = GetParent(detailsContent);
        Require(detailsScroll && ClientWidth(detailsScroll) == ClientWidth(detailsViewport),
                "scrollbar is hidden while description fits");
        RECT scrollBounds = {};
        GetWindowRect(detailsScroll, &scrollBounds);
        MapWindowPoints(nullptr, popup, reinterpret_cast<POINT*>(&scrollBounds), 2);
        MoveWindow(detailsScroll, scrollBounds.left, scrollBounds.top,
                   scrollBounds.right - scrollBounds.left, 48, TRUE);
        Require(ClientWidth(detailsScroll) > ClientWidth(detailsViewport), "overflow reveals custom scrollbar");
        fs::create_directories(L"ui-artifacts");
        Capture(detailsScroll, L"ui-artifacts/scrollbar-overflow.bmp");
        SendMessageW(details, WM_MOUSEWHEEL, MAKEWPARAM(0, static_cast<WORD>(-WHEEL_DELTA)), 0);
        Require(ContentTop(detailsContent) < 0, "wheel moves description in custom viewport");
        const int barX = ClientWidth(detailsScroll) - 6;
        SendMessageW(detailsScroll, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(barX, 8));
        SendMessageW(detailsScroll, WM_MOUSEMOVE, MK_LBUTTON, MAKELPARAM(barX, 44));
        SendMessageW(detailsScroll, WM_LBUTTONUP, 0, MAKELPARAM(barX, 44));
        Require(IsWindowVisible(popup) && ContentTop(detailsContent) < 0, "scrollbar drag keeps completion popup open");
        GUITHREADINFO gui = {sizeof(GUITHREADINFO)};
        GetGUIThreadInfo(GetWindowThreadProcessId(main, nullptr), &gui);
        Require(gui.hwndFocus == control(1003), "scrollbar drag preserves edit focus");
        MoveWindow(detailsScroll, scrollBounds.left, scrollBounds.top,
                   scrollBounds.right - scrollBounds.left, scrollBounds.bottom - scrollBounds.top, TRUE);
        Require(ClientWidth(detailsScroll) == ClientWidth(detailsViewport) && ContentTop(detailsContent) == 0,
                "scrollbar disappears and offset resets when description fits again");
        Require(Text(details).find(L"S02") != std::wstring::npos, "suggestion includes actual filename");
        Require(Text(details).find(L"Не подходит: notes.txt") != std::wstring::npos, "suggestion explains excluded file");
        fs::create_directories(L"ui-artifacts");
        Capture(main, L"ui-artifacts/main.bmp");
        Capture(popup, L"ui-artifacts/suggestions.bmp");
        PostMessageW(control(1003), WM_KEYDOWN, VK_RETURN, 0);
        Wait([&] { return !IsWindowVisible(popup); }, "Enter closes popup after applying");
        Require(Text(control(1004)) == L"Название — сезон $2, серия $3$4", "Enter chooses suggestion");
        Require(Text(control(1009)).find(L"сезон 02, серия 03.mkv") != std::wstring::npos, "preview shows proposed names");
        Require(fs::exists(fixture / L"Show.S02E03.1080p.mkv"), "Enter in popup must not rename files");
        PostMessageW(control(1004), WM_KEYDOWN, VK_RETURN, static_cast<LPARAM>(1LL << 30));
        Sleep(80);
        Require(fs::exists(fixture / L"Show.S02E03.1080p.mkv"), "held Enter must not rename after applying suggestion");

        SendMessageW(control(1004), WM_SETTEXT, 0, reinterpret_cast<LPARAM>(L"Серия $9"));
        SendMessageW(control(1004), EM_SETSEL, 7, 7);
        Require(Text(control(1004)) == L"Серия $9", "existing group is present before completion");
        Require(SendMessageW(control(1004), EM_GETSEL, 0, 0) == MAKELONG(7, 7), "caret is inside the group token");
        SendMessageW(main, WM_COMMAND, MAKEWPARAM(1012, BN_CLICKED), reinterpret_cast<LPARAM>(control(1012)));
        Wait([&] { return IsWindowVisible(popup); }, "replacement suggestions open");
        Require(Text(details).find(L"группы 1") != std::wstring::npos, "typed dollar filters to groups");
        const HWND listScroll = FindControl(popup, 2303);
        const HWND listContent = GetParent(list);
        Require(ClientWidth(listScroll) > ClientWidth(GetParent(listContent)), "long suggestion list uses custom scrollbar");
        PostMessageW(control(1004), WM_KEYDOWN, VK_UP, 0);
        Wait([&] { return ContentTop(listContent) < 0; }, "keyboard selection scrolls last suggestion into view");
        PostMessageW(control(1004), WM_KEYDOWN, VK_DOWN, 0);
        Wait([&] { return ContentTop(listContent) == 0; }, "wrapping selection scrolls back to first suggestion");
        PostMessageW(control(1004), WM_KEYDOWN, VK_DOWN, 0);
        Wait([&] { return SendMessageW(list, LB_GETCURSEL, 0, 0) == 1; }, "Down selects next suggestion");
        PostMessageW(control(1004), WM_KEYDOWN, VK_RETURN, 0);
        Wait([&] { return Text(control(1004)) == L"Серия $2"; }, "completion inside a token replaces it without merging group numbers");

        for (const auto* value : {L"X$&", L"X$$"}) {
            SendMessageW(control(1004), WM_SETTEXT, 0, reinterpret_cast<LPARAM>(value));
            SendMessageW(control(1004), EM_SETSEL, 3, 3);
            SendMessageW(main, WM_COMMAND, MAKEWPARAM(1012, BN_CLICKED), reinterpret_cast<LPARAM>(control(1012)));
            Require(SendMessageW(list, LB_GETCOUNT, 0, 0) == 1, "special dollar token filters to itself");
            PostMessageW(control(1004), WM_KEYDOWN, VK_RETURN, 0);
            Wait([&] { return !IsWindowVisible(popup); }, "special dollar completion closes popup");
            Require(Text(control(1004)) == value, "special dollar token is completed without duplication");
        }

        SendMessageW(main, WM_COMMAND, MAKEWPARAM(1005, BN_CLICKED), reinterpret_cast<LPARAM>(control(1005)));
        SendMessageW(control(1003), WM_SETTEXT, 0, reinterpret_cast<LPARAM>(L""));
        SendMessageW(control(1004), WM_SETTEXT, 0, reinterpret_cast<LPARAM>(L"Название {n:03}"));
        SendMessageW(control(1004), EM_SETSEL, 13, 13);
        SendMessageW(main, WM_COMMAND, MAKEWPARAM(1012, BN_CLICKED), reinterpret_cast<LPARAM>(control(1012)));
        Require(SendMessageW(list, LB_GETCOUNT, 0, 0) == 2, "counter prefix filters width variants");
        SendMessageW(list, WM_LBUTTONDBLCLK, MK_LBUTTON, MAKELPARAM(20, 36));
        Wait([&] { return Text(control(1004)) == L"Название {n:03}"; }, "double click inserts padded counter");
        Require(Text(control(1009)).find(L"Название 001.txt") != std::wstring::npos, "counter preview preserves extension");
        SendMessageW(main, WM_COMMAND, MAKEWPARAM(1012, BN_CLICKED), reinterpret_cast<LPARAM>(control(1012)));
        PostMessageW(control(1004), WM_KEYDOWN, VK_ESCAPE, 0);
        Wait([&] { return !IsWindowVisible(popup); }, "Escape closes dropdown");
        Require(Text(control(1004)) == L"Название {n:03}", "Escape leaves template unchanged");
        std::cout << "All UI smoke checks passed\n";
    } catch (const std::exception& error) {
        std::cerr << "FAIL: " << error.what() << '\n';
        outcome = 1;
    }
    if (main) PostMessageW(main, WM_CLOSE, 0, 0);
    if (process.hProcess) {
        if (WaitForSingleObject(process.hProcess, 2000) == WAIT_TIMEOUT) TerminateProcess(process.hProcess, 1);
        CloseHandle(process.hProcess);
        CloseHandle(process.hThread);
    }
    std::error_code ec;
    if (fixtureCreated) fs::remove_all(fixture, ec);
    if (testDesktop) {
        SetThreadDesktop(originalDesktop);
        CloseDesktop(testDesktop);
    }
    return outcome;
}
