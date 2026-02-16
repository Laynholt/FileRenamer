
#include "Application.h"

#include "ExplorerPathProvider.h"
#include "RenamerService.h"
#include "Tooltil.h"
#include "UiRenderer.h"
#include "UpdateService.h"
#include "resource.h"

#include <windowsx.h>
#include <commctrl.h>
#include <objbase.h>
#include <shobjidl.h>

#include <algorithm>
#include <cctype>
#include <cwctype>
#include <filesystem>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "shell32.lib")

namespace fs = std::filesystem;


namespace {
const wchar_t* WINDOW_CLASS_NAME = L"FileRenamerWinApiClass";
const wchar_t* INFO_WINDOW_CLASS_NAME = L"FileRenamerInfoWindowClass";
const wchar_t* MESSAGE_WINDOW_CLASS_NAME = L"FileRenamerMessageWindowClass";
const wchar_t* APP_VERSION = L"1.0.2";

enum ControlId {
    ID_FOLDER_EDIT = 1001,
    ID_BROWSE_BUTTON = 1002,
    ID_PATTERN_EDIT = 1003,
    ID_REPLACEMENT_EDIT = 1004,
    ID_REGEX_CHECKBOX = 1005,
    ID_IGNORE_CASE_CHECKBOX = 1006,
    ID_RENAME_BUTTON = 1007,
    ID_CURRENT_PREVIEW = 1008,
    ID_RESULT_PREVIEW = 1009,
    ID_HELP_BUTTON = 1010
};

enum MenuId {
    ID_MENU_HELP_HOTKEYS = 2001,
    ID_MENU_HELP_ABOUT = 2002,
    ID_MENU_HELP_SEPARATOR = 2003,
    ID_MENU_CONTEXT_COPY = 2004
};

enum InfoControlId {
    ID_INFO_TEXT = 2101,
    ID_INFO_CLOSE = 2102,
    ID_INFO_CHECK_UPDATES = 2103
};

enum MessageControlId {
    ID_MESSAGE_TEXT = 2201,
    ID_MESSAGE_PRIMARY = 2202,
    ID_MESSAGE_SECONDARY = 2203
};

struct InfoWindowState {
    Application* owner;
    int kind;
    HWND textControl;
    HWND closeButton;
    HWND checkUpdatesButton;
    std::wstring text;
    HFONT font;
    HBRUSH editBrush;
};

struct MessageWindowState {
    Application* owner;
    HWND textControl;
    HWND primaryButton;
    HWND secondaryButton;
    HFONT font;
    HBRUSH editBrush;
    std::wstring text;
    std::wstring primaryButtonText;
    std::wstring secondaryButtonText;
    bool hasSecondaryButton;
    int result;
    int* resultOut;
};

constexpr UINT_PTR TEXT_CONTEXT_SUBCLASS_ID = 1;

std::wstring Trim(const std::wstring& text) {
    size_t begin = 0;
    while (begin < text.size() && std::iswspace(text[begin])) {
        ++begin;
    }

    size_t end = text.size();
    while (end > begin && std::iswspace(text[end - 1])) {
        --end;
    }

    return text.substr(begin, end - begin);
}

std::wstring PathCompareKey(const std::wstring& rawPath) {
    if (rawPath.empty()) {
        return L"";
    }

    std::error_code absoluteEc;
    const fs::path absolutePath = fs::absolute(fs::path(rawPath), absoluteEc);
    std::wstring key = absoluteEc
        ? fs::path(rawPath).lexically_normal().wstring()
        : absolutePath.lexically_normal().wstring();

    std::transform(key.begin(), key.end(), key.begin(), [](wchar_t ch) {
        return static_cast<wchar_t>(std::towlower(ch));
    });
    return key;
}

const wchar_t* GetHelpMenuItemText(UINT itemId) {
    switch (itemId) {
    case ID_MENU_HELP_HOTKEYS:
        return L"Горячие клавиши";
    case ID_MENU_HELP_ABOUT:
        return L"О программе";
    case ID_MENU_CONTEXT_COPY:
        return L"Копировать";
    default:
        return nullptr;
    }
}

const wchar_t* GetMenuItemText(UINT itemId, ULONG_PTR itemData) {
    if (itemData != 0) {
        return reinterpret_cast<const wchar_t*>(itemData);
    }
    return GetHelpMenuItemText(itemId);
}
} // namespace

Application::Application()
    : m_hInstance(nullptr)
    , m_hWnd(nullptr)
    , m_hFolderLabel(nullptr)
    , m_hFolderEdit(nullptr)
    , m_hBrowseButton(nullptr)
    , m_hPatternLabel(nullptr)
    , m_hPatternEdit(nullptr)
    , m_hReplacementLabel(nullptr)
    , m_hReplacementEdit(nullptr)
    , m_hRegexCheckbox(nullptr)
    , m_hIgnoreCaseCheckbox(nullptr)
    , m_hRenameButton(nullptr)
    , m_hHelpButton(nullptr)
    , m_hStatusLabel(nullptr)
    , m_hCurrentLabel(nullptr)
    , m_hResultLabel(nullptr)
    , m_hCurrentPreview(nullptr)
    , m_hResultPreview(nullptr)
    , m_hHotkeysWindow(nullptr)
    , m_hAboutWindow(nullptr)
    , m_hHelpMenu(nullptr)
    , m_hBackgroundBrush(nullptr)
    , m_hCardBrush(nullptr)
    , m_hFont(nullptr)
    , m_hMonoFont(nullptr)
    , m_gdiplusToken(0)
    , m_comInitialized(false)
    , m_useRegex(false)
    , m_ignoreCase(false)
    , m_infoWindowClassRegistered(false)
    , m_messageWindowClassRegistered(false)
    , m_hoveredControl(nullptr)
    , m_pressedControl(nullptr) {
}

Application::~Application() {
    Shutdown();
}

bool Application::Initialize(HINSTANCE hInstance) {
    m_hInstance = hInstance;

    Gdiplus::GdiplusStartupInput gdiplusStartupInput;
    if (Gdiplus::GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, nullptr) != Gdiplus::Ok) {
        MessageBox(nullptr, L"Не удалось инициализировать GDI+", L"Ошибка", MB_OK | MB_ICONERROR);
        return false;
    }

    const HRESULT comResult = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (SUCCEEDED(comResult)) {
        m_comInitialized = true;
    }

    INITCOMMONCONTROLSEX icex = {};
    icex.dwSize = sizeof(icex);
    icex.dwICC = ICC_WIN95_CLASSES;
    InitCommonControlsEx(&icex);

    m_hBackgroundBrush = CreateSolidBrush(RGB(26, 26, 26));
    m_hCardBrush = CreateSolidBrush(RGB(45, 45, 45));

    WNDCLASSEX wcex = {};
    wcex.cbSize = sizeof(wcex);
    wcex.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wcex.lpfnWndProc = WindowProc;
    wcex.hInstance = m_hInstance;
    wcex.hIcon = LoadIcon(m_hInstance, MAKEINTRESOURCE(IDI_MAIN_ICON));
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = m_hBackgroundBrush;
    wcex.lpszClassName = WINDOW_CLASS_NAME;
    wcex.hIconSm = LoadIcon(m_hInstance, MAKEINTRESOURCE(IDI_MAIN_ICON));

    if (!RegisterClassEx(&wcex) && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
        MessageBox(nullptr, L"Не удалось зарегистрировать класс окна", L"Ошибка", MB_OK | MB_ICONERROR);
        return false;
    }

    const int windowWidth = 980;
    const int windowHeight = 620;
    const int x = (GetSystemMetrics(SM_CXSCREEN) - windowWidth) / 2;
    const int y = (GetSystemMetrics(SM_CYSCREEN) - windowHeight) / 2;

    m_hWnd = CreateWindowEx(
        0,
        WINDOW_CLASS_NAME,
        Application::WINDOW_TITLE,
        WS_OVERLAPPEDWINDOW,
        x,
        y,
        windowWidth,
        windowHeight,
        nullptr,
        nullptr,
        m_hInstance,
        this
    );

    if (!m_hWnd) {
        MessageBox(nullptr, L"Не удалось создать главное окно", L"Ошибка", MB_OK | MB_ICONERROR);
        return false;
    }
    SetWindowTextW(m_hWnd, Application::WINDOW_TITLE);

    CreateControls();
    CreateHelpMenu();
    if (!RegisterInfoWindowClass()) {
        MessageBox(nullptr, L"Не удалось зарегистрировать класс информационного окна", L"Ошибка", MB_OK | MB_ICONERROR);
        return false;
    }
    if (!RegisterMessageWindowClass()) {
        MessageBox(nullptr, L"Не удалось зарегистрировать класс диалога сообщений", L"Ошибка", MB_OK | MB_ICONERROR);
        return false;
    }
    m_updateService = std::make_unique<UpdateService>();

    RECT clientRect = {};
    GetClientRect(m_hWnd, &clientRect);
    OnResize(clientRect.right - clientRect.left, clientRect.bottom - clientRect.top);

    SetTimer(m_hWnd, EXPLORER_SYNC_TIMER_ID, EXPLORER_SYNC_INTERVAL_MS, nullptr);

    PrefillFolderFromExplorer();
    UpdatePreview();
    return true;
}

int Application::Run() {
    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0)) {
        if (m_tooltil) {
            m_tooltil->RelayEvent(msg);
        }

        bool handled = false;
        if ((msg.message == WM_KEYDOWN || msg.message == WM_CHAR) &&
            (msg.hwnd == m_hWnd || (m_hWnd && IsChild(m_hWnd, msg.hwnd)))) {
            HWND focused = GetFocus();

            if (msg.message == WM_KEYDOWN) {
                switch (msg.wParam) {
                case VK_TAB:
                    SelectFolder();
                    handled = true;
                    break;
                case VK_RETURN:
                    RenameFiles();
                    handled = true;
                    break;
                case VK_ESCAPE:
                    if (focused == m_hFolderEdit || focused == m_hPatternEdit || focused == m_hReplacementEdit) {
                        SetFocus(m_hWnd);
                        handled = true;
                    }
                    break;
                case VK_DOWN:
                    if (focused == m_hPatternEdit) {
                        SetFocus(m_hReplacementEdit);
                        const int replacementLen = GetWindowTextLengthW(m_hReplacementEdit);
                        SendMessage(m_hReplacementEdit, EM_SETSEL, replacementLen, replacementLen);
                        handled = true;
                    }
                    break;
                case VK_UP:
                    if (focused == m_hReplacementEdit) {
                        SetFocus(m_hPatternEdit);
                        const int patternLen = GetWindowTextLengthW(m_hPatternEdit);
                        SendMessage(m_hPatternEdit, EM_SETSEL, patternLen, patternLen);
                        handled = true;
                    }
                    break;
                default:
                    break;
                }
            } else if (msg.message == WM_CHAR) {
                const wchar_t ch = static_cast<wchar_t>(msg.wParam);
                const bool inEditableInput =
                    focused == m_hFolderEdit ||
                    focused == m_hPatternEdit ||
                    focused == m_hReplacementEdit ||
                    msg.hwnd == m_hFolderEdit ||
                    msg.hwnd == m_hPatternEdit ||
                    msg.hwnd == m_hReplacementEdit;
                if (!inEditableInput && IsCharAlphaNumericW(ch) != FALSE) {
                    SetFocus(m_hPatternEdit);
                    const int patternLen = GetWindowTextLengthW(m_hPatternEdit);
                    SendMessage(m_hPatternEdit, EM_SETSEL, patternLen, patternLen);
                    SendMessage(m_hPatternEdit, WM_CHAR, msg.wParam, msg.lParam);
                    handled = true;
                }
            }
        }

        if (handled) {
            continue;
        }

        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return static_cast<int>(msg.wParam);
}

void Application::Shutdown() {
    if (m_hWnd && IsWindow(m_hWnd)) {
        KillTimer(m_hWnd, EXPLORER_SYNC_TIMER_ID);
    }

    m_tooltil.reset();

    if (m_hHelpMenu) {
        DestroyMenu(m_hHelpMenu);
        m_hHelpMenu = nullptr;
    }

    m_updateService.reset();

    if (m_hFont) {
        DeleteObject(m_hFont);
        m_hFont = nullptr;
    }

    if (m_hMonoFont) {
        DeleteObject(m_hMonoFont);
        m_hMonoFont = nullptr;
    }

    if (m_hCardBrush) {
        DeleteObject(m_hCardBrush);
        m_hCardBrush = nullptr;
    }

    if (m_hBackgroundBrush) {
        DeleteObject(m_hBackgroundBrush);
        m_hBackgroundBrush = nullptr;
    }

    if (m_gdiplusToken != 0) {
        Gdiplus::GdiplusShutdown(m_gdiplusToken);
        m_gdiplusToken = 0;
    }

    if (m_comInitialized) {
        CoUninitialize();
        m_comInitialized = false;
    }
}

LRESULT CALLBACK Application::WindowProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    Application* app = reinterpret_cast<Application*>(GetWindowLongPtr(hWnd, GWLP_USERDATA));

    if (message == WM_NCCREATE) {
        const CREATESTRUCT* createStruct = reinterpret_cast<CREATESTRUCT*>(lParam);
        app = static_cast<Application*>(createStruct->lpCreateParams);
        SetWindowLongPtr(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(app));
        if (app) {
            app->m_hWnd = hWnd;
        }
        return TRUE;
    }

    if (app) {
        return app->HandleMessage(message, wParam, lParam);
    }

    return DefWindowProc(hWnd, message, wParam, lParam);
}

LRESULT Application::HandleMessage(UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_COMMAND:
        if (lParam == 0) {
            OnMenuCommand(LOWORD(wParam));
        } else {
            OnCommand(LOWORD(wParam), HIWORD(wParam));
        }
        return 0;

    case WM_MEASUREITEM:
        {
            MEASUREITEMSTRUCT* mis = reinterpret_cast<MEASUREITEMSTRUCT*>(lParam);
            if (!mis || mis->CtlType != ODT_MENU) {
                break;
            }

            if (mis->itemID == ID_MENU_HELP_SEPARATOR) {
                mis->itemWidth = 60;
                mis->itemHeight = 10;
                return TRUE;
            }

            const wchar_t* text = GetMenuItemText(mis->itemID, mis->itemData);
            if (!text) {
                break;
            }

            HDC hdc = GetDC(m_hWnd);
            if (!hdc) {
                break;
            }

            HFONT hMenuFont = m_hFont ? m_hFont : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
            HFONT hOldFont = static_cast<HFONT>(SelectObject(hdc, hMenuFont));
            SIZE textSize = {};
            GetTextExtentPoint32W(hdc, text, static_cast<int>(wcslen(text)), &textSize);
            mis->itemWidth = textSize.cx + 38;
            mis->itemHeight = (textSize.cy + 10 > 26) ? (textSize.cy + 10) : 26;
            SelectObject(hdc, hOldFont);
            ReleaseDC(m_hWnd, hdc);
            return TRUE;
        }

    case WM_DRAWITEM:
        {
            DRAWITEMSTRUCT* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
            if (!dis) {
                break;
            }

            if (dis->CtlType == ODT_MENU) {
                if (dis->itemID == ID_MENU_HELP_SEPARATOR) {
                    HBRUSH hBgBrush = CreateSolidBrush(RGB(45, 45, 45));
                    FillRect(dis->hDC, &dis->rcItem, hBgBrush);
                    DeleteObject(hBgBrush);

                    HPEN hPen = CreatePen(PS_SOLID, 1, RGB(95, 95, 95));
                    HPEN hOldPen = static_cast<HPEN>(SelectObject(dis->hDC, hPen));
                    const int y = dis->rcItem.top + ((dis->rcItem.bottom - dis->rcItem.top) / 2);
                    const int padding = 12;
                    MoveToEx(dis->hDC, dis->rcItem.left + padding, y, nullptr);
                    LineTo(dis->hDC, dis->rcItem.right - padding, y);
                    SelectObject(dis->hDC, hOldPen);
                    DeleteObject(hPen);
                    return TRUE;
                }

                const wchar_t* text = GetMenuItemText(dis->itemID, dis->itemData);
                if (text) {
                    const bool isSelected = (dis->itemState & ODS_SELECTED) != 0;
                    const bool isDisabled = (dis->itemState & ODS_DISABLED) != 0;

                    HBRUSH hBgBrush = CreateSolidBrush(isSelected ? RGB(68, 68, 68) : RGB(45, 45, 45));
                    FillRect(dis->hDC, &dis->rcItem, hBgBrush);
                    DeleteObject(hBgBrush);

                    if (isSelected) {
                        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(85, 85, 85));
                        HPEN hOldPen = static_cast<HPEN>(SelectObject(dis->hDC, hPen));
                        HBRUSH hOldBrush = static_cast<HBRUSH>(SelectObject(dis->hDC, GetStockObject(NULL_BRUSH)));
                        Rectangle(dis->hDC, dis->rcItem.left, dis->rcItem.top, dis->rcItem.right, dis->rcItem.bottom);
                        SelectObject(dis->hDC, hOldBrush);
                        SelectObject(dis->hDC, hOldPen);
                        DeleteObject(hPen);
                    }

                    HFONT hMenuFont = m_hFont ? m_hFont : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
                    HFONT hOldFont = static_cast<HFONT>(SelectObject(dis->hDC, hMenuFont));
                    SetBkMode(dis->hDC, TRANSPARENT);
                    SetTextColor(dis->hDC, isDisabled ? RGB(150, 150, 150) : RGB(255, 255, 255));

                    RECT textRect = dis->rcItem;
                    textRect.left += 14;
                    DrawTextW(dis->hDC, text, -1, &textRect, DT_SINGLELINE | DT_VCENTER | DT_LEFT);
                    SelectObject(dis->hDC, hOldFont);
                    return TRUE;
                }
                break;
            }

            if (dis->CtlType != ODT_BUTTON) {
                break;
            }

            if (dis->CtlID == ID_REGEX_CHECKBOX || dis->CtlID == ID_IGNORE_CASE_CHECKBOX) {
                wchar_t text[256] = {};
                GetWindowTextW(dis->hwndItem, text, 256);
                const bool isPressed = m_pressedControl == dis->hwndItem || (dis->itemState & ODS_SELECTED) != 0;
                const bool hasFocus = (dis->itemState & ODS_FOCUS) != 0;
                const bool enabled = (dis->itemState & ODS_DISABLED) == 0;
                const bool isHot = m_hoveredControl == dis->hwndItem;
                const bool checked = (dis->CtlID == ID_REGEX_CHECKBOX) ? m_useRegex : m_ignoreCase;
                UiRenderer::DrawCustomCheckbox(dis->hDC, dis->hwndItem, text, checked, isHot, isPressed, enabled, hasFocus);
                return TRUE;
            }

            wchar_t text[256] = {};
            GetWindowTextW(dis->hwndItem, text, 256);
            const bool isPressed = m_pressedControl == dis->hwndItem || (dis->itemState & ODS_SELECTED) != 0;
            float hoverAlpha = 0.0f;
            auto hoverIt = m_buttonHoverAlpha.find(dis->hwndItem);
            if (hoverIt != m_buttonHoverAlpha.end()) {
                hoverAlpha = hoverIt->second;
            }
            UiRenderer::DrawCustomButton(dis->hDC, dis->hwndItem, text, isPressed, hoverAlpha);
            return TRUE;
        }

    case WM_MOUSEMOVE:
        {
            TRACKMOUSEEVENT tme = {};
            tme.cbSize = sizeof(tme);
            tme.dwFlags = TME_LEAVE;
            tme.hwndTrack = m_hWnd;
            TrackMouseEvent(&tme);

            POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
            UpdateHoverState(pt);
        }
        return 0;

    case WM_SETCURSOR:
        if (LOWORD(lParam) == HTCLIENT) {
            POINT cursorPos = {};
            GetCursorPos(&cursorPos);
            ScreenToClient(m_hWnd, &cursorPos);
            UpdateHoverState(cursorPos);
        }
        return DefWindowProc(m_hWnd, message, wParam, lParam);

    case WM_MOUSELEAVE:
        if (m_hoveredControl) {
            HWND oldHovered = m_hoveredControl;
            m_hoveredControl = nullptr;
            InvalidateRect(oldHovered, nullptr, TRUE);
        }
        for (auto& pair : m_buttonHoverAlpha) {
            pair.second = 0.0f;
        }
        return 0;

    case WM_LBUTTONDOWN:
        {
            POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
            if (IsPointInControl(m_hBrowseButton, pt)) {
                m_pressedControl = m_hBrowseButton;
            } else if (IsPointInControl(m_hRenameButton, pt)) {
                m_pressedControl = m_hRenameButton;
            } else if (IsPointInControl(m_hHelpButton, pt)) {
                m_pressedControl = m_hHelpButton;
            } else if (IsPointInControl(m_hRegexCheckbox, pt)) {
                m_pressedControl = m_hRegexCheckbox;
            } else if (IsPointInControl(m_hIgnoreCaseCheckbox, pt)) {
                m_pressedControl = m_hIgnoreCaseCheckbox;
            } else {
                m_pressedControl = nullptr;
            }

            if (m_pressedControl) {
                InvalidateRect(m_pressedControl, nullptr, TRUE);
            }
        }
        return 0;

    case WM_LBUTTONUP:
        if (m_pressedControl) {
            HWND oldPressed = m_pressedControl;
            m_pressedControl = nullptr;
            InvalidateRect(oldPressed, nullptr, TRUE);
        }
        return 0;

    case WM_CTLCOLORSTATIC:
        {
            HDC hdc = reinterpret_cast<HDC>(wParam);
            SetTextColor(hdc, RGB(228, 231, 236));
            SetBkColor(hdc, RGB(45, 45, 45));
            return reinterpret_cast<INT_PTR>(m_hCardBrush);
        }

    case WM_CTLCOLOREDIT:
        {
            HDC hdc = reinterpret_cast<HDC>(wParam);
            SetTextColor(hdc, RGB(243, 244, 246));
            SetBkColor(hdc, RGB(45, 45, 45));
            return reinterpret_cast<INT_PTR>(m_hCardBrush);
        }

    case WM_ERASEBKGND:
        return 1;

    case WM_PAINT:
        OnPaint();
        UiRenderer::DrawEditBorder(m_hWnd, m_hFolderEdit);
        UiRenderer::DrawEditBorder(m_hWnd, m_hPatternEdit);
        UiRenderer::DrawEditBorder(m_hWnd, m_hReplacementEdit);
        UiRenderer::DrawEditBorder(m_hWnd, m_hCurrentPreview);
        UiRenderer::DrawEditBorder(m_hWnd, m_hResultPreview);
        return 0;

    case WM_SIZE:
        OnResize(LOWORD(lParam), HIWORD(lParam));
        return 0;

    case WM_ACTIVATE:
        if (LOWORD(wParam) != WA_INACTIVE) {
            SyncFolderFromExplorer();
        }
        return 0;

    case WM_TIMER:
        if (wParam == EXPLORER_SYNC_TIMER_ID) {
            SyncFolderFromExplorer();
        }
        return 0;

    case WM_GETMINMAXINFO:
        {
            MINMAXINFO* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
            mmi->ptMinTrackSize.x = MIN_WINDOW_WIDTH;
            mmi->ptMinTrackSize.y = MIN_WINDOW_HEIGHT;
        }
        return 0;

    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProc(m_hWnd, message, wParam, lParam);
}
void Application::CreateControls() {
    m_hFont = CreateFont(
        -15,
        0,
        0,
        0,
        FW_NORMAL,
        FALSE,
        FALSE,
        FALSE,
        DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS,
        CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY,
        DEFAULT_PITCH | FF_SWISS,
        L"Segoe UI"
    );

    m_hMonoFont = CreateFont(
        -15,
        0,
        0,
        0,
        FW_NORMAL,
        FALSE,
        FALSE,
        FALSE,
        DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS,
        CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY,
        FIXED_PITCH | FF_MODERN,
        L"Consolas"
    );

    m_hFolderLabel = CreateWindowEx(
        0,
        L"STATIC",
        L"Папка:",
        WS_VISIBLE | WS_CHILD,
        0,
        0,
        0,
        0,
        m_hWnd,
        nullptr,
        m_hInstance,
        nullptr
    );

    m_hFolderEdit = CreateWindowEx(
        0,
        L"EDIT",
        L"",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | ES_AUTOHSCROLL,
        0,
        0,
        0,
        0,
        m_hWnd,
        reinterpret_cast<HMENU>(ID_FOLDER_EDIT),
        m_hInstance,
        nullptr
    );

    m_hBrowseButton = CreateWindowEx(
        0,
        L"BUTTON",
        L"Обзор...",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_OWNERDRAW,
        0,
        0,
        0,
        0,
        m_hWnd,
        reinterpret_cast<HMENU>(ID_BROWSE_BUTTON),
        m_hInstance,
        nullptr
    );

    m_hPatternLabel = CreateWindowEx(
        0,
        L"STATIC",
        L"Паттерн:",
        WS_VISIBLE | WS_CHILD,
        0,
        0,
        0,
        0,
        m_hWnd,
        nullptr,
        m_hInstance,
        nullptr
    );

    m_hPatternEdit = CreateWindowEx(
        0,
        L"EDIT",
        L"",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | ES_AUTOHSCROLL,
        0,
        0,
        0,
        0,
        m_hWnd,
        reinterpret_cast<HMENU>(ID_PATTERN_EDIT),
        m_hInstance,
        nullptr
    );

    m_hReplacementLabel = CreateWindowEx(
        0,
        L"STATIC",
        L"Шаблон замены:",
        WS_VISIBLE | WS_CHILD,
        0,
        0,
        0,
        0,
        m_hWnd,
        nullptr,
        m_hInstance,
        nullptr
    );

    m_hReplacementEdit = CreateWindowEx(
        0,
        L"EDIT",
        L"",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | ES_AUTOHSCROLL,
        0,
        0,
        0,
        0,
        m_hWnd,
        reinterpret_cast<HMENU>(ID_REPLACEMENT_EDIT),
        m_hInstance,
        nullptr
    );

    m_hRegexCheckbox = CreateWindowEx(
        0,
        L"BUTTON",
        L"Использовать regex",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX | BS_OWNERDRAW,
        0,
        0,
        0,
        0,
        m_hWnd,
        reinterpret_cast<HMENU>(ID_REGEX_CHECKBOX),
        m_hInstance,
        nullptr
    );

    m_hIgnoreCaseCheckbox = CreateWindowEx(
        0,
        L"BUTTON",
        L"Игнорировать регистр",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX | BS_OWNERDRAW,
        0,
        0,
        0,
        0,
        m_hWnd,
        reinterpret_cast<HMENU>(ID_IGNORE_CASE_CHECKBOX),
        m_hInstance,
        nullptr
    );

    m_hRenameButton = CreateWindowEx(
        0,
        L"BUTTON",
        L"Переименовать",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_OWNERDRAW,
        0,
        0,
        0,
        0,
        m_hWnd,
        reinterpret_cast<HMENU>(ID_RENAME_BUTTON),
        m_hInstance,
        nullptr
    );

    m_hHelpButton = CreateWindowEx(
        0,
        L"BUTTON",
        L"Справка",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_OWNERDRAW,
        0,
        0,
        0,
        0,
        m_hWnd,
        reinterpret_cast<HMENU>(ID_HELP_BUTTON),
        m_hInstance,
        nullptr
    );

    m_hStatusLabel = CreateWindowEx(
        0,
        L"STATIC",
        L"",
        WS_VISIBLE | WS_CHILD,
        0,
        0,
        0,
        0,
        m_hWnd,
        nullptr,
        m_hInstance,
        nullptr
    );

    m_hCurrentLabel = CreateWindowEx(
        0,
        L"STATIC",
        L"Элементы с совпадением",
        WS_VISIBLE | WS_CHILD,
        0,
        0,
        0,
        0,
        m_hWnd,
        nullptr,
        m_hInstance,
        nullptr
    );

    m_hResultLabel = CreateWindowEx(
        0,
        L"STATIC",
        L"После замены",
        WS_VISIBLE | WS_CHILD,
        0,
        0,
        0,
        0,
        m_hWnd,
        nullptr,
        m_hInstance,
        nullptr
    );

    m_hCurrentPreview = CreateWindowEx(
        0,
        L"EDIT",
        L"",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | ES_AUTOVSCROLL | ES_MULTILINE | ES_READONLY,
        0,
        0,
        0,
        0,
        m_hWnd,
        reinterpret_cast<HMENU>(ID_CURRENT_PREVIEW),
        m_hInstance,
        nullptr
    );

    m_hResultPreview = CreateWindowEx(
        0,
        L"EDIT",
        L"",
        WS_VISIBLE | WS_CHILD | WS_TABSTOP | ES_AUTOVSCROLL | ES_MULTILINE | ES_READONLY,
        0,
        0,
        0,
        0,
        m_hWnd,
        reinterpret_cast<HMENU>(ID_RESULT_PREVIEW),
        m_hInstance,
        nullptr
    );

    SetWindowSubclass(m_hFolderEdit, TextEditSubclassProc, TEXT_CONTEXT_SUBCLASS_ID, reinterpret_cast<DWORD_PTR>(this));
    SetWindowSubclass(m_hPatternEdit, TextEditSubclassProc, TEXT_CONTEXT_SUBCLASS_ID, reinterpret_cast<DWORD_PTR>(this));
    SetWindowSubclass(m_hReplacementEdit, TextEditSubclassProc, TEXT_CONTEXT_SUBCLASS_ID, reinterpret_cast<DWORD_PTR>(this));
    SetWindowSubclass(m_hCurrentPreview, TextEditSubclassProc, TEXT_CONTEXT_SUBCLASS_ID, reinterpret_cast<DWORD_PTR>(this));
    SetWindowSubclass(m_hResultPreview, TextEditSubclassProc, TEXT_CONTEXT_SUBCLASS_ID, reinterpret_cast<DWORD_PTR>(this));

    SendMessage(m_hRegexCheckbox, BM_SETCHECK, BST_UNCHECKED, 0);
    SendMessage(m_hIgnoreCaseCheckbox, BM_SETCHECK, BST_UNCHECKED, 0);

    EnumChildWindows(
        m_hWnd,
        [](HWND child, LPARAM fontParam) -> BOOL {
            SendMessage(child, WM_SETFONT, fontParam, MAKELPARAM(FALSE, 0));
            return TRUE;
        },
        reinterpret_cast<LPARAM>(m_hFont)
    );

    SendMessage(m_hCurrentPreview, WM_SETFONT, reinterpret_cast<WPARAM>(m_hMonoFont), MAKELPARAM(FALSE, 0));
    SendMessage(m_hResultPreview, WM_SETFONT, reinterpret_cast<WPARAM>(m_hMonoFont), MAKELPARAM(FALSE, 0));

    m_buttonHoverAlpha[m_hBrowseButton] = 0.0f;
    m_buttonHoverAlpha[m_hRenameButton] = 0.0f;
    m_buttonHoverAlpha[m_hHelpButton] = 0.0f;

    m_tooltil = std::make_unique<Tooltil>();
    if (m_tooltil->Initialize(m_hWnd)) {
        m_tooltil->SetStyle(m_hFont, RGB(45, 45, 45), RGB(235, 235, 235));

        const std::wstring patternTooltip = L"Текст или regex-шаблон, который нужно найти в имени файла или папки.";
        const std::wstring replacementTooltip =
            L"Текст замены. Оставьте пустым, чтобы удалить найденный паттерн.\r\n"
            L"С пустым паттерном: <text добавляет text в начало, >text добавляет text в конец имени.";

        m_tooltil->AddTool(m_hPatternLabel, patternTooltip);
        m_tooltil->AddTool(m_hPatternEdit, patternTooltip);
        m_tooltil->AddTool(m_hReplacementLabel, replacementTooltip);
        m_tooltil->AddTool(m_hReplacementEdit, replacementTooltip);
    } else {
        m_tooltil.reset();
    }
}

void Application::CreateHelpMenu() {
    if (m_hHelpMenu) {
        DestroyMenu(m_hHelpMenu);
        m_hHelpMenu = nullptr;
    }

    m_hHelpMenu = CreatePopupMenu();
    if (!m_hHelpMenu) {
        return;
    }

    AppendMenuW(m_hHelpMenu, MF_OWNERDRAW, ID_MENU_HELP_HOTKEYS, GetMenuItemText(ID_MENU_HELP_HOTKEYS, 0));
    AppendMenuW(m_hHelpMenu, MF_OWNERDRAW, ID_MENU_HELP_SEPARATOR, L"");
    AppendMenuW(m_hHelpMenu, MF_OWNERDRAW, ID_MENU_HELP_ABOUT, GetMenuItemText(ID_MENU_HELP_ABOUT, 0));

    MENUINFO popupMenuInfo = {};
    popupMenuInfo.cbSize = sizeof(MENUINFO);
    popupMenuInfo.fMask = MIM_BACKGROUND;
    popupMenuInfo.hbrBack = m_hCardBrush;
    SetMenuInfo(m_hHelpMenu, &popupMenuInfo);
}

void Application::OnResize(int width, int height) {
    if (width <= 0 || height <= 0) {
        return;
    }

    const int outerMargin = 12;
    const int cardPadding = 16;
    const int cardGap = 12;

    const int formCardTop = outerMargin;
    const int formCardHeight = 250;
    const int formCardBottom = formCardTop + formCardHeight;
    const int previewCardTop = formCardBottom + cardGap;
    const int previewCardBottom = height - outerMargin;

    const int contentLeft = outerMargin + cardPadding;
    const int contentRight = width - outerMargin - cardPadding;
    const int labelWidth = 150;
    const int controlLeft = contentLeft + labelWidth;
    const int rowSpacing = 38;
    const int rowTop = formCardTop + 22;
    const int editHeight = 26;

    const int browseWidth = 120;
    const int browseLeft = contentRight - browseWidth;
    const int folderEditWidth = std::max(140, browseLeft - 10 - controlLeft);

    MoveWindow(m_hFolderLabel, contentLeft, rowTop + 3, labelWidth - 8, 22, TRUE);
    MoveWindow(m_hFolderEdit, controlLeft, rowTop, folderEditWidth, editHeight, TRUE);
    MoveWindow(m_hBrowseButton, browseLeft, rowTop - 1, browseWidth, 30, TRUE);

    MoveWindow(m_hPatternLabel, contentLeft, rowTop + rowSpacing + 3, labelWidth - 8, 22, TRUE);
    MoveWindow(m_hPatternEdit, controlLeft, rowTop + rowSpacing, contentRight - controlLeft, editHeight, TRUE);

    MoveWindow(m_hReplacementLabel, contentLeft, rowTop + rowSpacing * 2 + 3, labelWidth - 8, 22, TRUE);
    MoveWindow(m_hReplacementEdit, controlLeft, rowTop + rowSpacing * 2, contentRight - controlLeft, editHeight, TRUE);

    const int actionRowY = rowTop + rowSpacing * 3;
    MoveWindow(m_hRegexCheckbox, controlLeft, actionRowY, 210, 26, TRUE);
    MoveWindow(m_hIgnoreCaseCheckbox, controlLeft + 216, actionRowY, 220, 26, TRUE);
    MoveWindow(m_hRenameButton, contentRight - 155, actionRowY - 1, 155, 30, TRUE);
    MoveWindow(m_hHelpButton, contentRight - 155, actionRowY + 34, 155, 28, TRUE);

    MoveWindow(m_hStatusLabel, contentLeft, actionRowY + 68, contentRight - contentLeft, 22, TRUE);

    const int previewInnerLeft = outerMargin + cardPadding;
    const int previewInnerRight = width - outerMargin - cardPadding;
    const int previewTitleReserve = 40;
    const int headerY = previewCardTop + previewTitleReserve;
    const int columnGap = 12;
    const int columnWidth = std::max(180, (previewInnerRight - previewInnerLeft - columnGap) / 2);
    const int rightColumnLeft = previewInnerRight - columnWidth;

    MoveWindow(m_hCurrentLabel, previewInnerLeft, headerY, columnWidth, 22, TRUE);
    MoveWindow(m_hResultLabel, rightColumnLeft, headerY, columnWidth, 22, TRUE);

    const int previewTop = headerY + 24;
    const int previewHeight = std::max(120, previewCardBottom - cardPadding - previewTop);
    MoveWindow(m_hCurrentPreview, previewInnerLeft, previewTop, columnWidth, previewHeight, TRUE);
    MoveWindow(m_hResultPreview, rightColumnLeft, previewTop, columnWidth, previewHeight, TRUE);

    InvalidateRect(m_hWnd, nullptr, TRUE);
}

void Application::OnPaint() {
    PAINTSTRUCT ps = {};
    HDC hdc = BeginPaint(m_hWnd, &ps);

    RECT clientRect = {};
    GetClientRect(m_hWnd, &clientRect);
    UiRenderer::DrawBackground(hdc, clientRect);

    const int outerMargin = 12;
    const int formCardHeight = 250;
    const int formCardTop = outerMargin;
    const int formCardBottom = formCardTop + formCardHeight;
    const int previewCardTop = formCardBottom + 12;

    RECT formCard = {
        outerMargin,
        formCardTop,
        clientRect.right - outerMargin,
        formCardBottom
    };

    RECT previewCard = {
        outerMargin,
        previewCardTop,
        clientRect.right - outerMargin,
        clientRect.bottom - outerMargin
    };

    UiRenderer::DrawCard(hdc, formCard);
    UiRenderer::DrawCard(hdc, previewCard, L"Предпросмотр");

    EndPaint(m_hWnd, &ps);
}

void Application::OnCommand(UINT controlId, UINT notifyCode) {
    if (notifyCode == EN_CHANGE) {
        if (controlId == ID_FOLDER_EDIT || controlId == ID_PATTERN_EDIT || controlId == ID_REPLACEMENT_EDIT) {
            UpdatePreview();
        }
        return;
    }

    if (notifyCode != BN_CLICKED) {
        return;
    }

    switch (controlId) {
    case ID_BROWSE_BUTTON:
        SelectFolder();
        break;

    case ID_REGEX_CHECKBOX:
        m_useRegex = !m_useRegex;
        SendMessage(m_hRegexCheckbox, BM_SETCHECK, m_useRegex ? BST_CHECKED : BST_UNCHECKED, 0);
        InvalidateRect(m_hRegexCheckbox, nullptr, TRUE);
        UpdatePreview();
        break;
    case ID_IGNORE_CASE_CHECKBOX:
        m_ignoreCase = !m_ignoreCase;
        SendMessage(m_hIgnoreCaseCheckbox, BM_SETCHECK, m_ignoreCase ? BST_CHECKED : BST_UNCHECKED, 0);
        InvalidateRect(m_hIgnoreCaseCheckbox, nullptr, TRUE);
        UpdatePreview();
        break;

    case ID_RENAME_BUTTON:
        RenameFiles();
        break;

    case ID_HELP_BUTTON:
        ShowHelpMenu();
        break;

    default:
        break;
    }
}

void Application::OnMenuCommand(UINT menuId) {
    switch (menuId) {
    case ID_MENU_HELP_HOTKEYS:
        ShowHotkeysWindow();
        break;
    case ID_MENU_HELP_ABOUT:
        ShowAboutWindow();
        break;
    default:
        break;
    }
}

bool Application::RegisterInfoWindowClass() {
    if (m_infoWindowClassRegistered) {
        return true;
    }

    WNDCLASSEX wcex = {};
    wcex.cbSize = sizeof(wcex);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = InfoWindowProc;
    wcex.hInstance = m_hInstance;
    wcex.hIcon = LoadIcon(m_hInstance, MAKEINTRESOURCE(IDI_MAIN_ICON));
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = nullptr;
    wcex.lpszClassName = INFO_WINDOW_CLASS_NAME;
    wcex.hIconSm = LoadIcon(m_hInstance, MAKEINTRESOURCE(IDI_MAIN_ICON));

    if (!RegisterClassEx(&wcex) && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
        return false;
    }

    m_infoWindowClassRegistered = true;
    return true;
}

bool Application::RegisterMessageWindowClass() {
    if (m_messageWindowClassRegistered) {
        return true;
    }

    WNDCLASSEX wcex = {};
    wcex.cbSize = sizeof(wcex);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = MessageWindowProc;
    wcex.hInstance = m_hInstance;
    wcex.hIcon = LoadIcon(m_hInstance, MAKEINTRESOURCE(IDI_MAIN_ICON));
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = nullptr;
    wcex.lpszClassName = MESSAGE_WINDOW_CLASS_NAME;
    wcex.hIconSm = LoadIcon(m_hInstance, MAKEINTRESOURCE(IDI_MAIN_ICON));

    if (!RegisterClassEx(&wcex) && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
        return false;
    }

    m_messageWindowClassRegistered = true;
    return true;
}

void Application::ShowHelpMenu() {
    if (!m_hHelpMenu || !m_hHelpButton || !IsWindow(m_hHelpButton)) {
        return;
    }

    RECT buttonRect = {};
    GetWindowRect(m_hHelpButton, &buttonRect);

    SetForegroundWindow(m_hWnd);
    TrackPopupMenu(
        m_hHelpMenu,
        TPM_LEFTALIGN | TPM_TOPALIGN | TPM_LEFTBUTTON,
        buttonRect.left,
        buttonRect.bottom + 2,
        0,
        m_hWnd,
        nullptr
    );
    PostMessage(m_hWnd, WM_NULL, 0, 0);
}

void Application::ShowTextContextMenu(HWND targetControl, LPARAM lParam) {
    if (!targetControl || !IsWindow(targetControl)) {
        return;
    }

    HMENU contextMenu = CreatePopupMenu();
    if (!contextMenu) {
        return;
    }

    AppendMenuW(contextMenu, MF_OWNERDRAW, ID_MENU_CONTEXT_COPY, GetMenuItemText(ID_MENU_CONTEXT_COPY, 0));

    MENUINFO popupMenuInfo = {};
    popupMenuInfo.cbSize = sizeof(MENUINFO);
    popupMenuInfo.fMask = MIM_BACKGROUND;
    popupMenuInfo.hbrBack = m_hCardBrush;
    SetMenuInfo(contextMenu, &popupMenuInfo);

    DWORD selectionStart = 0;
    DWORD selectionEnd = 0;
    SendMessageW(targetControl, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd));
    const bool hasSelection = selectionEnd > selectionStart;
    const bool hasText = GetWindowTextLengthW(targetControl) > 0;

    EnableMenuItem(contextMenu, ID_MENU_CONTEXT_COPY, MF_BYCOMMAND | (hasText ? MF_ENABLED : MF_GRAYED));

    POINT popupPoint = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    if (popupPoint.x == -1 && popupPoint.y == -1) {
        if (!GetCaretPos(&popupPoint)) {
            RECT controlRect = {};
            GetWindowRect(targetControl, &controlRect);
            popupPoint.x = controlRect.left + 10;
            popupPoint.y = controlRect.top + 10;
        } else {
            ClientToScreen(targetControl, &popupPoint);
        }
    }

    HWND popupHostWindow = GetAncestor(targetControl, GA_ROOT);
    if (!popupHostWindow || !IsWindow(popupHostWindow)) {
        popupHostWindow = m_hWnd;
    }
    SetForegroundWindow(popupHostWindow);

    HWND popupOwner = (m_hWnd && IsWindow(m_hWnd)) ? m_hWnd : popupHostWindow;
    const UINT selectedCommand = TrackPopupMenu(
        contextMenu,
        TPM_LEFTALIGN | TPM_TOPALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD | TPM_NONOTIFY,
        popupPoint.x,
        popupPoint.y,
        0,
        popupOwner,
        nullptr
    );

    if (selectedCommand == ID_MENU_CONTEXT_COPY && hasText) {
        const int textLength = GetWindowTextLengthW(targetControl);
        if (textLength > 0) {
            std::wstring controlText(static_cast<size_t>(textLength) + 1, L'\0');
            GetWindowTextW(targetControl, controlText.data(), textLength + 1);
            controlText.resize(static_cast<size_t>(textLength));

            std::wstring textToCopy = controlText;
            if (hasSelection) {
                size_t start = static_cast<size_t>(selectionStart);
                size_t end = static_cast<size_t>(selectionEnd);
                if (start > controlText.size()) {
                    start = controlText.size();
                }
                if (end > controlText.size()) {
                    end = controlText.size();
                }
                if (end < start) {
                    end = start;
                }
                textToCopy = controlText.substr(start, end - start);
            }

            if (OpenClipboard(popupHostWindow)) {
                EmptyClipboard();

                const size_t bytes = (textToCopy.size() + 1) * sizeof(wchar_t);
                HGLOBAL memory = GlobalAlloc(GMEM_MOVEABLE, bytes);
                if (memory) {
                    void* memoryData = GlobalLock(memory);
                    if (memoryData) {
                        wcscpy_s(static_cast<wchar_t*>(memoryData), textToCopy.size() + 1, textToCopy.c_str());
                        GlobalUnlock(memory);
                        if (!SetClipboardData(CF_UNICODETEXT, memory)) {
                            GlobalFree(memory);
                        }
                    } else {
                        GlobalFree(memory);
                    }
                }

                CloseClipboard();
            }
        }
    }

    DestroyMenu(contextMenu);
    PostMessage(popupOwner, WM_NULL, 0, 0);
}

void Application::ShowHotkeysWindow() {
    const std::wstring hotkeysText =
        L"Горячие клавиши:\r\n\r\n"
        L"Tab\t— открыть выбор папки\r\n"
        L"Enter\t— запустить переименование\r\n"
        L"Esc\t— снять фокус с поля ввода\r\n"
        L"Down\t— из поля Паттерн перейти в Шаблон замены\r\n"
        L"Up\t— из поля Шаблон замены перейти в Паттерн\r\n"
        L"Буква/цифра вне полей\t— фокус в Паттерн и ввод символа";

    CreateOrActivateInfoWindow(
        InfoWindowKind::Hotkeys,
        m_hHotkeysWindow,
        GetMenuItemText(ID_MENU_HELP_HOTKEYS, 0),
        hotkeysText
    );
}

void Application::ShowAboutWindow() {
    std::wstring aboutText;
    aboutText.reserve(256);
    aboutText += L"FileRenamer\r\n";
    aboutText += L"Версия: ";
    aboutText += APP_VERSION;
    aboutText += L"\r\n";
    aboutText += L"Автор: laynholt\r\n\r\n";
    aboutText += L"Приложение для массового переименования файлов и папок ";
    aboutText += L"по строке или regex с предпросмотром результата.";

    CreateOrActivateInfoWindow(
        InfoWindowKind::About,
        m_hAboutWindow,
        GetMenuItemText(ID_MENU_HELP_ABOUT, 0),
        aboutText
    );
}

void Application::CreateOrActivateInfoWindow(InfoWindowKind kind, HWND& targetHandle, const wchar_t* title, const std::wstring& bodyText) {
    const wchar_t* effectiveTitle = title;
    if (!effectiveTitle || effectiveTitle[0] == L'\0') {
        const UINT menuItemId = (kind == InfoWindowKind::Hotkeys) ? ID_MENU_HELP_HOTKEYS : ID_MENU_HELP_ABOUT;
        effectiveTitle = GetMenuItemText(menuItemId, 0);
        if (!effectiveTitle) {
            effectiveTitle = L"";
        }
    }

    if (targetHandle && IsWindow(targetHandle)) {
        SetWindowTextW(targetHandle, effectiveTitle);
        ShowWindow(targetHandle, SW_SHOWNORMAL);
        SetForegroundWindow(targetHandle);
        return;
    }

    const int windowWidth = (kind == InfoWindowKind::Hotkeys) ? 640 : 520;
    const int windowHeight = (kind == InfoWindowKind::Hotkeys) ? 360 : 300;

    RECT parentRect = {};
    GetWindowRect(m_hWnd, &parentRect);
    int x = parentRect.left + ((parentRect.right - parentRect.left) - windowWidth) / 2;
    int y = parentRect.top + ((parentRect.bottom - parentRect.top) - windowHeight) / 2;
    if (x < 0) {
        x = 40;
    }
    if (y < 0) {
        y = 40;
    }

    auto* state = new InfoWindowState{
        this,
        static_cast<int>(kind),
        nullptr,
        nullptr,
        nullptr,
        bodyText,
        m_hFont,
        nullptr
    };

    targetHandle = CreateWindowEx(
        WS_EX_DLGMODALFRAME,
        INFO_WINDOW_CLASS_NAME,
        effectiveTitle,
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
        x,
        y,
        windowWidth,
        windowHeight,
        m_hWnd,
        nullptr,
        m_hInstance,
        state
    );

    if (!targetHandle) {
        delete state;
        SetStatusText(L"Не удалось открыть окно справки");
        return;
    }

    SetWindowTextW(targetHandle, effectiveTitle);
    ShowWindow(targetHandle, SW_SHOW);
    UpdateWindow(targetHandle);
}

void Application::OnInfoWindowClosed(InfoWindowKind kind) {
    if (kind == InfoWindowKind::Hotkeys) {
        m_hHotkeysWindow = nullptr;
    } else {
        m_hAboutWindow = nullptr;
    }
}

int Application::ShowStyledMessageDialog(const wchar_t* title,
                                         const std::wstring& bodyText,
                                         const wchar_t* primaryButtonText,
                                         const wchar_t* secondaryButtonText) {
    if (!m_hWnd || !IsWindow(m_hWnd)) {
        return IDCANCEL;
    }

    if (!RegisterMessageWindowClass()) {
        const UINT fallbackType = (secondaryButtonText && secondaryButtonText[0] != L'\0') ? MB_YESNO : MB_OK;
        return MessageBoxW(m_hWnd, bodyText.c_str(), title ? title : L"", fallbackType | MB_ICONINFORMATION);
    }

    int result = IDCANCEL;
    auto* state = new MessageWindowState{
        this,
        nullptr,
        nullptr,
        nullptr,
        m_hFont,
        nullptr,
        bodyText,
        (primaryButtonText && primaryButtonText[0] != L'\0') ? primaryButtonText : L"Закрыть",
        (secondaryButtonText ? secondaryButtonText : L""),
        (secondaryButtonText && secondaryButtonText[0] != L'\0'),
        IDCANCEL,
        &result
    };

    const int windowWidth = 560;
    const int windowHeight = state->hasSecondaryButton ? 320 : 290;
    RECT parentRect = {};
    GetWindowRect(m_hWnd, &parentRect);
    int x = parentRect.left + ((parentRect.right - parentRect.left) - windowWidth) / 2;
    int y = parentRect.top + ((parentRect.bottom - parentRect.top) - windowHeight) / 2;
    if (x < 0) {
        x = 50;
    }
    if (y < 0) {
        y = 50;
    }

    HWND dialogWindow = CreateWindowEx(
        WS_EX_DLGMODALFRAME,
        MESSAGE_WINDOW_CLASS_NAME,
        (title && title[0] != L'\0') ? title : L"",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
        x,
        y,
        windowWidth,
        windowHeight,
        m_hWnd,
        nullptr,
        m_hInstance,
        state
    );

    if (!dialogWindow) {
        const UINT fallbackType = state->hasSecondaryButton ? MB_YESNO : MB_OK;
        delete state;
        return MessageBoxW(m_hWnd, bodyText.c_str(), title ? title : L"", fallbackType | MB_ICONINFORMATION);
    }

    EnableWindow(m_hWnd, FALSE);
    ShowWindow(dialogWindow, SW_SHOW);
    UpdateWindow(dialogWindow);
    SetForegroundWindow(dialogWindow);

    MSG msg = {};
    while (IsWindow(dialogWindow)) {
        const BOOL getMessageResult = GetMessage(&msg, nullptr, 0, 0);
        if (getMessageResult == -1) {
            break;
        }
        if (getMessageResult == 0) {
            PostQuitMessage(static_cast<int>(msg.wParam));
            break;
        }

        if (!IsDialogMessage(dialogWindow, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    EnableWindow(m_hWnd, TRUE);
    SetForegroundWindow(m_hWnd);
    SetFocus(m_hWnd);
    return result;
}

void Application::ShowStyledMessage(const std::wstring& title, const std::wstring& message) {
    ShowStyledMessageDialog(title.c_str(), message, L"ОК");
}

void Application::CheckForUpdates() {
    const wchar_t* checkUpdatesTitle = L"Проверка обновлений";
    const wchar_t* updateTitle = L"Обновление";

    if (!m_updateService) {
        m_updateService = std::make_unique<UpdateService>();
    }
    if (!m_updateService) {
        ShowStyledMessageDialog(checkUpdatesTitle, L"Сервис обновлений не инициализирован", L"Закрыть");
        return;
    }

    SetCursor(LoadCursor(nullptr, IDC_WAIT));
    SetStatusText(L"Проверка обновлений...");
    const UpdateCheckResult checkResult = m_updateService->CheckForUpdates(APP_VERSION);
    SetCursor(LoadCursor(nullptr, IDC_ARROW));

    if (!checkResult.success) {
        std::wstring message = L"Не удалось проверить обновления.\r\n\r\n";
        message += checkResult.errorMessage;
        ShowStyledMessageDialog(checkUpdatesTitle, message, L"Закрыть");
        SetStatusText(L"Ошибка проверки обновлений");
        return;
    }

    const std::wstring latestVersion = checkResult.latestVersion.empty() ? checkResult.latestTag : checkResult.latestVersion;
    if (!checkResult.updateAvailable) {
        std::wstring message = L"Установлена последняя версия приложения.";
        if (!latestVersion.empty()) {
            message += L"\r\n\r\nТекущая версия: ";
            message += APP_VERSION;
            message += L"\r\nПоследний релиз: ";
            message += latestVersion;
        }
        ShowStyledMessageDialog(checkUpdatesTitle, message, L"Закрыть");
        SetStatusText(L"Установлена последняя версия");
        return;
    }

    std::wstring message = L"Доступна новая версия: ";
    message += latestVersion.empty() ? checkResult.latestTag : latestVersion;
    message += L"\r\nТекущая версия: ";
    message += APP_VERSION;
    message += L"\r\n\r\nСкачать и установить обновление сейчас?";

    const int userDecision = ShowStyledMessageDialog(
        checkUpdatesTitle,
        message,
        L"Скачать",
        L"Отмена"
    );

    if (userDecision != IDYES) {
        SetStatusText(L"Обновление отменено");
        return;
    }

    wchar_t tempPath[MAX_PATH] = {};
    const DWORD tempPathLength = GetTempPathW(MAX_PATH, tempPath);
    if (tempPathLength == 0 || tempPathLength >= MAX_PATH) {
        ShowStyledMessageDialog(updateTitle, L"Не удалось определить временную директорию", L"Закрыть");
        SetStatusText(L"Ошибка загрузки обновления");
        return;
    }

    std::wstring downloadedExePath = tempPath;
    if (!downloadedExePath.empty() && downloadedExePath.back() != L'\\') {
        downloadedExePath.push_back(L'\\');
    }
    downloadedExePath += L"FileRenamer_update.exe";

    std::wstring downloadError;
    SetCursor(LoadCursor(nullptr, IDC_WAIT));
    SetStatusText(L"Загрузка обновления...");
    const bool downloaded = m_updateService->DownloadReleaseExecutable(checkResult.latestTag, downloadedExePath, downloadError);
    SetCursor(LoadCursor(nullptr, IDC_ARROW));

    if (!downloaded) {
        std::wstring errorMessage = L"Не удалось скачать обновление.\r\n\r\n";
        errorMessage += downloadError;
        ShowStyledMessageDialog(updateTitle, errorMessage, L"Закрыть");
        SetStatusText(L"Ошибка загрузки обновления");
        return;
    }

    wchar_t currentExePath[MAX_PATH] = {};
    const DWORD currentExePathLength = GetModuleFileNameW(nullptr, currentExePath, MAX_PATH);
    if (currentExePathLength == 0 || currentExePathLength >= MAX_PATH) {
        DeleteFileW(downloadedExePath.c_str());
        ShowStyledMessageDialog(updateTitle, L"Не удалось определить путь текущего приложения", L"Закрыть");
        SetStatusText(L"Ошибка обновления");
        return;
    }

    std::wstring launchError;
    if (!m_updateService->LaunchUpdaterProcess(GetCurrentProcessId(), downloadedExePath, currentExePath, launchError)) {
        DeleteFileW(downloadedExePath.c_str());
        std::wstring errorMessage = L"Не удалось запустить установку обновления.\r\n\r\n";
        errorMessage += launchError;
        ShowStyledMessageDialog(updateTitle, errorMessage, L"Закрыть");
        SetStatusText(L"Ошибка обновления");
        return;
    }

    SetStatusText(L"Обновление готово. Перезапуск...");
    DestroyWindow(m_hWnd);
}

LRESULT CALLBACK Application::InfoWindowProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    auto* state = reinterpret_cast<InfoWindowState*>(GetWindowLongPtr(hWnd, GWLP_USERDATA));

    switch (message) {
    case WM_NCCREATE:
        {
            CREATESTRUCT* create = reinterpret_cast<CREATESTRUCT*>(lParam);
            auto* initialState = reinterpret_cast<InfoWindowState*>(create->lpCreateParams);
            SetWindowLongPtr(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(initialState));
            return TRUE;
        }

    case WM_CREATE:
        if (state) {
            state->editBrush = CreateSolidBrush(RGB(45, 45, 45));
            const InfoWindowKind infoKind = static_cast<InfoWindowKind>(state->kind);

            state->textControl = CreateWindowEx(
                0,
                L"EDIT",
                state->text.c_str(),
                WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_WANTRETURN | ES_READONLY,
                12, 12, 100, 100,
                hWnd,
                reinterpret_cast<HMENU>(ID_INFO_TEXT),
                GetModuleHandle(nullptr),
                nullptr
            );
            SetWindowSubclass(
                state->textControl,
                TextEditSubclassProc,
                TEXT_CONTEXT_SUBCLASS_ID,
                reinterpret_cast<DWORD_PTR>(state->owner)
            );

            state->closeButton = CreateWindowEx(
                0,
                L"BUTTON",
                L"Закрыть",
                WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON | BS_OWNERDRAW | WS_TABSTOP,
                12, 12, 120, 30,
                hWnd,
                reinterpret_cast<HMENU>(ID_INFO_CLOSE),
                GetModuleHandle(nullptr),
                nullptr
            );

            if (infoKind == InfoWindowKind::About) {
                state->checkUpdatesButton = CreateWindowEx(
                    0,
                    L"BUTTON",
                    L"Проверить обновления",
                    WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_TABSTOP,
                    12, 12, 210, 30,
                    hWnd,
                    reinterpret_cast<HMENU>(ID_INFO_CHECK_UPDATES),
                    GetModuleHandle(nullptr),
                    nullptr
                );
            }

            if (state->font) {
                SendMessage(state->textControl, WM_SETFONT, reinterpret_cast<WPARAM>(state->font), MAKELPARAM(FALSE, 0));
                SendMessage(state->closeButton, WM_SETFONT, reinterpret_cast<WPARAM>(state->font), MAKELPARAM(FALSE, 0));
                if (state->checkUpdatesButton) {
                    SendMessage(state->checkUpdatesButton, WM_SETFONT, reinterpret_cast<WPARAM>(state->font), MAKELPARAM(FALSE, 0));
                }
            }
        }
        return 0;

    case WM_SIZE:
        if (state) {
            const int width = LOWORD(lParam);
            const int height = HIWORD(lParam);
            const int margin = 20;
            const int buttonWidth = 120;
            const int buttonHeight = 30;
            const int checkUpdatesWidth = 210;

            MoveWindow(state->textControl, margin, margin, width - (margin * 2), height - (margin * 3) - buttonHeight, TRUE);
            MoveWindow(state->closeButton, width - margin - buttonWidth, height - margin - buttonHeight, buttonWidth, buttonHeight, TRUE);
            if (state->checkUpdatesButton) {
                MoveWindow(state->checkUpdatesButton, margin, height - margin - buttonHeight, checkUpdatesWidth, buttonHeight, TRUE);
            }
        }
        return 0;

    case WM_ERASEBKGND:
        return 1;

    case WM_PAINT:
        {
            PAINTSTRUCT ps = {};
            HDC hdc = BeginPaint(hWnd, &ps);

            RECT clientRect = {};
            GetClientRect(hWnd, &clientRect);
            UiRenderer::DrawBackground(hdc, clientRect);

            RECT cardRect = {8, 8, clientRect.right - 8, clientRect.bottom - 8};
            UiRenderer::DrawCard(hdc, cardRect);

            EndPaint(hWnd, &ps);

            if (state && state->textControl) {
                UiRenderer::DrawEditBorder(hWnd, state->textControl);
            }
        }
        return 0;

    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLOREDIT:
        if (state && state->editBrush) {
            HDC hdcControl = reinterpret_cast<HDC>(wParam);
            SetTextColor(hdcControl, RGB(255, 255, 255));
            SetBkColor(hdcControl, RGB(45, 45, 45));
            return reinterpret_cast<INT_PTR>(state->editBrush);
        }
        break;

    case WM_DRAWITEM:
        {
            DRAWITEMSTRUCT* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
            if (dis && dis->CtlType == ODT_BUTTON &&
                (dis->CtlID == ID_INFO_CLOSE || dis->CtlID == ID_INFO_CHECK_UPDATES)) {
                wchar_t buttonText[128] = {};
                GetWindowTextW(dis->hwndItem, buttonText, 128);
                const bool isPressed = (dis->itemState & ODS_SELECTED) != 0;
                const float hoverAlpha = (dis->itemState & ODS_HOTLIGHT) ? 1.0f : 0.0f;
                UiRenderer::DrawCustomButton(dis->hDC, dis->hwndItem, buttonText, isPressed, hoverAlpha);
                return TRUE;
            }
        }
        break;

    case WM_COMMAND:
        if (LOWORD(wParam) == ID_INFO_CHECK_UPDATES && state && state->owner) {
            state->owner->CheckForUpdates();
            return 0;
        }
        if (LOWORD(wParam) == ID_INFO_CLOSE || LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) {
            DestroyWindow(hWnd);
            return 0;
        }
        break;

    case WM_CLOSE:
        DestroyWindow(hWnd);
        return 0;

    case WM_NCDESTROY:
        if (state) {
            if (state->editBrush) {
                DeleteObject(state->editBrush);
                state->editBrush = nullptr;
            }
            if (state->owner) {
                state->owner->OnInfoWindowClosed(static_cast<InfoWindowKind>(state->kind));
            }
            delete state;
            SetWindowLongPtr(hWnd, GWLP_USERDATA, 0);
        }
        return DefWindowProc(hWnd, message, wParam, lParam);
    }

    return DefWindowProc(hWnd, message, wParam, lParam);
}

LRESULT CALLBACK Application::MessageWindowProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    auto* state = reinterpret_cast<MessageWindowState*>(GetWindowLongPtr(hWnd, GWLP_USERDATA));

    switch (message) {
    case WM_NCCREATE:
        {
            CREATESTRUCT* create = reinterpret_cast<CREATESTRUCT*>(lParam);
            auto* initialState = reinterpret_cast<MessageWindowState*>(create->lpCreateParams);
            SetWindowLongPtr(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(initialState));
            return TRUE;
        }

    case WM_CREATE:
        if (state) {
            state->editBrush = CreateSolidBrush(RGB(45, 45, 45));
            state->textControl = CreateWindowEx(
                0,
                L"EDIT",
                state->text.c_str(),
                WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_WANTRETURN | ES_READONLY,
                12, 12, 100, 100,
                hWnd,
                reinterpret_cast<HMENU>(ID_MESSAGE_TEXT),
                GetModuleHandle(nullptr),
                nullptr
            );
            SetWindowSubclass(
                state->textControl,
                TextEditSubclassProc,
                TEXT_CONTEXT_SUBCLASS_ID,
                reinterpret_cast<DWORD_PTR>(state->owner)
            );

            state->primaryButton = CreateWindowEx(
                0,
                L"BUTTON",
                state->primaryButtonText.c_str(),
                WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON | BS_OWNERDRAW | WS_TABSTOP,
                12, 12, 140, 32,
                hWnd,
                reinterpret_cast<HMENU>(ID_MESSAGE_PRIMARY),
                GetModuleHandle(nullptr),
                nullptr
            );

            if (state->hasSecondaryButton) {
                state->secondaryButton = CreateWindowEx(
                    0,
                    L"BUTTON",
                    state->secondaryButtonText.c_str(),
                    WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_TABSTOP,
                    12, 12, 140, 32,
                    hWnd,
                    reinterpret_cast<HMENU>(ID_MESSAGE_SECONDARY),
                    GetModuleHandle(nullptr),
                    nullptr
                );
            }

            if (state->font) {
                SendMessage(state->textControl, WM_SETFONT, reinterpret_cast<WPARAM>(state->font), MAKELPARAM(FALSE, 0));
                SendMessage(state->primaryButton, WM_SETFONT, reinterpret_cast<WPARAM>(state->font), MAKELPARAM(FALSE, 0));
                if (state->secondaryButton) {
                    SendMessage(state->secondaryButton, WM_SETFONT, reinterpret_cast<WPARAM>(state->font), MAKELPARAM(FALSE, 0));
                }
            }
        }
        return 0;

    case WM_SIZE:
        if (state) {
            const int width = LOWORD(lParam);
            const int height = HIWORD(lParam);
            const int margin = 20;
            const int buttonWidth = 140;
            const int buttonHeight = 32;
            const int buttonGap = 10;

            MoveWindow(
                state->textControl,
                margin,
                margin,
                width - (margin * 2),
                height - (margin * 3) - buttonHeight,
                TRUE
            );

            const int buttonY = height - margin - buttonHeight;
            if (state->hasSecondaryButton && state->secondaryButton) {
                const int primaryX = width - margin - buttonWidth;
                const int secondaryX = primaryX - buttonGap - buttonWidth;
                MoveWindow(state->secondaryButton, secondaryX, buttonY, buttonWidth, buttonHeight, TRUE);
                MoveWindow(state->primaryButton, primaryX, buttonY, buttonWidth, buttonHeight, TRUE);
            } else {
                MoveWindow(state->primaryButton, width - margin - buttonWidth, buttonY, buttonWidth, buttonHeight, TRUE);
            }
        }
        return 0;

    case WM_ERASEBKGND:
        return 1;

    case WM_PAINT:
        {
            PAINTSTRUCT ps = {};
            HDC hdc = BeginPaint(hWnd, &ps);

            RECT clientRect = {};
            GetClientRect(hWnd, &clientRect);
            UiRenderer::DrawBackground(hdc, clientRect);

            RECT cardRect = {8, 8, clientRect.right - 8, clientRect.bottom - 8};
            UiRenderer::DrawCard(hdc, cardRect);

            EndPaint(hWnd, &ps);

            if (state && state->textControl) {
                UiRenderer::DrawEditBorder(hWnd, state->textControl);
            }
        }
        return 0;

    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLOREDIT:
        if (state && state->editBrush) {
            HDC hdcControl = reinterpret_cast<HDC>(wParam);
            SetTextColor(hdcControl, RGB(255, 255, 255));
            SetBkColor(hdcControl, RGB(45, 45, 45));
            return reinterpret_cast<INT_PTR>(state->editBrush);
        }
        break;

    case WM_DRAWITEM:
        {
            DRAWITEMSTRUCT* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
            if (dis && dis->CtlType == ODT_BUTTON &&
                (dis->CtlID == ID_MESSAGE_PRIMARY || dis->CtlID == ID_MESSAGE_SECONDARY)) {
                wchar_t buttonText[128] = {};
                GetWindowTextW(dis->hwndItem, buttonText, 128);
                const bool isPressed = (dis->itemState & ODS_SELECTED) != 0;
                const float hoverAlpha = (dis->itemState & ODS_HOTLIGHT) ? 1.0f : 0.0f;
                UiRenderer::DrawCustomButton(dis->hDC, dis->hwndItem, buttonText, isPressed, hoverAlpha);
                return TRUE;
            }
        }
        break;

    case WM_COMMAND:
        if (state) {
            const UINT commandId = LOWORD(wParam);
            if (commandId == ID_MESSAGE_PRIMARY || commandId == IDOK) {
                state->result = state->hasSecondaryButton ? IDYES : IDOK;
                DestroyWindow(hWnd);
                return 0;
            }
            if (commandId == ID_MESSAGE_SECONDARY || commandId == IDCANCEL) {
                state->result = state->hasSecondaryButton ? IDNO : IDCANCEL;
                DestroyWindow(hWnd);
                return 0;
            }
        }
        break;

    case WM_CLOSE:
        if (state) {
            state->result = state->hasSecondaryButton ? IDNO : IDCANCEL;
        }
        DestroyWindow(hWnd);
        return 0;

    case WM_NCDESTROY:
        if (state) {
            if (state->editBrush) {
                DeleteObject(state->editBrush);
                state->editBrush = nullptr;
            }
            if (state->resultOut) {
                *state->resultOut = state->result;
            }
            delete state;
            SetWindowLongPtr(hWnd, GWLP_USERDATA, 0);
        }
        return DefWindowProc(hWnd, message, wParam, lParam);
    }

    return DefWindowProc(hWnd, message, wParam, lParam);
}

LRESULT CALLBACK Application::TextEditSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
    UNREFERENCED_PARAMETER(uIdSubclass);
    UNREFERENCED_PARAMETER(wParam);
    auto* app = reinterpret_cast<Application*>(dwRefData);

    if (uMsg == WM_CONTEXTMENU && app) {
        app->ShowTextContextMenu(hWnd, lParam);
        return 0;
    }

    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

void Application::UpdatePreview() {
    const std::wstring folderText = Trim(GetEditText(m_hFolderEdit));
    const std::wstring pattern = GetEditText(m_hPatternEdit);
    const std::wstring replacement = GetEditText(m_hReplacementEdit);
    RenamerCore::CollectResult result = RenamerCore::CollectOperations(
        folderText,
        pattern,
        replacement,
        m_useRegex,
        m_ignoreCase,
        PREVIEW_LIMIT
    );

    if (result.operations.empty()) {
        SetStatusText(result.status);
        SetEditText(m_hCurrentPreview, L"");
        SetEditText(m_hResultPreview, L"");
        return;
    }

    const size_t visibleCount = result.operations.size();

    std::vector<std::wstring> currentNames;
    std::vector<std::wstring> newNames;
    currentNames.reserve(visibleCount + 1);
    newNames.reserve(visibleCount + 1);

    for (size_t index = 0; index < visibleCount; ++index) {
        const auto& operation = result.operations[index];
        const std::wstring directorySuffix = operation.isDirectory ? L"\\" : L"";
        currentNames.push_back(operation.oldName + directorySuffix);
        newNames.push_back(operation.newName + directorySuffix);
    }

    std::wstring status = result.status;
    const size_t hiddenCount = result.totalCount > visibleCount ? (result.totalCount - visibleCount) : 0;
    if (hiddenCount > 0) {
        const std::wstring suffix = L"... и еще " + std::to_wstring(hiddenCount) + L" элементов";
        currentNames.push_back(suffix);
        newNames.push_back(suffix);
        status += L". Показано: " + std::to_wstring(visibleCount)
            + L" (лимит " + std::to_wstring(PREVIEW_LIMIT) + L").";
    }

    SetStatusText(status);
    SetEditText(m_hCurrentPreview, JoinLines(currentNames));
    SetEditText(m_hResultPreview, JoinLines(newNames));
}

void Application::RenameFiles() {
    const std::wstring folderText = Trim(GetEditText(m_hFolderEdit));
    const std::wstring pattern = GetEditText(m_hPatternEdit);
    const std::wstring replacement = GetEditText(m_hReplacementEdit);
    RenamerCore::CollectResult collectResult = RenamerCore::CollectOperations(
        folderText,
        pattern,
        replacement,
        m_useRegex,
        m_ignoreCase
    );

    if (collectResult.operations.empty()) {
        ShowStyledMessage(L"Внимание", collectResult.status);
        return;
    }

    const RenamerCore::ExecuteResult executeResult = RenamerCore::ExecuteRename(collectResult.operations);
    if (executeResult.status == RenamerCore::ExecuteStatus::NoChanges) {
        ShowStyledMessage(L"Готово", executeResult.message);
        return;
    }

    if (executeResult.status == RenamerCore::ExecuteStatus::Error) {
        ShowStyledMessage(L"Ошибка", executeResult.message);
        UpdatePreview();
        return;
    }

    const std::wstring successMessage = L"Переименовано элементов: " + std::to_wstring(executeResult.renamedCount);
    ShowStyledMessage(L"Готово", successMessage);
    UpdatePreview();
}
void Application::SelectFolder() {
    const std::wstring selectedFolder = BrowseForFolder();
    if (!selectedFolder.empty()) {
        m_lastExplorerFolder = selectedFolder;
        SetEditText(m_hFolderEdit, selectedFolder);
    }
}

std::wstring Application::BrowseForFolder() const {
    IFileOpenDialog* dialog = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dialog));
    if (FAILED(hr) || !dialog) {
        return L"";
    }

    DWORD options = 0;
    dialog->GetOptions(&options);
    dialog->SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST);
    dialog->SetTitle(L"Выберите папку");

    std::wstring selectedFolder;
    if (dialog->Show(m_hWnd) == S_OK) {
        IShellItem* item = nullptr;
        if (SUCCEEDED(dialog->GetResult(&item)) && item) {
            PWSTR folderPath = nullptr;
            if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &folderPath)) && folderPath) {
                selectedFolder = folderPath;
                CoTaskMemFree(folderPath);
            }
            item->Release();
        }
    }

    dialog->Release();
    return selectedFolder;
}

void Application::PrefillFolderFromExplorer() {
    std::wstring folder = ExplorerPathProvider::GetActiveExplorerPath(false);
    if (folder.empty()) {
        std::error_code currentEc;
        folder = fs::current_path(currentEc).wstring();
    }

    if (!folder.empty()) {
        std::error_code dirEc;
        if (fs::is_directory(folder, dirEc)) {
            m_lastExplorerFolder = folder;
            SetEditText(m_hFolderEdit, folder);
        }
    }
}

void Application::SyncFolderFromExplorer() {
    const std::wstring folder = ExplorerPathProvider::GetActiveExplorerPath(true);
    if (folder.empty()) {
        return;
    }

    if (PathCompareKey(folder) == PathCompareKey(m_lastExplorerFolder)) {
        return;
    }

    m_lastExplorerFolder = folder;
    if (PathCompareKey(GetEditText(m_hFolderEdit)) != PathCompareKey(folder)) {
        SetEditText(m_hFolderEdit, folder);
    }
}

std::wstring Application::GetEditText(HWND control) const {
    if (!control) {
        return L"";
    }

    const int len = GetWindowTextLengthW(control);
    if (len <= 0) {
        return L"";
    }

    std::wstring text(static_cast<size_t>(len) + 1, L'\0');
    GetWindowTextW(control, text.data(), len + 1);
    text.resize(wcslen(text.c_str()));
    return text;
}

void Application::SetEditText(HWND control, const std::wstring& text) {
    if (control) {
        SetWindowTextW(control, text.c_str());
    }
}

void Application::SetStatusText(const std::wstring& text) {
    if (m_hStatusLabel) {
        SetWindowTextW(m_hStatusLabel, text.c_str());
    }
}

void Application::UpdateHoverState(POINT clientPoint) {
    HWND hovered = nullptr;

    if (IsPointInControl(m_hBrowseButton, clientPoint)) {
        hovered = m_hBrowseButton;
    } else if (IsPointInControl(m_hRenameButton, clientPoint)) {
        hovered = m_hRenameButton;
    } else if (IsPointInControl(m_hHelpButton, clientPoint)) {
        hovered = m_hHelpButton;
    } else if (IsPointInControl(m_hRegexCheckbox, clientPoint)) {
        hovered = m_hRegexCheckbox;
    } else if (IsPointInControl(m_hIgnoreCaseCheckbox, clientPoint)) {
        hovered = m_hIgnoreCaseCheckbox;
    }

    if (hovered == m_hoveredControl) {
        return;
    }

    HWND previous = m_hoveredControl;
    m_hoveredControl = hovered;

    for (auto& pair : m_buttonHoverAlpha) {
        pair.second = pair.first == m_hoveredControl ? 1.0f : 0.0f;
    }

    if (previous) {
        InvalidateRect(previous, nullptr, TRUE);
    }
    if (m_hoveredControl) {
        InvalidateRect(m_hoveredControl, nullptr, TRUE);
    }
}

bool Application::IsPointInControl(HWND control, POINT clientPoint) const {
    if (!control || !IsWindow(control)) {
        return false;
    }

    RECT rect = {};
    GetWindowRect(control, &rect);
    ScreenToClient(m_hWnd, reinterpret_cast<LPPOINT>(&rect.left));
    ScreenToClient(m_hWnd, reinterpret_cast<LPPOINT>(&rect.right));
    return PtInRect(&rect, clientPoint) != FALSE;
}

std::wstring Application::JoinLines(const std::vector<std::wstring>& lines) {
    if (lines.empty()) {
        return L"";
    }

    std::wstring result;
    for (size_t index = 0; index < lines.size(); ++index) {
        result += lines[index];
        if (index + 1 < lines.size()) {
            result += L"\r\n";
        }
    }

    return result;
}
