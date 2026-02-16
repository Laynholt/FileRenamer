
#include "Application.h"

#include "ExplorerPathProvider.h"
#include "RenamerService.h"
#include "UiRenderer.h"
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
const wchar_t* STYLED_DIALOG_CLASS_NAME = L"FileRenamerStyledDialogClass";

enum ControlId {
    ID_FOLDER_EDIT = 1001,
    ID_BROWSE_BUTTON = 1002,
    ID_PATTERN_EDIT = 1003,
    ID_REPLACEMENT_EDIT = 1004,
    ID_REGEX_CHECKBOX = 1005,
    ID_IGNORE_CASE_CHECKBOX = 1006,
    ID_RENAME_BUTTON = 1007,
    ID_CURRENT_PREVIEW = 1008,
    ID_RESULT_PREVIEW = 1009
};

enum StyledDialogControlId {
    ID_STYLED_DIALOG_TEXT = 2001,
    ID_STYLED_DIALOG_OK = 2002
};

struct StyledDialogState {
    Application* owner;
    std::wstring message;
    HFONT font;
    HBRUSH editBrush;
    HWND textControl;
    HWND okButton;
};

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
    , m_hStatusLabel(nullptr)
    , m_hCurrentLabel(nullptr)
    , m_hResultLabel(nullptr)
    , m_hCurrentPreview(nullptr)
    , m_hResultPreview(nullptr)
    , m_hBackgroundBrush(nullptr)
    , m_hCardBrush(nullptr)
    , m_hFont(nullptr)
    , m_hMonoFont(nullptr)
    , m_gdiplusToken(0)
    , m_comInitialized(false)
    , m_useRegex(false)
    , m_ignoreCase(false)
    , m_styledDialogClassRegistered(false)
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
    if (!RegisterStyledDialogClass()) {
        MessageBox(nullptr, L"Не удалось зарегистрировать класс диалога сообщений", L"Ошибка", MB_OK | MB_ICONERROR);
        return false;
    }

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
        OnCommand(LOWORD(wParam), HIWORD(wParam));
        return 0;

    case WM_DRAWITEM:
        {
            DRAWITEMSTRUCT* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
            if (!dis || dis->CtlType != ODT_BUTTON) {
                break;
            }

            if (dis->CtlID == ID_REGEX_CHECKBOX || dis->CtlID == ID_IGNORE_CASE_CHECKBOX) {
                wchar_t text[256] = {};
                GetWindowText(dis->hwndItem, text, 256);
                const bool isPressed = m_pressedControl == dis->hwndItem || (dis->itemState & ODS_SELECTED) != 0;
                const bool hasFocus = (dis->itemState & ODS_FOCUS) != 0;
                const bool enabled = (dis->itemState & ODS_DISABLED) == 0;
                const bool isHot = m_hoveredControl == dis->hwndItem;
                const bool checked = (dis->CtlID == ID_REGEX_CHECKBOX) ? m_useRegex : m_ignoreCase;
                UiRenderer::DrawCustomCheckbox(dis->hDC, dis->hwndItem, text, checked, isHot, isPressed, enabled, hasFocus);
                return TRUE;
            }

            wchar_t text[256] = {};
            GetWindowText(dis->hwndItem, text, 256);
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
}

void Application::OnResize(int width, int height) {
    if (width <= 0 || height <= 0) {
        return;
    }

    const int outerMargin = 12;
    const int cardPadding = 16;
    const int cardGap = 12;

    const int formCardTop = outerMargin;
    const int formCardHeight = 220;
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

    MoveWindow(m_hStatusLabel, contentLeft, actionRowY + 36, contentRight - contentLeft, 22, TRUE);

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
    const int formCardHeight = 220;
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

    default:
        break;
    }
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

bool Application::RegisterStyledDialogClass() {
    if (m_styledDialogClassRegistered) {
        return true;
    }

    WNDCLASSEX wcex = {};
    wcex.cbSize = sizeof(wcex);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = StyledDialogProc;
    wcex.hInstance = m_hInstance;
    wcex.hIcon = LoadIcon(m_hInstance, MAKEINTRESOURCE(IDI_MAIN_ICON));
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = nullptr;
    wcex.lpszClassName = STYLED_DIALOG_CLASS_NAME;
    wcex.hIconSm = LoadIcon(m_hInstance, MAKEINTRESOURCE(IDI_MAIN_ICON));

    if (!RegisterClassEx(&wcex) && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
        return false;
    }

    m_styledDialogClassRegistered = true;
    return true;
}

void Application::ShowStyledMessage(const std::wstring& title, const std::wstring& message) {
    if (!RegisterStyledDialogClass()) {
        MessageBox(m_hWnd, message.c_str(), title.c_str(), MB_OK | MB_ICONINFORMATION);
        return;
    }

    const int dialogWidth = 520;
    const int dialogHeight = 280;

    RECT parentRect = {};
    GetWindowRect(m_hWnd, &parentRect);
    const int x = parentRect.left + ((parentRect.right - parentRect.left) - dialogWidth) / 2;
    const int y = parentRect.top + ((parentRect.bottom - parentRect.top) - dialogHeight) / 2;

    auto* state = new StyledDialogState{
        this,
        message,
        m_hFont,
        nullptr,
        nullptr,
        nullptr
    };

    EnableWindow(m_hWnd, FALSE);

    HWND dialog = CreateWindowEx(
        WS_EX_DLGMODALFRAME,
        STYLED_DIALOG_CLASS_NAME,
        title.c_str(),
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
        x,
        y,
        dialogWidth,
        dialogHeight,
        m_hWnd,
        nullptr,
        m_hInstance,
        state
    );

    if (!dialog) {
        EnableWindow(m_hWnd, TRUE);
        delete state;
        MessageBox(m_hWnd, message.c_str(), title.c_str(), MB_OK | MB_ICONINFORMATION);
        return;
    }

    ShowWindow(dialog, SW_SHOW);
    UpdateWindow(dialog);
    SetForegroundWindow(dialog);

    MSG msg = {};
    BOOL getMessageResult = TRUE;
    while (IsWindow(dialog) && (getMessageResult = GetMessage(&msg, nullptr, 0, 0)) != 0) {
        if (!IsDialogMessage(dialog, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    EnableWindow(m_hWnd, TRUE);
    SetForegroundWindow(m_hWnd);
    SetFocus(m_hWnd);

    if (getMessageResult == 0) {
        PostQuitMessage(static_cast<int>(msg.wParam));
    }
}

LRESULT CALLBACK Application::StyledDialogProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    auto* state = reinterpret_cast<StyledDialogState*>(GetWindowLongPtr(hWnd, GWLP_USERDATA));

    switch (message) {
    case WM_NCCREATE:
        {
            CREATESTRUCT* create = reinterpret_cast<CREATESTRUCT*>(lParam);
            auto* initialState = reinterpret_cast<StyledDialogState*>(create->lpCreateParams);
            SetWindowLongPtr(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(initialState));
            return TRUE;
        }

    case WM_CREATE:
        if (state) {
            state->editBrush = CreateSolidBrush(RGB(45, 45, 45));

            state->textControl = CreateWindowEx(
                0,
                L"EDIT",
                state->message.c_str(),
                WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
                16,
                16,
                100,
                100,
                hWnd,
                reinterpret_cast<HMENU>(ID_STYLED_DIALOG_TEXT),
                GetModuleHandle(nullptr),
                nullptr
            );

            state->okButton = CreateWindowEx(
                0,
                L"BUTTON",
                L"OK",
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON | BS_OWNERDRAW,
                16,
                16,
                120,
                32,
                hWnd,
                reinterpret_cast<HMENU>(ID_STYLED_DIALOG_OK),
                GetModuleHandle(nullptr),
                nullptr
            );

            if (state->font) {
                SendMessage(state->textControl, WM_SETFONT, reinterpret_cast<WPARAM>(state->font), MAKELPARAM(FALSE, 0));
                SendMessage(state->okButton, WM_SETFONT, reinterpret_cast<WPARAM>(state->font), MAKELPARAM(FALSE, 0));
            }
        }
        return 0;

    case WM_SIZE:
        if (state) {
            const int width = LOWORD(lParam);
            const int height = HIWORD(lParam);
            const int margin = 20;
            const int buttonWidth = 120;
            const int buttonHeight = 32;

            MoveWindow(state->textControl, margin, margin, width - (margin * 2), height - (margin * 3) - buttonHeight, TRUE);
            MoveWindow(state->okButton, width - margin - buttonWidth, height - margin - buttonHeight, buttonWidth, buttonHeight, TRUE);
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
            SetTextColor(hdcControl, RGB(243, 244, 246));
            SetBkColor(hdcControl, RGB(45, 45, 45));
            return reinterpret_cast<INT_PTR>(state->editBrush);
        }
        break;

    case WM_DRAWITEM:
        {
            DRAWITEMSTRUCT* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
            if (dis && dis->CtlType == ODT_BUTTON && dis->CtlID == ID_STYLED_DIALOG_OK) {
                const bool isPressed = (dis->itemState & ODS_SELECTED) != 0;
                const float hoverAlpha = (dis->itemState & ODS_HOTLIGHT) ? 1.0f : 0.0f;
                UiRenderer::DrawCustomButton(dis->hDC, dis->hwndItem, L"OK", isPressed, hoverAlpha);
                return TRUE;
            }
        }
        break;

    case WM_COMMAND:
        if (LOWORD(wParam) == ID_STYLED_DIALOG_OK || LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) {
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
            delete state;
            SetWindowLongPtr(hWnd, GWLP_USERDATA, 0);
        }
        return DefWindowProc(hWnd, message, wParam, lParam);
    }

    return DefWindowProc(hWnd, message, wParam, lParam);
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
