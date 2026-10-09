#include "Application.h"
#include <commctrl.h>
#include <windowsx.h>
#include <algorithm>
#include <wcw/Controls.h>
#include <wcw/Runtime.h>
#include <wcw/Geometry.h>

namespace {
constexpr wchar_t SUGGESTION_CLASS[] = L"FileRenamerSuggestions";
constexpr int LIST_ID = 2301;
constexpr int DETAILS_ID = 2302;
constexpr int LIST_SCROLL_ID = 2303;
constexpr int DETAILS_SCROLL_ID = 2304;
constexpr float SCROLLBAR_WIDTH_DIP = 12;

// Complete a partially typed token, but never replace unrelated text or a
// selection. Used both for filtering the list and for inserting its choice.
DWORD TokenStart(const std::wstring& text, DWORD start, DWORD end) {
    if (start != end || start == 0 || start > text.size()) return start;
    const auto brace = text.rfind(L'{', start - 1);
    const auto dollar = text.rfind(L'$', start - 1);
    if (brace != std::wstring::npos) {
        const auto prefix = text.substr(brace, start - brace);
        if (prefix == L"{" || prefix == L"{n" ||
            (prefix.rfind(L"{n:", 0) == 0 && prefix.find_first_not_of(L"0123456789", 3) == std::wstring::npos)) {
            return static_cast<DWORD>(brace);
        }
    }
    if (dollar != std::wstring::npos) {
        size_t runStart = dollar;
        while (runStart > 0 && text[runStart - 1] == L'$') --runStart;
        if ((dollar - runStart + 1) % 2 == 0) {
            if (start == dollar + 1) return static_cast<DWORD>(dollar - 1); // $$
        } else {
            const auto tail = text.substr(dollar + 1, start - dollar - 1);
            if (tail == L"&" || tail.find_first_not_of(L"0123456789") == std::wstring::npos) {
                return static_cast<DWORD>(dollar);
            }
        }
    }
    return start;
}

DWORD TokenEnd(const std::wstring& text, DWORD start, DWORD caret) {
    size_t end = caret;
    if (text[start] == L'$') {
        const auto prefix = text.substr(start + 1, caret - start - 1);
        if (prefix.empty() && end < text.size() && (text[end] == L'&' || text[end] == L'$')) {
            ++end;
        } else if (prefix.find_first_not_of(L"0123456789") == std::wstring::npos) {
            while (end < text.size() && text[end] >= L'0' && text[end] <= L'9') ++end;
        }
    } else {
        while (end < text.size() && std::wstring(L"n:0123456789").find(text[end]) != std::wstring::npos) ++end;
        end = end < text.size() && text[end] == L'}' ? end + 1 : caret;
    }
    return static_cast<DWORD>(end);
}
} // namespace

void Application::ScheduleSuggestions() {
    if (m_applyingSuggestion) return;
    const HWND focus = GetFocus();
    if ((focus == m_hPatternEdit && m_useRegex) || focus == m_hReplacementEdit) {
        SetTimer(m_hWnd, SUGGESTION_TIMER_ID, 220, nullptr);
    } else {
        HideSuggestions();
    }
}

void Application::HideSuggestions() {
    if (m_hWnd) KillTimer(m_hWnd, SUGGESTION_TIMER_ID);
    if (m_hSuggestionWindow) ShowWindow(m_hSuggestionWindow, SW_HIDE);
    m_suggestionTarget = nullptr;
}

void Application::ShowSuggestions(HWND target) {
    KillTimer(m_hWnd, SUGGESTION_TIMER_ID);
    if (m_applyingSuggestion || (target != m_hPatternEdit && target != m_hReplacementEdit) ||
        (target == m_hPatternEdit && !m_useRegex)) {
        HideSuggestions();
        return;
    }
    const auto entries = RenamerCore::CollectOperations(GetEditText(m_hFolderEdit), L"", L"", false, false);
    const auto pattern = GetEditText(m_hPatternEdit);
    const bool patterns = target == m_hPatternEdit;
    m_suggestions = patterns ? RenamerCore::SuggestPatterns(entries.operations, m_ignoreCase)
        : RenamerCore::SuggestReplacements(pattern, m_useRegex, m_ignoreCase, entries.operations,
            pattern == m_groupLabelPattern ? m_groupLabels : std::vector<std::wstring>{});

    if (!patterns) {
        DWORD start = 0, end = 0;
        SendMessageW(target, EM_GETSEL, reinterpret_cast<WPARAM>(&start), reinterpret_cast<LPARAM>(&end));
        const auto text = GetEditText(target);
        const DWORD tokenStart = TokenStart(text, start, end);
        if (tokenStart < start) {
            const auto prefix = text.substr(tokenStart, start - tokenStart);
            m_suggestions.erase(std::remove_if(m_suggestions.begin(), m_suggestions.end(), [&](const auto& s) {
                return s.replacement.rfind(prefix, 0) != 0;
            }), m_suggestions.end());
        }
    }

    if (!m_hSuggestionWindow) {
        if (!m_widgetsInitialized) {
            if (!wcw::Initialize(m_hInstance)) {
                SetStatusText(L"Не удалось создать элементы подсказок.");
                return;
            }
            m_widgetsInitialized = true;
            wcw::SetTheme(wcw::DarkTheme());
        }
        WNDCLASSW wc = {};
        wc.lpfnWndProc = SuggestionWindowProc;
        wc.hInstance = m_hInstance;
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.lpszClassName = SUGGESTION_CLASS;
        if (!RegisterClassW(&wc) && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) return;
        m_hSuggestionWindow = CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
            SUGGESTION_CLASS, L"Подсказки переименования", WS_POPUP | WS_BORDER,
            0, 0, 0, 0, m_hWnd, nullptr, m_hInstance, this);
        if (!m_hSuggestionWindow) return;
        auto createScroll = [&](int id) {
            wcw::ScrollViewOptions options;
            options.parent = m_hSuggestionWindow;
            options.id = id;
            options.style = WS_VISIBLE;
            options.appearance.background = wcw::Color::FromRgb(45, 45, 45);
            options.appearance.focusWidthDip = 0;
            options.trackAppearance.background = wcw::Color::FromRgb(55, 55, 55);
            options.trackAppearance.trackThicknessDip = 4;
            options.thumbAppearance.accent = wcw::Color::FromRgb(125, 125, 125);
            options.thumbAppearance.thumbSizeDip = SCROLLBAR_WIDTH_DIP;
            options.thumbAppearance.cornerRadiusDip = 4;
            HWND scroll = wcw::CreateScrollView(options);
            SetWindowSubclass(scroll, SuggestionScrollProc, 1, reinterpret_cast<DWORD_PTR>(this));
            SetWindowSubclass(wcw::GetScrollContentWindow(scroll), SuggestionScrollProc, 1,
                reinterpret_cast<DWORD_PTR>(this));
            return scroll;
        };
        m_hSuggestionListScroll = createScroll(LIST_SCROLL_ID);
        m_hSuggestionDetailsScroll = createScroll(DETAILS_SCROLL_ID);
        if (!m_hSuggestionListScroll || !m_hSuggestionDetailsScroll) {
            DestroyWindow(m_hSuggestionWindow);
            m_hSuggestionWindow = nullptr;
            return;
        }
        m_hSuggestionList = CreateWindowExW(0, L"LISTBOX", L"", WS_CHILD | WS_VISIBLE |
            LBS_NOTIFY | LBS_NOINTEGRALHEIGHT, 0, 0, 0, 0, wcw::GetScrollContentWindow(m_hSuggestionListScroll),
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(LIST_ID)), m_hInstance, nullptr);
        m_hSuggestionDetails = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE |
            ES_MULTILINE | ES_READONLY, 0, 0, 0, 0, wcw::GetScrollContentWindow(m_hSuggestionDetailsScroll),
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(DETAILS_ID)), m_hInstance, nullptr);
        for (HWND control : {m_hSuggestionList, m_hSuggestionDetails}) {
            SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(m_hFont), FALSE);
            SetWindowSubclass(control, SuggestionListProc, 1, reinterpret_cast<DWORD_PTR>(this));
        }
        SendMessageW(m_hSuggestionList, LB_SETITEMHEIGHT, 0, 24);
    }
    m_suggestionTarget = target;
    SendMessageW(m_hSuggestionList, WM_SETREDRAW, FALSE, 0);
    SendMessageW(m_hSuggestionList, LB_RESETCONTENT, 0, 0);
    for (const auto& s : m_suggestions) {
        const std::wstring label = s.title + (patterns ? L"  ·  " + std::to_wstring(s.matchCount) +
            L" из " + std::to_wstring(s.totalCount) : L"");
        SendMessageW(m_hSuggestionList, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(label.c_str()));
    }
    if (m_suggestions.empty()) {
        SendMessageW(m_hSuggestionList, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Нет подходящих подсказок"));
    }
    SendMessageW(m_hSuggestionList, LB_SETCURSEL, 0, 0);
    SendMessageW(m_hSuggestionList, WM_SETREDRAW, TRUE, 0);

    RECT anchor = {};
    GetWindowRect(target, &anchor);
    MONITORINFO monitor = {sizeof(MONITORINFO)};
    GetMonitorInfoW(MonitorFromWindow(target, MONITOR_DEFAULTTONEAREST), &monitor);
    const int width = std::min(std::max(620, static_cast<int>(anchor.right - anchor.left + 110)),
        static_cast<int>(monitor.rcWork.right - monitor.rcWork.left));
    const int height = std::min(370, static_cast<int>(monitor.rcWork.bottom - monitor.rcWork.top));
    const int x = std::clamp(static_cast<int>(anchor.left), static_cast<int>(monitor.rcWork.left),
        static_cast<int>(monitor.rcWork.right) - width);
    const int y = std::clamp(static_cast<int>(anchor.bottom + 3), static_cast<int>(monitor.rcWork.top),
        static_cast<int>(monitor.rcWork.bottom) - height);
    const unsigned dpi = GetDpiForWindow(m_hSuggestionWindow);
    const int listHeight = std::max(1, static_cast<int>(m_suggestions.size())) * 24;
    const int listBar = listHeight > 120 ? wcw::DipToPx(SCROLLBAR_WIDTH_DIP, dpi) : 0;
    MoveWindow(m_hSuggestionListScroll, 10, 10, width - 22, 120, TRUE);
    MoveWindow(m_hSuggestionList, 0, 0, width - 22 - listBar, std::max(120, listHeight), TRUE);
    wcw::SetScrollContentExtent(m_hSuggestionListScroll, {0, wcw::PxToDip(listHeight, dpi)});
    wcw::SetScrollOffset(m_hSuggestionListScroll, {});
    MoveWindow(m_hSuggestionDetailsScroll, 12, 142, width - 26, height - 154, TRUE);
    UpdateSuggestionDetails();
    SetWindowPos(m_hSuggestionWindow, HWND_TOP, x, y, width, height, SWP_NOACTIVATE | SWP_SHOWWINDOW);
    InvalidateRect(m_hSuggestionWindow, nullptr, TRUE);
}

void Application::UpdateSuggestionDetails() {
    const auto selected = SendMessageW(m_hSuggestionList, LB_GETCURSEL, 0, 0);
    if (selected == LB_ERR || static_cast<size_t>(selected) >= m_suggestions.size()) {
        SetEditText(m_hSuggestionDetails,
            L"Для подбора паттерна выберите папку с похожими именами. Предлагаются сезоны, серии, числовые части и частые замены.\r\n"
            L"В поле замены: Ctrl+Space — счётчики и доступные группы. Удалите незавершённый токен, чтобы увидеть весь список.\r\n\r\nEsc — закрыть.");
        LayoutSuggestionDetails();
        return;
    }
    const auto& s = m_suggestions[static_cast<size_t>(selected)];
    const bool patterns = m_suggestionTarget == m_hPatternEdit;
    std::wstring text = patterns
        ? L"Enter / двойной щелчок — заполнить паттерн и замену; Esc — закрыть.\r\n"
        : L"Enter / двойной щелчок — вставить в позицию курсора; Esc — закрыть.\r\n";
    text += s.description + L"\r\n\r\n";
    if (patterns) {
        text += L"Паттерн: " + s.pattern + L"\r\nЗамена: " +
            (s.replacement.empty() ? L"(пусто — удалить совпадение)" : s.replacement) + L"\r\n";
    }
    if (!s.exampleBefore.empty()) text += s.exampleBefore + L"  →  " + s.exampleAfter + L"\r\n";
    if (!s.unmatchedExample.empty()) text += L"Не подходит: " + s.unmatchedExample;
    SetEditText(m_hSuggestionDetails, text);
    LayoutSuggestionDetails();

    // The list itself is full-height; the custom viewport scrolls to the selection.
    const unsigned dpi = GetDpiForWindow(m_hSuggestionListScroll);
    RECT viewport = {};
    GetClientRect(m_hSuggestionListScroll, &viewport);
    const auto offset = wcw::GetScrollOffset(m_hSuggestionListScroll).value_or(wcw::ScrollOffsetDip{});
    float y = offset.y;
    const float top = wcw::PxToDip(static_cast<int>(selected) * 24, dpi);
    const float bottom = top + wcw::PxToDip(24, dpi);
    const float page = wcw::PxToDip(viewport.bottom, dpi);
    if (top < y) y = top;
    else if (bottom > y + page) y = bottom - page;
    wcw::SetScrollOffset(m_hSuggestionListScroll, {0, y});
}

void Application::LayoutSuggestionDetails() {
    if (!m_hSuggestionDetails || !m_hSuggestionDetailsScroll) return;
    RECT bounds = {};
    GetClientRect(m_hSuggestionDetailsScroll, &bounds);
    if (bounds.right <= 0 || bounds.bottom <= 0) return;
    HDC dc = GetDC(m_hSuggestionDetails);
    const auto oldFont = SelectObject(dc, m_hFont);
    TEXTMETRICW metrics = {};
    GetTextMetricsW(dc, &metrics);
    SelectObject(dc, oldFont);
    ReleaseDC(m_hSuggestionDetails, dc);
    const unsigned dpi = GetDpiForWindow(m_hSuggestionDetailsScroll);
    int width = bounds.right;
    auto measure = [&] {
        MoveWindow(m_hSuggestionDetails, 0, 0, width, bounds.bottom, FALSE);
        return static_cast<int>(SendMessageW(m_hSuggestionDetails, EM_GETLINECOUNT, 0, 0)) *
            std::max(1L, metrics.tmHeight) + 4;
    };
    int contentHeight = measure();
    if (contentHeight > bounds.bottom) {
        width = std::max(1, width - wcw::DipToPx(SCROLLBAR_WIDTH_DIP, dpi));
        contentHeight = measure();
    }
    MoveWindow(m_hSuggestionDetails, 0, 0, width, contentHeight, TRUE);
    wcw::SetScrollContentExtent(m_hSuggestionDetailsScroll, {0, wcw::PxToDip(contentHeight, dpi)});
    wcw::SetScrollOffset(m_hSuggestionDetailsScroll, {});
}

void Application::ApplySuggestion() {
    const auto selected = SendMessageW(m_hSuggestionList, LB_GETCURSEL, 0, 0);
    if (selected == LB_ERR || static_cast<size_t>(selected) >= m_suggestions.size()) {
        HideSuggestions();
        return;
    }
    const auto s = m_suggestions[static_cast<size_t>(selected)];
    const HWND target = m_suggestionTarget;
    m_applyingSuggestion = true;
    if (target == m_hPatternEdit) {
        SetEditText(m_hPatternEdit, s.pattern);
        SetEditText(m_hReplacementEdit, s.replacement);
        m_groupLabelPattern = s.pattern;
        m_groupLabels = s.groups;
        SetFocus(m_hReplacementEdit);
        // Select the placeholder title so typing a name replaces it immediately.
        const size_t placeholder = s.replacement.rfind(L"Название", 0) == 0 ? 8 : 0;
        SendMessageW(m_hReplacementEdit, EM_SETSEL, 0, static_cast<LPARAM>(placeholder));
    } else if (target == m_hReplacementEdit) {
        DWORD start = 0, end = 0;
        SendMessageW(target, EM_GETSEL, reinterpret_cast<WPARAM>(&start), reinterpret_cast<LPARAM>(&end));
        const auto text = GetEditText(target);
        const DWORD tokenStart = TokenStart(text, start, end);
        if (tokenStart < start) end = TokenEnd(text, tokenStart, start);
        start = tokenStart;
        SendMessageW(target, EM_SETSEL, start, end);
        SendMessageW(target, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(s.replacement.c_str()));
    }
    HideSuggestions();
    m_applyingSuggestion = false;
    UpdatePreview();
}

bool Application::HandleSuggestionKey(WPARAM key) {
    const HWND focus = GetFocus();
    if (key == VK_SPACE && (GetKeyState(VK_CONTROL) & 0x8000) &&
        (focus == m_hPatternEdit || focus == m_hReplacementEdit)) {
        ShowSuggestions(focus);
        return true;
    }
    if (!m_hSuggestionWindow || !IsWindowVisible(m_hSuggestionWindow) || focus != m_suggestionTarget) return false;
    if (key == VK_ESCAPE) {
        HideSuggestions();
        return true;
    }
    if (key == VK_RETURN) {
        ApplySuggestion();
        return true;
    }
    if (key == VK_UP || key == VK_DOWN) {
        const int count = static_cast<int>(m_suggestions.size());
        if (count > 0) {
            const int selected = static_cast<int>(SendMessageW(m_hSuggestionList, LB_GETCURSEL, 0, 0));
            const int next = (selected + (key == VK_DOWN ? 1 : count - 1)) % count;
            SendMessageW(m_hSuggestionList, LB_SETCURSEL, next, 0);
            UpdateSuggestionDetails();
        }
        return true;
    }
    if (key == VK_TAB) HideSuggestions();
    return false;
}

LRESULT CALLBACK Application::SuggestionWindowProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam) {
    auto* app = reinterpret_cast<Application*>(GetWindowLongPtrW(window, GWLP_USERDATA));
    if (message == WM_NCCREATE) {
        app = static_cast<Application*>(reinterpret_cast<CREATESTRUCTW*>(lParam)->lpCreateParams);
        SetWindowLongPtrW(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(app));
    }
    if (!app) return DefWindowProcW(window, message, wParam, lParam);
    switch (message) {
    case WM_MOUSEACTIVATE: return MA_NOACTIVATE;
    case WM_ERASEBKGND:
        {
            RECT rect;
            GetClientRect(window, &rect);
            FillRect(reinterpret_cast<HDC>(wParam), &rect, app->m_hCardBrush);
        }
        return 1;
    case WM_CTLCOLORLISTBOX:
    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORSTATIC:
        SetTextColor(reinterpret_cast<HDC>(wParam), RGB(243, 244, 246));
        SetBkColor(reinterpret_cast<HDC>(wParam), RGB(45, 45, 45));
        return reinterpret_cast<LRESULT>(app->m_hCardBrush);
    case WM_COMMAND:
        if (LOWORD(wParam) == LIST_ID && HIWORD(wParam) == LBN_SELCHANGE) app->UpdateSuggestionDetails();
        return 0;
    }
    return DefWindowProcW(window, message, wParam, lParam);
}

LRESULT CALLBACK Application::SuggestionListProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam,
    UINT_PTR, DWORD_PTR data) {
    auto* app = reinterpret_cast<Application*>(data);
    if (message == WM_MOUSEACTIVATE) return MA_NOACTIVATE;
    if (message == WM_LBUTTONDOWN || message == WM_LBUTTONDBLCLK) {
        if (window == app->m_hSuggestionList) {
            const auto hit = SendMessageW(window, LB_ITEMFROMPOINT, 0, lParam);
            if (!HIWORD(hit)) {
                SendMessageW(window, LB_SETCURSEL, LOWORD(hit), 0);
                app->UpdateSuggestionDetails();
                if (message == WM_LBUTTONDBLCLK) app->ApplySuggestion();
            }
        }
        // Keep the caret/selection in the original edit control.
        return 0;
    }
    if (message == WM_MOUSEWHEEL) {
        SendMessageW(window == app->m_hSuggestionDetails ? app->m_hSuggestionDetailsScroll : app->m_hSuggestionListScroll,
            message, wParam, lParam);
        return 0;
    }
    return DefSubclassProc(window, message, wParam, lParam);
}

LRESULT CALLBACK Application::SuggestionScrollProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam,
    UINT_PTR, DWORD_PTR data) {
    auto* app = reinterpret_cast<Application*>(data);
    if (message == WM_MOUSEACTIVATE) return MA_NOACTIVATE;
    if (message == WM_CTLCOLORLISTBOX || message == WM_CTLCOLOREDIT || message == WM_CTLCOLORSTATIC || message == WM_COMMAND) {
        return SendMessageW(app->m_hSuggestionWindow, message, wParam, lParam);
    }
    if (message == WM_LBUTTONDOWN &&
        (window == app->m_hSuggestionListScroll || window == app->m_hSuggestionDetailsScroll)) {
        // WCW focuses its scrollbar when dragging. Keep completion in the edit
        // field and suppress dismissal during this brief focus transfer.
        const bool wasApplying = app->m_applyingSuggestion;
        app->m_applyingSuggestion = true;
        const auto result = DefSubclassProc(window, message, wParam, lParam);
        if (app->m_suggestionTarget) SetFocus(app->m_suggestionTarget);
        app->m_applyingSuggestion = wasApplying;
        return result;
    }
    const auto result = DefSubclassProc(window, message, wParam, lParam);
    if (message == WM_SIZE && window == app->m_hSuggestionDetailsScroll) app->LayoutSuggestionDetails();
    return result;
}
