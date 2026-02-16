#include "ToolTip.h"

#include <commctrl.h>

ToolTip::ToolTip()
    : m_hToolTip(nullptr) {
}

ToolTip::~ToolTip() {
    Destroy();
}

bool ToolTip::Initialize(HWND ownerWindow) {
    if (!ownerWindow || !IsWindow(ownerWindow)) {
        return false;
    }

    Destroy();

    m_hToolTip = CreateWindowExW(
        WS_EX_TOPMOST,
        TOOLTIPS_CLASSW,
        nullptr,
        WS_POPUP | TTS_ALWAYSTIP | TTS_NOPREFIX,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        ownerWindow,
        nullptr,
        GetModuleHandleW(nullptr),
        nullptr
    );

    if (!m_hToolTip) {
        return false;
    }

    SetWindowPos(m_hToolTip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    SendMessageW(m_hToolTip, TTM_ACTIVATE, TRUE, 0);
    SendMessageW(m_hToolTip, TTM_SETDELAYTIME, TTDT_INITIAL, 350);
    SendMessageW(m_hToolTip, TTM_SETMAXTIPWIDTH, 0, 420);

    return true;
}

void ToolTip::SetStyle(HFONT font, COLORREF backgroundColor, COLORREF textColor) const {
    if (!m_hToolTip || !IsWindow(m_hToolTip)) {
        return;
    }

    SendMessageW(m_hToolTip, TTM_SETTIPBKCOLOR, static_cast<WPARAM>(backgroundColor), 0);
    SendMessageW(m_hToolTip, TTM_SETTIPTEXTCOLOR, static_cast<WPARAM>(textColor), 0);

    if (font) {
        SendMessageW(m_hToolTip, WM_SETFONT, reinterpret_cast<WPARAM>(font), MAKELPARAM(TRUE, 0));
    }
}

bool ToolTip::AddTool(HWND control, const std::wstring& text) {
    if (!m_hToolTip || !control || !IsWindow(control) || text.empty()) {
        return false;
    }

    HWND parentWindow = GetParent(control);
    if (!parentWindow) {
        return false;
    }

    m_textByControl[control] = text;

    TOOLINFOW toolInfo = {};
    toolInfo.cbSize = TTTOOLINFOW_V2_SIZE;
    toolInfo.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
    toolInfo.hwnd = parentWindow;
    toolInfo.uId = reinterpret_cast<UINT_PTR>(control);
    toolInfo.lpszText = const_cast<LPWSTR>(m_textByControl[control].c_str());

    SendMessageW(m_hToolTip, TTM_DELTOOLW, 0, reinterpret_cast<LPARAM>(&toolInfo));
    return SendMessageW(m_hToolTip, TTM_ADDTOOLW, 0, reinterpret_cast<LPARAM>(&toolInfo)) != FALSE;
}

void ToolTip::RelayEvent(const MSG& message) const {
    if (!m_hToolTip || !IsWindow(m_hToolTip)) {
        return;
    }

    switch (message.message) {
    case WM_MOUSEMOVE:
    case WM_LBUTTONDOWN:
    case WM_LBUTTONUP:
    case WM_RBUTTONDOWN:
    case WM_RBUTTONUP:
    case WM_MBUTTONDOWN:
    case WM_MBUTTONUP:
    case WM_MOUSEWHEEL:
    case WM_NCMOUSEMOVE:
    case WM_NCLBUTTONDOWN:
    case WM_NCLBUTTONUP:
    case WM_NCRBUTTONDOWN:
    case WM_NCRBUTTONUP:
    case WM_NCMBUTTONDOWN:
    case WM_NCMBUTTONUP:
        {
            MSG relayMessage = message;
            SendMessageW(m_hToolTip, TTM_RELAYEVENT, 0, reinterpret_cast<LPARAM>(&relayMessage));
        }
        break;
    default:
        break;
    }
}

void ToolTip::Destroy() {
    if (m_hToolTip && IsWindow(m_hToolTip)) {
        DestroyWindow(m_hToolTip);
    }

    m_hToolTip = nullptr;
    m_textByControl.clear();
}
