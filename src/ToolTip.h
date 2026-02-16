#pragma once

#include <windows.h>

#include <map>
#include <string>

class ToolTip {
public:
    ToolTip();
    ~ToolTip();

    ToolTip(const ToolTip&) = delete;
    ToolTip& operator=(const ToolTip&) = delete;

    bool Initialize(HWND ownerWindow);
    void SetStyle(HFONT font, COLORREF backgroundColor, COLORREF textColor) const;
    bool AddTool(HWND control, const std::wstring& text);
    void RelayEvent(const MSG& message) const;
    void Destroy();

private:
    HWND m_hToolTip;
    std::map<HWND, std::wstring> m_textByControl;
};
