#pragma once

#include <windows.h>
#include <string>

class UiRenderer {
public:
    static void DrawCustomButton(HDC hdc, HWND button, const std::wstring& text, bool isPressed, float hoverAlpha);
    static void DrawCustomCheckbox(HDC hdc, HWND control, const std::wstring& text, bool checked, bool hot, bool pressed, bool enabled, bool focused);
    static void DrawBackground(HDC hdc, const RECT& rect);
    static void DrawCard(HDC hdc, const RECT& rect, const std::wstring& title = L"");
    static void DrawEditBorder(HWND parentWindow, HWND editControl);
};
