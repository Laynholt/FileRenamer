#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <windows.h>
#include <gdiplus.h>

#include "RenamerService.h"

#include <map>
#include <memory>
#include <string>
#include <vector>

class UpdateService;

class Application {
public:
    static constexpr const wchar_t* WINDOW_TITLE = L"FileRenamer";

    Application();
    ~Application();

    bool Initialize(HINSTANCE hInstance);
    int Run();
    void Shutdown();

    HWND GetMainWindow() const { return m_hWnd; }

private:
    enum class InfoWindowKind {
        Hotkeys,
        About
    };

    static LRESULT CALLBACK WindowProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam);
    LRESULT HandleMessage(UINT message, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK TextEditSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);

    void CreateControls();
    void CreateHelpMenu();
    void OnResize(int width, int height);
    void OnPaint();
    void OnCommand(UINT controlId, UINT notifyCode);
    void OnMenuCommand(UINT menuId);

    void UpdatePreview();
    void RenameFiles();

    void SelectFolder();
    std::wstring BrowseForFolder() const;

    void PrefillFolderFromExplorer();
    void SyncFolderFromExplorer();

    bool RegisterInfoWindowClass();
    bool RegisterMessageWindowClass();
    void ShowHelpMenu();
    void ShowHotkeysWindow();
    void ShowAboutWindow();
    void CreateOrActivateInfoWindow(InfoWindowKind kind, HWND& targetHandle, const wchar_t* title, const std::wstring& bodyText);
    void OnInfoWindowClosed(InfoWindowKind kind);
    void CheckForUpdates();
    int ShowStyledMessageDialog(const wchar_t* title,
                                const std::wstring& bodyText,
                                const wchar_t* primaryButtonText,
                                const wchar_t* secondaryButtonText = nullptr);
    void ShowStyledMessage(const std::wstring& title, const std::wstring& message);
    void ShowTextContextMenu(HWND targetControl, LPARAM lParam);
    static LRESULT CALLBACK InfoWindowProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK MessageWindowProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam);

    std::wstring GetEditText(HWND control) const;
    void SetEditText(HWND control, const std::wstring& text);
    void SetStatusText(const std::wstring& text);

    void UpdateHoverState(POINT clientPoint);
    bool IsPointInControl(HWND control, POINT clientPoint) const;

    static std::wstring JoinLines(const std::vector<std::wstring>& lines);

    HINSTANCE m_hInstance;
    HWND m_hWnd;

    HWND m_hFolderLabel;
    HWND m_hFolderEdit;
    HWND m_hBrowseButton;

    HWND m_hPatternLabel;
    HWND m_hPatternEdit;

    HWND m_hReplacementLabel;
    HWND m_hReplacementEdit;

    HWND m_hRegexCheckbox;
    HWND m_hIgnoreCaseCheckbox;
    HWND m_hRenameButton;
    HWND m_hHelpButton;

    HWND m_hStatusLabel;

    HWND m_hCurrentLabel;
    HWND m_hResultLabel;
    HWND m_hCurrentPreview;
    HWND m_hResultPreview;
    HWND m_hHotkeysWindow;
    HWND m_hAboutWindow;
    HMENU m_hHelpMenu;

    HBRUSH m_hBackgroundBrush;
    HBRUSH m_hCardBrush;

    HFONT m_hFont;
    HFONT m_hMonoFont;

    ULONG_PTR m_gdiplusToken;
    bool m_comInitialized;
    bool m_useRegex;
    bool m_ignoreCase;
    bool m_infoWindowClassRegistered;
    bool m_messageWindowClassRegistered;

    HWND m_hoveredControl;
    HWND m_pressedControl;
    std::map<HWND, float> m_buttonHoverAlpha;
    std::unique_ptr<UpdateService> m_updateService;

    std::wstring m_lastExplorerFolder;

    static constexpr int PREVIEW_LIMIT = 400;
    static constexpr UINT_PTR EXPLORER_SYNC_TIMER_ID = 1;
    static constexpr UINT EXPLORER_SYNC_INTERVAL_MS = 300;
    static constexpr int MIN_WINDOW_WIDTH = 860;
    static constexpr int MIN_WINDOW_HEIGHT = 620;
};
