#include "Application.h"

#include <windows.h>

int APIENTRY wWinMain(
    _In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR lpCmdLine,
    _In_ int nCmdShow
) {
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    Application app;
    if (!app.Initialize(hInstance)) {
        MessageBox(nullptr, L"Не удалось инициализировать приложение", L"Ошибка", MB_OK | MB_ICONERROR);
        return -1;
    }

    HWND hWnd = app.GetMainWindow();
    ShowWindow(hWnd, nCmdShow == SW_SHOWMINIMIZED ? SW_SHOWNORMAL : nCmdShow);
    UpdateWindow(hWnd);

    return app.Run();
}
