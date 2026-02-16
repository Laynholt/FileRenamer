#include "ExplorerPathProvider.h"

#include <windows.h>
#include <exdisp.h>
#include <objbase.h>
#include <shlwapi.h>

#include <filesystem>
#include <iterator>

namespace fs = std::filesystem;

namespace ExplorerPathProvider {

std::wstring GetActiveExplorerPath(bool activeOnly) {
    IShellWindows* shellWindows = nullptr;
    const HRESULT shellResult = CoCreateInstance(CLSID_ShellWindows, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&shellWindows));
    if (FAILED(shellResult) || !shellWindows) {
        return L"";
    }

    long count = 0;
    shellWindows->get_Count(&count);

    const HWND foregroundWindow = GetForegroundWindow();
    const HWND foregroundRoot = foregroundWindow ? GetAncestor(foregroundWindow, GA_ROOT) : nullptr;
    const HWND foregroundRootOwner = foregroundWindow ? GetAncestor(foregroundWindow, GA_ROOTOWNER) : nullptr;
    std::wstring fallbackPath;

    for (long index = 0; index < count; ++index) {
        VARIANT itemIndex;
        VariantInit(&itemIndex);
        itemIndex.vt = VT_I4;
        itemIndex.lVal = index;

        IDispatch* dispatch = nullptr;
        const HRESULT itemResult = shellWindows->Item(itemIndex, &dispatch);
        VariantClear(&itemIndex);
        if (FAILED(itemResult) || !dispatch) {
            continue;
        }

        IWebBrowserApp* browserApp = nullptr;
        if (FAILED(dispatch->QueryInterface(IID_PPV_ARGS(&browserApp))) || !browserApp) {
            dispatch->Release();
            continue;
        }
        dispatch->Release();

        SHANDLE_PTR browserHwnd = 0;
        browserApp->get_HWND(&browserHwnd);

        BSTR locationUrl = nullptr;
        browserApp->get_LocationURL(&locationUrl);
        browserApp->Release();

        if (!locationUrl || SysStringLen(locationUrl) == 0) {
            if (locationUrl) {
                SysFreeString(locationUrl);
            }
            continue;
        }

        wchar_t pathBuffer[4096] = {};
        DWORD pathBufferSize = static_cast<DWORD>(std::size(pathBuffer));

        std::wstring folderPath;
        if (SUCCEEDED(PathCreateFromUrlW(locationUrl, pathBuffer, &pathBufferSize, 0))) {
            folderPath = pathBuffer;
        } else {
            folderPath = locationUrl;
        }

        SysFreeString(locationUrl);

        std::error_code dirEc;
        if (!fs::is_directory(folderPath, dirEc)) {
            continue;
        }

        const HWND explorerHwnd = reinterpret_cast<HWND>(browserHwnd);
        const bool isForegroundExplorer =
            explorerHwnd == foregroundWindow ||
            explorerHwnd == foregroundRoot ||
            explorerHwnd == foregroundRootOwner ||
            (foregroundWindow && IsChild(explorerHwnd, foregroundWindow));

        if (isForegroundExplorer) {
            shellWindows->Release();
            return folderPath;
        }

        if (fallbackPath.empty()) {
            fallbackPath = folderPath;
        }
    }

    shellWindows->Release();
    if (!activeOnly) {
        return fallbackPath;
    }

    return L"";
}

} // namespace ExplorerPathProvider
