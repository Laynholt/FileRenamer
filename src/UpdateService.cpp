#include "UpdateService.h"

#include <windows.h>
#include <winhttp.h>

#include <algorithm>
#include <cwctype>
#include <sstream>
#include <vector>

namespace {
constexpr wchar_t kGitHubHost[] = L"github.com";
constexpr wchar_t kLatestReleasePath[] = L"/Laynholt/FileRenamer/releases/latest";
constexpr wchar_t kReleaseDownloadPrefix[] = L"/Laynholt/FileRenamer/releases/download/";
constexpr wchar_t kReleaseExeName[] = L"FileRenamer.exe";
constexpr wchar_t kUserAgent[] = L"FileRenamer-Updater/1.0";

class WinHttpHandle {
public:
    explicit WinHttpHandle(HINTERNET handle = nullptr)
        : m_handle(handle) {
    }

    ~WinHttpHandle() {
        if (m_handle) {
            WinHttpCloseHandle(m_handle);
            m_handle = nullptr;
        }
    }

    WinHttpHandle(const WinHttpHandle&) = delete;
    WinHttpHandle& operator=(const WinHttpHandle&) = delete;

    WinHttpHandle(WinHttpHandle&& other) noexcept
        : m_handle(other.m_handle) {
        other.m_handle = nullptr;
    }

    WinHttpHandle& operator=(WinHttpHandle&& other) noexcept {
        if (this == &other) {
            return *this;
        }
        if (m_handle) {
            WinHttpCloseHandle(m_handle);
        }
        m_handle = other.m_handle;
        other.m_handle = nullptr;
        return *this;
    }

    HINTERNET get() const {
        return m_handle;
    }

    explicit operator bool() const {
        return m_handle != nullptr;
    }

private:
    HINTERNET m_handle;
};

std::wstring FormatWin32Error(DWORD errorCode) {
    wchar_t* buffer = nullptr;
    const DWORD chars = FormatMessageW(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        nullptr,
        errorCode,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        reinterpret_cast<LPWSTR>(&buffer),
        0,
        nullptr
    );

    if (chars == 0 || !buffer) {
        std::wstringstream stream;
        stream << L"Код ошибки: " << errorCode;
        return stream.str();
    }

    std::wstring message(buffer, chars);
    LocalFree(buffer);

    while (!message.empty() && (message.back() == L'\r' || message.back() == L'\n' || message.back() == L' ')) {
        message.pop_back();
    }
    return message;
}

bool QueryRequestOptionString(HINTERNET request, DWORD option, std::wstring& value) {
    DWORD bytes = 0;
    if (!WinHttpQueryOption(request, option, WINHTTP_NO_OUTPUT_BUFFER, &bytes)) {
        if (GetLastError() != ERROR_INSUFFICIENT_BUFFER) {
            return false;
        }
    }

    if (bytes == 0) {
        return false;
    }

    std::wstring buffer(bytes / sizeof(wchar_t), L'\0');
    if (!WinHttpQueryOption(request, option, buffer.data(), &bytes)) {
        return false;
    }

    value.assign(buffer.c_str());
    return !value.empty();
}

bool QueryLocationHeader(HINTERNET request, std::wstring& location) {
    DWORD bytes = 0;
    if (!WinHttpQueryHeaders(
            request,
            WINHTTP_QUERY_LOCATION,
            WINHTTP_HEADER_NAME_BY_INDEX,
            WINHTTP_NO_OUTPUT_BUFFER,
            &bytes,
            WINHTTP_NO_HEADER_INDEX)) {
        if (GetLastError() != ERROR_INSUFFICIENT_BUFFER) {
            return false;
        }
    }

    if (bytes == 0) {
        return false;
    }

    std::wstring buffer(bytes / sizeof(wchar_t), L'\0');
    if (!WinHttpQueryHeaders(
            request,
            WINHTTP_QUERY_LOCATION,
            WINHTTP_HEADER_NAME_BY_INDEX,
            buffer.data(),
            &bytes,
            WINHTTP_NO_HEADER_INDEX)) {
        return false;
    }

    location.assign(buffer.c_str());
    return !location.empty();
}

std::wstring ExtractTagFromUrl(const std::wstring& url) {
    const std::wstring token = L"/releases/tag/";
    const size_t tagStart = url.find(token);
    if (tagStart == std::wstring::npos) {
        return L"";
    }

    size_t valueStart = tagStart + token.size();
    size_t valueEnd = url.find_first_of(L"?#", valueStart);
    if (valueEnd == std::wstring::npos) {
        valueEnd = url.size();
    }
    if (valueEnd <= valueStart) {
        return L"";
    }

    std::wstring tag = url.substr(valueStart, valueEnd - valueStart);
    while (!tag.empty() && tag.back() == L'/') {
        tag.pop_back();
    }
    return tag;
}

std::vector<int> ParseVersionParts(const std::wstring& version) {
    std::vector<int> parts;
    int current = -1;

    for (wchar_t ch : version) {
        if (iswdigit(ch)) {
            if (current < 0) {
                current = 0;
            }
            current = (current * 10) + static_cast<int>(ch - L'0');
            continue;
        }

        if (current >= 0) {
            parts.push_back(current);
            current = -1;
        }

        if (ch != L'.') {
            break;
        }
    }

    if (current >= 0) {
        parts.push_back(current);
    }

    return parts;
}

std::wstring EscapePowerShellSingleQuoted(const std::wstring& value) {
    std::wstring escaped;
    escaped.reserve(value.size());

    for (wchar_t ch : value) {
        if (ch == L'\'') {
            escaped += L"''";
        } else {
            escaped.push_back(ch);
        }
    }

    return escaped;
}
} // namespace

UpdateCheckResult UpdateService::CheckForUpdates(const std::wstring& currentVersion) const {
    UpdateCheckResult result = {};

    std::wstring latestTag;
    if (!ResolveLatestReleaseTag(latestTag, result.errorMessage)) {
        result.success = false;
        return result;
    }

    const std::wstring latestVersion = NormalizeVersionFromTag(latestTag);
    const int compareResult = CompareVersions(currentVersion, latestVersion);

    result.success = true;
    result.updateAvailable = (compareResult < 0);
    result.latestTag = latestTag;
    result.latestVersion = latestVersion;
    return result;
}

bool UpdateService::DownloadReleaseExecutable(const std::wstring& tag, const std::wstring& destinationPath, std::wstring& errorMessage) const {
    if (tag.empty() || destinationPath.empty()) {
        errorMessage = L"Неверные параметры загрузки обновления";
        return false;
    }

    const std::wstring requestPath = std::wstring(kReleaseDownloadPrefix) + tag + L"/" + kReleaseExeName;

    WinHttpHandle session(WinHttpOpen(
        kUserAgent,
        WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
        WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS,
        0
    ));
    if (!session) {
        errorMessage = L"Не удалось инициализировать WinHTTP: " + FormatWin32Error(GetLastError());
        return false;
    }

    WinHttpHandle connection(WinHttpConnect(session.get(), kGitHubHost, INTERNET_DEFAULT_HTTPS_PORT, 0));
    if (!connection) {
        errorMessage = L"Не удалось подключиться к GitHub: " + FormatWin32Error(GetLastError());
        return false;
    }

    WinHttpHandle request(WinHttpOpenRequest(
        connection.get(),
        L"GET",
        requestPath.c_str(),
        nullptr,
        WINHTTP_NO_REFERER,
        WINHTTP_DEFAULT_ACCEPT_TYPES,
        WINHTTP_FLAG_SECURE
    ));
    if (!request) {
        errorMessage = L"Не удалось создать HTTP-запрос: " + FormatWin32Error(GetLastError());
        return false;
    }

    if (!WinHttpSendRequest(
            request.get(),
            WINHTTP_NO_ADDITIONAL_HEADERS,
            0,
            WINHTTP_NO_REQUEST_DATA,
            0,
            0,
            0)) {
        errorMessage = L"Не удалось отправить запрос на загрузку: " + FormatWin32Error(GetLastError());
        return false;
    }

    if (!WinHttpReceiveResponse(request.get(), nullptr)) {
        errorMessage = L"Не удалось получить ответ сервера: " + FormatWin32Error(GetLastError());
        return false;
    }

    DWORD statusCode = 0;
    DWORD statusSize = sizeof(statusCode);
    if (!WinHttpQueryHeaders(
            request.get(),
            WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
            WINHTTP_HEADER_NAME_BY_INDEX,
            &statusCode,
            &statusSize,
            WINHTTP_NO_HEADER_INDEX)) {
        errorMessage = L"Не удалось получить HTTP-статус: " + FormatWin32Error(GetLastError());
        return false;
    }

    if (statusCode != 200) {
        std::wstringstream stream;
        stream << L"Сервер вернул HTTP " << statusCode << L" при загрузке обновления";
        errorMessage = stream.str();
        return false;
    }

    HANDLE fileHandle = CreateFileW(
        destinationPath.c_str(),
        GENERIC_WRITE,
        0,
        nullptr,
        CREATE_ALWAYS,
        FILE_ATTRIBUTE_NORMAL,
        nullptr
    );
    if (fileHandle == INVALID_HANDLE_VALUE) {
        errorMessage = L"Не удалось создать файл обновления: " + FormatWin32Error(GetLastError());
        return false;
    }

    bool readOk = true;
    while (true) {
        DWORD bytesAvailable = 0;
        if (!WinHttpQueryDataAvailable(request.get(), &bytesAvailable)) {
            errorMessage = L"Ошибка получения данных обновления: " + FormatWin32Error(GetLastError());
            readOk = false;
            break;
        }

        if (bytesAvailable == 0) {
            break;
        }

        std::vector<BYTE> buffer(bytesAvailable);
        DWORD bytesRead = 0;
        if (!WinHttpReadData(request.get(), buffer.data(), bytesAvailable, &bytesRead)) {
            errorMessage = L"Ошибка чтения данных обновления: " + FormatWin32Error(GetLastError());
            readOk = false;
            break;
        }

        DWORD bytesWritten = 0;
        if (!WriteFile(fileHandle, buffer.data(), bytesRead, &bytesWritten, nullptr) || bytesWritten != bytesRead) {
            errorMessage = L"Ошибка записи файла обновления: " + FormatWin32Error(GetLastError());
            readOk = false;
            break;
        }
    }

    CloseHandle(fileHandle);

    if (!readOk) {
        DeleteFileW(destinationPath.c_str());
        return false;
    }

    return true;
}

bool UpdateService::LaunchUpdaterProcess(DWORD currentProcessId,
                                         const std::wstring& downloadedExePath,
                                         const std::wstring& targetExePath,
                                         std::wstring& errorMessage) const {
    if (downloadedExePath.empty() || targetExePath.empty()) {
        errorMessage = L"Неверные параметры запуска установщика обновления";
        return false;
    }

    std::wstringstream script;
    script << L"$pidToWait=" << currentProcessId << L";";
    script << L"$download='" << EscapePowerShellSingleQuoted(downloadedExePath) << L"';";
    script << L"$target='" << EscapePowerShellSingleQuoted(targetExePath) << L"';";
    script << L"while (Get-Process -Id $pidToWait -ErrorAction SilentlyContinue) { Start-Sleep -Milliseconds 500 };";
    script << L"Copy-Item -LiteralPath $download -Destination $target -Force;";
    script << L"Start-Process -FilePath $target;";
    script << L"Remove-Item -LiteralPath $download -Force -ErrorAction SilentlyContinue;";

    std::wstring commandLine =
        L"powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command \"" +
        script.str() + L"\"";

    STARTUPINFOW startupInfo = {};
    startupInfo.cb = sizeof(startupInfo);
    PROCESS_INFORMATION processInfo = {};
    std::vector<wchar_t> mutableCommand(commandLine.begin(), commandLine.end());
    mutableCommand.push_back(L'\0');

    if (!CreateProcessW(
            nullptr,
            mutableCommand.data(),
            nullptr,
            nullptr,
            FALSE,
            CREATE_NO_WINDOW,
            nullptr,
            nullptr,
            &startupInfo,
            &processInfo)) {
        errorMessage = L"Не удалось запустить процесс установки обновления: " + FormatWin32Error(GetLastError());
        return false;
    }

    CloseHandle(processInfo.hThread);
    CloseHandle(processInfo.hProcess);
    return true;
}

bool UpdateService::ResolveLatestReleaseTag(std::wstring& latestTag, std::wstring& errorMessage) const {
    latestTag.clear();

    WinHttpHandle session(WinHttpOpen(
        kUserAgent,
        WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
        WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS,
        0
    ));
    if (!session) {
        errorMessage = L"Не удалось инициализировать WinHTTP: " + FormatWin32Error(GetLastError());
        return false;
    }

    WinHttpHandle connection(WinHttpConnect(session.get(), kGitHubHost, INTERNET_DEFAULT_HTTPS_PORT, 0));
    if (!connection) {
        errorMessage = L"Не удалось подключиться к GitHub: " + FormatWin32Error(GetLastError());
        return false;
    }

    WinHttpHandle request(WinHttpOpenRequest(
        connection.get(),
        L"GET",
        kLatestReleasePath,
        nullptr,
        WINHTTP_NO_REFERER,
        WINHTTP_DEFAULT_ACCEPT_TYPES,
        WINHTTP_FLAG_SECURE
    ));
    if (!request) {
        errorMessage = L"Не удалось создать HTTP-запрос: " + FormatWin32Error(GetLastError());
        return false;
    }

    if (!WinHttpSendRequest(
            request.get(),
            WINHTTP_NO_ADDITIONAL_HEADERS,
            0,
            WINHTTP_NO_REQUEST_DATA,
            0,
            0,
            0)) {
        errorMessage = L"Не удалось отправить запрос проверки обновлений: " + FormatWin32Error(GetLastError());
        return false;
    }

    if (!WinHttpReceiveResponse(request.get(), nullptr)) {
        errorMessage = L"Не удалось получить ответ сервера: " + FormatWin32Error(GetLastError());
        return false;
    }

    DWORD statusCode = 0;
    DWORD statusSize = sizeof(statusCode);
    if (!WinHttpQueryHeaders(
            request.get(),
            WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
            WINHTTP_HEADER_NAME_BY_INDEX,
            &statusCode,
            &statusSize,
            WINHTTP_NO_HEADER_INDEX)) {
        errorMessage = L"Не удалось получить HTTP-статус: " + FormatWin32Error(GetLastError());
        return false;
    }

    std::wstring sourceUrl;
    if (!QueryRequestOptionString(request.get(), WINHTTP_OPTION_URL, sourceUrl)) {
        QueryLocationHeader(request.get(), sourceUrl);
    }

    if (sourceUrl.empty()) {
        std::wstringstream stream;
        stream << L"Не удалось определить URL последнего релиза (HTTP " << statusCode << L")";
        errorMessage = stream.str();
        return false;
    }

    latestTag = ExtractTagFromUrl(sourceUrl);
    if (latestTag.empty()) {
        std::wstringstream stream;
        stream << L"Не удалось извлечь тег релиза из URL: " << sourceUrl;
        errorMessage = stream.str();
        return false;
    }

    return true;
}

std::wstring UpdateService::NormalizeVersionFromTag(const std::wstring& rawTag) {
    if (!rawTag.empty() && (rawTag.front() == L'v' || rawTag.front() == L'V')) {
        return rawTag.substr(1);
    }
    return rawTag;
}

int UpdateService::CompareVersions(const std::wstring& left, const std::wstring& right) {
    const std::vector<int> leftParts = ParseVersionParts(left);
    const std::vector<int> rightParts = ParseVersionParts(right);
    const size_t maxParts = (std::max)(leftParts.size(), rightParts.size());

    for (size_t index = 0; index < maxParts; ++index) {
        const int leftValue = (index < leftParts.size()) ? leftParts[index] : 0;
        const int rightValue = (index < rightParts.size()) ? rightParts[index] : 0;
        if (leftValue < rightValue) {
            return -1;
        }
        if (leftValue > rightValue) {
            return 1;
        }
    }

    return 0;
}
