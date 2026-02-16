#pragma once

#include <windows.h>

#include <string>

struct UpdateCheckResult {
    bool success;
    bool updateAvailable;
    std::wstring latestTag;
    std::wstring latestVersion;
    std::wstring errorMessage;
};

class UpdateService {
public:
    UpdateCheckResult CheckForUpdates(const std::wstring& currentVersion) const;
    bool DownloadReleaseExecutable(const std::wstring& tag, const std::wstring& destinationPath, std::wstring& errorMessage) const;
    bool LaunchUpdaterProcess(DWORD currentProcessId,
                              const std::wstring& downloadedExePath,
                              const std::wstring& targetExePath,
                              std::wstring& errorMessage) const;

private:
    bool ResolveLatestReleaseTag(std::wstring& latestTag, std::wstring& errorMessage) const;

    static std::wstring NormalizeVersionFromTag(const std::wstring& rawTag);
    static int CompareVersions(const std::wstring& left, const std::wstring& right);
};
