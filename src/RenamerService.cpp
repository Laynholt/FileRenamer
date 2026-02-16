#include "RenamerService.h"

#include <windows.h>
#include <objbase.h>

#include <algorithm>
#include <cctype>
#include <cwctype>
#include <optional>
#include <regex>
#include <set>

namespace fs = std::filesystem;

namespace {
struct EntryInfo {
    std::wstring name;
    bool isDirectory;
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

std::wstring ReplaceAll(const std::wstring& text, const std::wstring& pattern, const std::wstring& replacement) {
    if (pattern.empty()) {
        return text;
    }

    std::wstring result = text;
    size_t position = 0;

    while ((position = result.find(pattern, position)) != std::wstring::npos) {
        result.replace(position, pattern.size(), replacement);
        position += replacement.size();
    }

    return result;
}

std::wstring ToLowerCopy(const std::wstring& text) {
    std::wstring lowered = text;
    std::transform(lowered.begin(), lowered.end(), lowered.begin(), [](wchar_t ch) {
        return static_cast<wchar_t>(std::towlower(ch));
    });
    return lowered;
}

size_t FindCaseInsensitive(const std::wstring& text, const std::wstring& pattern, size_t start = 0) {
    if (pattern.empty() || start >= text.size()) {
        return std::wstring::npos;
    }

    const std::wstring lowerText = ToLowerCopy(text);
    const std::wstring lowerPattern = ToLowerCopy(pattern);
    return lowerText.find(lowerPattern, start);
}

std::wstring ReplaceAllCaseInsensitive(const std::wstring& text, const std::wstring& pattern, const std::wstring& replacement) {
    if (pattern.empty()) {
        return text;
    }

    const std::wstring lowerText = ToLowerCopy(text);
    const std::wstring lowerPattern = ToLowerCopy(pattern);

    std::wstring result;
    result.reserve(text.size());

    size_t sourcePos = 0;
    while (sourcePos < text.size()) {
        const size_t foundPos = lowerText.find(lowerPattern, sourcePos);
        if (foundPos == std::wstring::npos) {
            result.append(text.substr(sourcePos));
            break;
        }

        result.append(text.substr(sourcePos, foundPos - sourcePos));
        result.append(replacement);
        sourcePos = foundPos + pattern.size();
    }

    return result;
}

std::wstring PathKey(const fs::path& path) {
    std::error_code absoluteEc;
    const fs::path absolutePath = fs::absolute(path, absoluteEc);
    std::wstring key = absoluteEc ? path.wstring() : absolutePath.lexically_normal().wstring();
    std::transform(key.begin(), key.end(), key.begin(), [](wchar_t ch) {
        return static_cast<wchar_t>(std::towlower(ch));
    });
    return key;
}

std::wstring MakeTempSuffix() {
    GUID guid = {};
    if (FAILED(CoCreateGuid(&guid))) {
        return L".renamer_tmp_fallback_" + std::to_wstring(GetTickCount64());
    }

    wchar_t guidBuffer[64] = {};
    StringFromGUID2(guid, guidBuffer, 64);

    std::wstring token = guidBuffer;
    token.erase(
        std::remove_if(token.begin(), token.end(), [](wchar_t ch) {
            return ch == L'{' || ch == L'}' || ch == L'-';
        }),
        token.end()
    );

    return L".renamer_tmp_" + token;
}

} // namespace

namespace RenamerCore {

CollectResult CollectOperations(
    const std::wstring& folderText,
    const std::wstring& pattern,
    const std::wstring& replacement,
    bool useRegex,
    bool ignoreCase,
    std::size_t maxOperations
) {
    CollectResult result;
    result.totalCount = 0;

    const std::wstring folder = Trim(folderText);
    const bool hasPattern = !pattern.empty();

    if (folder.empty()) {
        result.status = L"Укажите папку.";
        return result;
    }

    fs::path folderPath(folder);
    std::error_code ec;
    if (!fs::exists(folderPath, ec) || !fs::is_directory(folderPath, ec)) {
        result.status = L"Папка не найдена.";
        return result;
    }

    std::optional<std::wregex> regexPattern;
    if (hasPattern && useRegex) {
        try {
            auto flags = std::regex_constants::ECMAScript;
            if (ignoreCase) {
                flags |= std::regex_constants::icase;
            }
            regexPattern.emplace(pattern, flags);
        } catch (const std::regex_error&) {
            result.status = L"Ошибка regex: некорректный шаблон.";
            return result;
        }
    }

    std::vector<EntryInfo> entries;
    for (fs::directory_iterator it(folderPath, ec), end; it != end; it.increment(ec)) {
        if (ec) {
            result.status = L"Не удалось прочитать содержимое папки.";
            return result;
        }

        std::error_code typeEc;
        const bool isDirectory = it->is_directory(typeEc);
        if (typeEc) {
            continue;
        }

        const bool isRegularFile = it->is_regular_file(typeEc);
        if (typeEc) {
            continue;
        }

        if (!isRegularFile && !isDirectory) {
            continue;
        }

        entries.push_back({ it->path().filename().wstring(), isDirectory });
    }

    std::sort(entries.begin(), entries.end(), [](const EntryInfo& left, const EntryInfo& right) {
        return left.name < right.name;
    });

    auto addOperation = [&](const RenameOperation& operation) {
        ++result.totalCount;
        if (maxOperations == 0 || result.operations.size() < maxOperations) {
            result.operations.push_back(operation);
        }
    };

    if (hasPattern) {
        for (const EntryInfo& entry : entries) {
            const std::wstring& name = entry.name;
            std::wstring newName;

            if (useRegex) {
                if (!std::regex_search(name, *regexPattern)) {
                    continue;
                }
                newName = std::regex_replace(name, *regexPattern, replacement);
            } else {
                if (ignoreCase) {
                    if (FindCaseInsensitive(name, pattern) == std::wstring::npos) {
                        continue;
                    }
                    newName = ReplaceAllCaseInsensitive(name, pattern, replacement);
                } else {
                    if (name.find(pattern) == std::wstring::npos) {
                        continue;
                    }
                    newName = ReplaceAll(name, pattern, replacement);
                }
            }

            addOperation({
                folderPath / name,
                folderPath / newName,
                name,
                newName,
                entry.isDirectory
            });
        }

        result.status = L"Найдено совпадений: " + std::to_wstring(result.totalCount);
        return result;
    }

    const bool isPrefixMode = !replacement.empty() && replacement.front() == L'<';
    const bool isSuffixMode = !replacement.empty() && replacement.front() == L'>';

    if (isPrefixMode || isSuffixMode) {
        for (const EntryInfo& entry : entries) {
            const std::wstring& name = entry.name;
            std::wstring newName;
            if (isPrefixMode) {
                newName = replacement.substr(1) + name;
            } else {
                if (entry.isDirectory) {
                    newName = name + replacement.substr(1);
                } else {
                    const fs::path filePath(name);
                    const std::wstring stem = filePath.stem().wstring();
                    const std::wstring ext = filePath.extension().wstring();
                    newName = stem + replacement.substr(1) + ext;
                }
            }

            addOperation({
                folderPath / name,
                folderPath / newName,
                name,
                newName,
                entry.isDirectory
            });
        }

        result.status = L"Паттерн пустой: массовый режим, элементов: " + std::to_wstring(result.totalCount);
        return result;
    }

    for (const EntryInfo& entry : entries) {
        const std::wstring& name = entry.name;
        addOperation({
            folderPath / name,
            folderPath / name,
            name,
            name,
            entry.isDirectory
        });
    }

    result.status = L"Паттерн пустой: показаны все элементы (" + std::to_wstring(result.totalCount) + L")";
    return result;
}

ExecuteResult ExecuteRename(const std::vector<RenameOperation>& operations) {
    std::vector<RenameOperation> toRename;
    toRename.reserve(operations.size());
    for (const RenameOperation& operation : operations) {
        if (operation.oldPath != operation.newPath) {
            toRename.push_back(operation);
        }
    }

    if (toRename.empty()) {
        return { ExecuteStatus::NoChanges, L"Изменений нет: имена уже соответствуют шаблону.", 0 };
    }

    std::set<std::wstring> uniqueNewPaths;
    for (const RenameOperation& operation : toRename) {
        const std::wstring key = PathKey(operation.newPath);
        if (!uniqueNewPaths.insert(key).second) {
            return { ExecuteStatus::Error, L"После замены есть дублирующиеся имена.", 0 };
        }
    }

    std::set<std::wstring> oldPathKeys;
    for (const RenameOperation& operation : toRename) {
        oldPathKeys.insert(PathKey(operation.oldPath));
    }

    std::vector<fs::path> conflicts;
    for (const RenameOperation& operation : toRename) {
        std::error_code existsEc;
        if (fs::exists(operation.newPath, existsEc) && oldPathKeys.find(PathKey(operation.newPath)) == oldPathKeys.end()) {
            conflicts.push_back(operation.newPath);
        }
    }

    if (!conflicts.empty()) {
        std::wstring message = L"Эти элементы уже существуют:\n";
        const size_t shown = (std::min)(conflicts.size(), static_cast<size_t>(10));
        for (size_t index = 0; index < shown; ++index) {
            message += conflicts[index].filename().wstring();
            if (index + 1 < shown) {
                message += L"\n";
            }
        }

        return { ExecuteStatus::Error, message, 0 };
    }

    struct TempMapping {
        fs::path tempPath;
        fs::path oldPath;
        fs::path targetPath;
    };

    std::vector<TempMapping> tempMapping;
    tempMapping.reserve(toRename.size());

    std::wstring errorMessage;
    bool failed = false;

    for (const RenameOperation& operation : toRename) {
        const fs::path tempPath = fs::path(operation.oldPath.wstring() + MakeTempSuffix());

        std::error_code renameEc;
        fs::rename(operation.oldPath, tempPath, renameEc);
        if (renameEc) {
            failed = true;
            errorMessage = L"Не удалось переименовать временный файл: " + operation.oldName;
            break;
        }

        tempMapping.push_back({ tempPath, operation.oldPath, operation.newPath });
    }

    if (!failed) {
        for (const TempMapping& mapping : tempMapping) {
            std::error_code renameEc;
            fs::rename(mapping.tempPath, mapping.targetPath, renameEc);
            if (renameEc) {
                failed = true;
                errorMessage = L"Не удалось завершить переименование: " + mapping.targetPath.filename().wstring();
                break;
            }
        }
    }

    if (failed) {
        for (const TempMapping& mapping : tempMapping) {
            std::error_code existsEc;
            if (fs::exists(mapping.tempPath, existsEc)) {
                std::error_code rollbackEc;
                fs::rename(mapping.tempPath, mapping.oldPath, rollbackEc);
            }
        }

        return { ExecuteStatus::Error, errorMessage, 0 };
    }

    return { ExecuteStatus::Success, L"", toRename.size() };
}

} // namespace RenamerCore
