#pragma once

#include <cstddef>
#include <filesystem>
#include <string>
#include <vector>

namespace RenamerCore {

struct RenameOperation {
    std::filesystem::path oldPath;
    std::filesystem::path newPath;
    std::wstring oldName;
    std::wstring newName;
    bool isDirectory;
};

struct CollectResult {
    std::vector<RenameOperation> operations;
    std::wstring status;
    std::size_t totalCount;
};

enum class ExecuteStatus {
    Success,
    NoChanges,
    Error
};

struct ExecuteResult {
    ExecuteStatus status;
    std::wstring message;
    std::size_t renamedCount;
};

CollectResult CollectOperations(
    const std::wstring& folderText,
    const std::wstring& pattern,
    const std::wstring& replacement,
    bool useRegex,
    bool ignoreCase,
    std::size_t maxOperations = 0
);

ExecuteResult ExecuteRename(const std::vector<RenameOperation>& operations);

} // namespace RenamerCore
