#pragma once
#include "RenamerService.h"

namespace RenamerCore {
struct RenameSuggestion {
    std::wstring title;
    std::wstring pattern;
    std::wstring replacement;
    std::wstring description;
    std::vector<std::wstring> groups;
    size_t matchCount = 0;
    size_t totalCount = 0;
    std::wstring exampleBefore;
    std::wstring exampleAfter;
    std::wstring unmatchedExample;
};

std::vector<RenameSuggestion> SuggestPatterns(const std::vector<RenameOperation>& entries, bool ignoreCase);
std::vector<RenameSuggestion> SuggestReplacements(const std::wstring& pattern, bool useRegex, bool ignoreCase,
    const std::vector<RenameOperation>& entries, const std::vector<std::wstring>& groupLabels = {});
} // namespace RenamerCore
