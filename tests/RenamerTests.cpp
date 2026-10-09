#include "RenamerService.h"
#include "RenameSuggestions.h"
#include <windows.h>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <stdexcept>
#include <algorithm>

namespace fs = std::filesystem;
using namespace RenamerCore;

void Require(bool condition, const char* message) {
    if (!condition) throw std::runtime_error(message);
}

struct Fixture {
    fs::path path = fs::temp_directory_path() / (L"FileRenamer-tests-" +
        std::to_wstring(GetCurrentProcessId()) + L"-" + std::to_wstring(GetTickCount64()));
    Fixture() { Require(fs::create_directory(path), "create isolated test directory"); }
    ~Fixture() { std::error_code ec; fs::remove_all(path, ec); }
    void File(const wchar_t* name) { std::ofstream(path / name) << "test"; }
    CollectResult Collect(const wchar_t* pattern, const wchar_t* replacement, bool regex = false,
                          size_t limit = 0) {
        return CollectOperations(path.wstring(), pattern, replacement, regex, false, limit);
    }
};

void CounterTests() {
    Fixture f;
    f.File(L"episode10.mkv");
    f.File(L"episode2.mkv");
    f.File(L"notes.txt");
    auto all = f.Collect(L"", L"Название {n:03}");
    Require(all.operations.size() == 3, "counter includes all items with empty pattern");
    Require(all.operations[0].oldName == L"episode2.mkv", "natural ordering");
    Require(all.operations[0].newName == L"Название 001.mkv", "counter replaces stem and preserves extension");
    Require(all.operations[1].newName == L"Название 002.mkv", "counter increments once per item");
    Require(all.operations[2].newName == L"Название 003.txt", "counter preserves different extensions");
    auto preview = f.Collect(L"", L"Название {n:03}", false, 1);
    Require(preview.totalCount == 3 && preview.operations[0].newName == all.operations[0].newName,
            "preview limit must not change numbering");
    auto regex = f.Collect(L"episode([0-9]+)", L"Серия {n} (исходная $1)", true);
    Require(regex.operations.size() == 2, "only matching items receive numbers");
    Require(regex.operations[1].newName == L"Серия 2 (исходная 10).mkv", "counter and capture groups coexist");
    auto prefix = f.Collect(L"", L"<{n}-");
    Require(prefix.operations[1].newName == L"2-episode10.mkv", "counter in prefix mode");
    auto suffix = f.Collect(L"", L">-{n:02}");
    Require(suffix.operations[0].newName == L"episode2-01.mkv", "counter in suffix mode");
    auto literal = f.Collect(L"episode", L"{n}-{n}");
    Require(literal.operations[1].newName == L"2-210.mkv", "repeated tokens use same item number");
    auto invalid = f.Collect(L"episode", L"{n:999999999999}");
    Require(invalid.operations[0].newName == L"{n:999999999999}2.mkv", "malformed counter stays literal");
    auto outcome = ExecuteRename(regex.operations);
    Require(outcome.status == ExecuteStatus::Success && fs::exists(f.path / L"Серия 2 (исходная 10).mkv"),
            "execute produces the names shown in preview");
}

void SuggestionTests() {
    Fixture f;
    f.File(L"Show.S02E03.1080p.mkv");
    f.File(L"Show S02E04 720p.mkv");
    f.File(L"Show_s02e05.mkv");
    f.File(L"notes.txt");
    const auto entries = f.Collect(L"", L"").operations;
    const auto suggestions = SuggestPatterns(entries, false);
    Require(!suggestions.empty(), "suggest actual patterns from filenames");
    auto season = std::find_if(suggestions.begin(), suggestions.end(), [](const auto& s) {
        return s.groups.size() == 4 && s.groups[1] == L"Номер сезона";
    });
    Require(season != suggestions.end(), "recognize season and episode across separator variants");
    Require(season->matchCount == 3 && season->totalCount == 4, "report actual coverage");
    Require(season->unmatchedExample == L"notes.txt", "show an excluded filename");
    const auto applied = f.Collect(season->pattern.c_str(), season->replacement.c_str(), true);
    Require(applied.operations.size() == 3, "suggested coverage agrees with rename engine");
    for (const auto& op : applied.operations) {
        Require(op.newName.find(L"сезон 02, серия 0") != std::wstring::npos, "preserve season and episode values");
        Require(fs::path(op.newName).extension() == L".mkv", "suggestions preserve extension");
        Require(op.newName.find(L"1080p") == std::wstring::npos, "suggestions remove quality suffix");
    }
    auto tokens = SuggestReplacements(season->pattern, true, false, entries, season->groups);
    auto group = std::find_if(tokens.begin(), tokens.end(), [](const auto& s) { return s.replacement == L"$3"; });
    Require(group != tokens.end() && group->title.find(L"Номер серии") != std::wstring::npos,
            "replacement capture explains its meaning");
    Require(!group->exampleAfter.empty(), "capture suggestion includes real matched text");
    auto invalid = SuggestReplacements(L"[", true, false, entries);
    Require(invalid.size() >= 2 && std::none_of(invalid.begin(), invalid.end(), [](const auto& s) {
        return s.replacement.find(L'$') != std::wstring::npos;
    }), "invalid regex keeps counters available without inventing captures");
    auto plain = SuggestReplacements(L"", false, false, {});
    Require(std::any_of(plain.begin(), plain.end(), [](const auto& s) { return s.replacement == L"{n:03}"; }),
            "counter suggestions available without regex or folder");
}

void InferredPatternTests() {
    Fixture f;
    f.File(L"Мой [проект]+_12.txt");
    f.File(L"Мой [проект]+_13.txt");
    f.File(L"Иной.txt");
    auto suggestions = SuggestPatterns(f.Collect(L"", L"").operations, false);
    auto common = std::find_if(suggestions.begin(), suggestions.end(), [](const auto& s) {
        return s.title.find(L"Общий шаблон") == 0;
    });
    Require(common != suggestions.end() && common->matchCount == 2, "infer common structure and escape regex punctuation");
    auto applied = f.Collect(common->pattern.c_str(), common->replacement.c_str(), true);
    Require(applied.operations[0].newName == L"Название 12.txt", "inferred replacement preserves variable number and extension");
}

void CounterEdgeTests() {
    Fixture f;
    f.File(L"a{n}.txt");
    f.File(L"b.txt");
    f.File(L"c");
    fs::create_directory(f.path / L"folder.with.dots");
    auto captured = f.Collect(L"^(.+)\\.txt$", L"$1-{n}.txt", true);
    Require(captured.operations[0].newName == L"a{n}-1.txt", "captured data is not evaluated as a counter");
    auto adjacent = f.Collect(L"^(.+)\\.txt$", L"$1{n}.txt", true);
    Require(adjacent.operations[0].newName == L"a{n}1.txt", "adjacent capture and counter must not become group 11");
    auto dollar = f.Collect(L"^(.+)\\.txt$", L"${n}.txt", true);
    Require(dollar.operations[0].newName == L"$1.txt", "literal dollar before counter must not become a capture");
    auto plain = f.Collect(L"", L"Item {n}");
    Require(plain.operations[2].newName == L"Item 3", "extensionless file remains extensionless");
    Require(plain.operations[3].newName == L"Item 4", "directory dots are not a file extension");
    auto unchanged = f.Collect(L"", L"Just text");
    Require(unchanged.operations[0].newName == L"a{n}.txt", "empty-pattern behavior without counter unchanged");
    auto perItem = f.Collect(L"[at]", L"{n}", true);
    Require(perItem.operations[0].newName == L"1{n}.1x1", "all regex matches within one item share its number");
    auto collision = f.Collect(L"^.*\\.txt$", L"same.txt", true);
    Require(ExecuteRename(collision.operations).status == ExecuteStatus::Error && fs::exists(f.path / L"b.txt"),
            "duplicate generated names rejected without renaming sources");
}

void CleanupSuggestionTests() {
    Fixture f;
    f.File(L"[Group] Show 2x03 1080p.mkv");
    f.File(L"[Other] Show 2x04.mkv");
    const auto suggestions = SuggestPatterns(f.Collect(L"", L"").operations, true);
    for (const auto& s : suggestions) {
        const auto applied = f.Collect(s.pattern.c_str(), s.replacement.c_str(), true);
        Require(s.matchCount == applied.totalCount, "each suggested coverage agrees with execution");
        Require(s.exampleAfter == applied.operations.front().newName, "each suggestion example uses the same regex semantics");
    }
    const auto bracket = std::find_if(suggestions.begin(), suggestions.end(), [](const auto& s) {
        return s.title == L"Убрать текст в квадратных скобках";
    });
    Require(bracket != suggestions.end(), "bracket cleanup is offered");
    Require(bracket->exampleAfter.find(L'[') == std::wstring::npos, "bracket cleanup removes tag");
    const auto season = std::find_if(suggestions.begin(), suggestions.end(), [](const auto& s) {
        return s.title == L"Сезон и серия: 2x03";
    });
    Require(season != suggestions.end() && season->matchCount == 2, "recognize alternative season notation");
    const auto numbers = std::find_if(suggestions.begin(), suggestions.end(), [](const auto& s) {
        return s.title == L"Выделить числа";
    });
    Require(numbers != suggestions.end(), "numeric capture suggestion exists");
    const auto captures = SuggestReplacements(numbers->pattern, true, true, f.Collect(L"", L"").operations, numbers->groups);
    Require(std::none_of(captures.begin(), captures.end(), [](const auto& s) { return s.replacement == L"$2"; }),
            "numeric suggestion has exactly one captured group");
}

void RegexCounterBoundaryTests() {
    Fixture f;
    f.File(L"aba.txt");
    struct Case { const wchar_t* pattern; const wchar_t* replacement; const wchar_t* expected; };
    const Case cases[] = {
        {L"^", L"{n}", L"1aba.txt"},
        {L"$", L"{n}", L"aba.txt1"},
        {L".*", L"{n}", L"11"},
        {L"a", L"$&-{n}", L"a-1ba-1.txt"},
        {L"(?=a)", L"N{n}", L"N1abN1a.txt"},
        {L"a", L"$${n}", L"$1b$1.txt"},
        {L"^(.*)$", L"$1{n}{n:03}", L"aba.txt1001"},
    };
    for (const auto& c : cases) {
        auto result = f.Collect(c.pattern, c.replacement, true);
        Require(result.operations.size() == 1 && result.operations[0].newName == c.expected,
                "counter formatting preserves zero-width and repeated regex matches");
    }
}

int main() {
    try {
        CounterTests();
        SuggestionTests();
        InferredPatternTests();
        CounterEdgeTests();
        CleanupSuggestionTests();
        RegexCounterBoundaryTests();
        std::cout << "All renamer tests passed\n";
        return 0;
    } catch (const std::exception& error) {
        std::cerr << "FAIL: " << error.what() << '\n';
        return 1;
    }
}
