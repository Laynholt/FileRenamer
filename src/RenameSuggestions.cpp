#include "RenameSuggestions.h"
#include <algorithm>
#include <regex>
#include <set>

namespace RenamerCore {
namespace {
auto RegexFlags(bool ignoreCase) {
    return std::regex_constants::ECMAScript |
        (ignoreCase ? std::regex_constants::icase : std::regex_constants::syntax_option_type{});
}

std::wstring EscapeRegex(const std::wstring& text) {
    std::wstring escaped;
    for (wchar_t ch : text) {
        if (std::wstring(L"\\^$.|?*+()[]{}").find(ch) != std::wstring::npos) escaped += L'\\';
        escaped += ch;
    }
    return escaped;
}

bool Measure(RenameSuggestion& suggestion, const std::vector<RenameOperation>& entries, bool ignoreCase) {
    const std::wregex expression(suggestion.pattern, RegexFlags(ignoreCase));
    suggestion.totalCount = entries.size();
    for (const auto& entry : entries) {
        if (std::regex_search(entry.oldName, expression)) {
            ++suggestion.matchCount;
            if (suggestion.exampleBefore.empty()) {
                suggestion.exampleBefore = entry.oldName;
                suggestion.exampleAfter = std::regex_replace(entry.oldName, expression, suggestion.replacement);
            }
        } else if (suggestion.unmatchedExample.empty()) {
            suggestion.unmatchedExample = entry.oldName;
        }
    }
    return suggestion.matchCount != 0;
}
} // namespace

std::vector<RenameSuggestion> SuggestPatterns(const std::vector<RenameOperation>& entries, bool ignoreCase) {
    std::vector<RenameSuggestion> result;
    if (entries.empty()) return result;
    auto add = [&](std::wstring title, std::wstring pattern, std::wstring replacement,
                   std::wstring description, std::vector<std::wstring> groups) {
        RenameSuggestion s;
        s.title = std::move(title);
        s.pattern = std::move(pattern);
        s.replacement = std::move(replacement);
        s.description = std::move(description);
        s.groups = std::move(groups);
        if (Measure(s, entries, ignoreCase)) result.push_back(std::move(s));
    };

    const std::vector<std::wstring> episodeGroups = {L"Название", L"Номер сезона", L"Номер серии", L"Расширение"};
    add(L"Сезон и серия: S02E03", LR"(^(.*?)[ ._-]*[sS]([0-9]{1,3})[ ._-]*[eE]([0-9]{1,4})(?:[ ._-].*)?(\.[^.]+)$)",
        L"Название — сезон $2, серия $3$4",
        L"$2 — сезон, $3 — серия, $4 — расширение. Убирает качество и прочий хвост. Замените слово «Название» своим.", episodeGroups);
    add(L"Сезон и серия: 2x03", LR"(^(.*?)[ ._-]*([0-9]{1,3})[xX]([0-9]{1,4})(?:[ ._-].*)?(\.[^.]+)$)",
        L"Название — сезон $2, серия $3$4",
        L"$2 — сезон, $3 — серия, $4 — расширение. Убирает хвост после номера серии. Замените слово «Название» своим.", episodeGroups);
    add(L"Номер в конце имени", LR"(^(.*?)[ ._-]+([0-9]+)(\.[^.]+)$)", L"Название $2$3",
        L"$2 — существующий номер, $3 — расширение. Сохраняет номер, заменяет название.",
        {L"Название", L"Номер", L"Расширение"});

    // Infer a few shared structures by keeping literal text and capturing numeric
    // runs. Bound candidate generation, but measure coverage against every entry.
    std::set<std::wstring> seen;
    std::vector<RenameSuggestion> inferred;
    for (size_t index = 0; index < entries.size() && index < 128 && seen.size() < 16; ++index) {
        const auto& entry = entries[index];
        if (entry.oldName.size() > 260) continue;
        const std::filesystem::path path(entry.oldName);
        const std::wstring stem = entry.isDirectory ? entry.oldName : path.stem().wstring();
        const std::wstring extension = entry.isDirectory ? L"" : path.extension().wstring();
        RenameSuggestion s;
        s.title = L"Общий шаблон: " + stem.substr(0, 45);
        s.pattern = L"^";
        s.replacement = L"Название";
        for (size_t i = 0; i < stem.size();) {
            if (stem[i] >= L'0' && stem[i] <= L'9') {
                do { ++i; } while (i < stem.size() && stem[i] >= L'0' && stem[i] <= L'9');
                s.pattern += L"([0-9]+)";
                s.groups.push_back(L"Число " + std::to_wstring(s.groups.size() + 1));
                s.replacement += L" $" + std::to_wstring(s.groups.size());
            } else if (std::wstring(L" ._-").find(stem[i]) != std::wstring::npos) {
                do { ++i; } while (i < stem.size() && std::wstring(L" ._-").find(stem[i]) != std::wstring::npos);
                s.pattern += L"[ ._-]+";
            } else {
                s.pattern += EscapeRegex(stem.substr(i++, 1));
            }
        }
        if (s.groups.empty() || s.groups.size() > 4) continue;
        if (!extension.empty()) {
            s.pattern += L"(" + EscapeRegex(extension) + L")";
            s.groups.push_back(L"Расширение");
            s.replacement += L"$" + std::to_wstring(s.groups.size());
        }
        s.pattern += L"$";
        if (!seen.insert(s.pattern).second) continue;
        s.description = L"Общий текст остаётся в паттерне, числа становятся группами $1, $2… Разделители могут отличаться. В замене задайте своё название; числа и расширение сохраняются.";
        if (Measure(s, entries, ignoreCase) && s.matchCount >= 2) inferred.push_back(std::move(s));
    }
    std::stable_sort(inferred.begin(), inferred.end(), [](const auto& a, const auto& b) { return a.matchCount > b.matchCount; });
    for (size_t i = 0; i < inferred.size() && i < 4; ++i) result.push_back(std::move(inferred[i]));

    add(L"Убрать текст в квадратных скобках", LR"(\[[^\]\r\n]*\][ ._-]*)", L"",
        L"Удаляет квадратные скобки вместе с содержимым и разделителями сразу после них. Шаблон замены пустой.", {});
    add(L"Убрать номер в начале", LR"(^[0-9]+[ ._-]+)", L"",
        L"Удаляет начальный номер и следующий разделитель. Остальная часть имени остаётся прежней.", {});
    add(L"Заменить разделители пробелом", LR"([ _-]+)", L" ",
        L"Заменяет подчёркивания, дефисы и повторяющиеся пробелы одним пробелом. Точки сохраняются.", {});
    add(L"Выделить числа", LR"(([0-9]+))", L"$1",
        L"Каждая последовательность цифр — группа $1. По умолчанию оставляет числа; измените замену, чтобы их заменить или удалить.", {L"Число"});
    return result;
}

std::vector<RenameSuggestion> SuggestReplacements(const std::wstring& pattern, bool useRegex, bool ignoreCase,
    const std::vector<RenameOperation>& entries, const std::vector<std::wstring>& groupLabels) {
    std::vector<RenameSuggestion> result;
    auto token = [&](const std::wstring& title, const std::wstring& value, const std::wstring& description) {
        RenameSuggestion s;
        s.title = title + L" — " + value;
        s.replacement = value;
        s.description = description;
        result.push_back(std::move(s));
    };
    token(L"Порядковый номер", L"{n}", L"1, 2, 3… в порядке списка, только среди подходящих элементов. Один номер на файл или папку. С пустым паттерном заменяет имя, сохраняя расширение файла.");
    token(L"Номер из двух цифр", L"{n:02}", L"01, 02, 03… Начинается с 1; после 99 продолжает 100, 101…");
    token(L"Номер из трёх цифр", L"{n:03}", L"001, 002, 003… Например: Название {n:03}. С пустым паттерном расширение сохраняется автоматически.");
    if (!useRegex || pattern.empty()) return result;
    try {
        const std::wregex expression(pattern, RegexFlags(ignoreCase));
        std::wsmatch match;
        for (const auto& entry : entries) {
            if (std::regex_search(entry.oldName, match, expression)) break;
        }
        for (size_t i = 1; i <= expression.mark_count() && i <= 99; ++i) {
            const std::wstring label = i <= groupLabels.size() ? groupLabels[i - 1] : L"Группа " + std::to_wstring(i);
            token(label, L"$" + std::to_wstring(i), L"Вставляет текст из группы " + std::to_wstring(i) +
                L" текущего regex. Это часть исходного имени, а не счётчик.");
            if (!match.empty() && i < match.size()) {
                result.back().exampleBefore = match.str(0);
                result.back().exampleAfter = match.str(i);
            }
        }
        token(L"Всё совпадение", L"$&", L"Вставляет весь фрагмент имени, найденный регулярным выражением.");
        token(L"Символ доллара", L"$$", L"В режиме regex вставляет один буквальный знак $.");
    } catch (const std::regex_error&) {
        // Counter suggestions remain useful while the user is typing a regex.
    }
    return result;
}
}
