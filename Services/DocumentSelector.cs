using AIDocAssistant.Models;

namespace AIDocAssistant.Services;

// Фаза 1'ден DataQuery алып, документтен точно керектүү абзацтарды тандайт
public static class DocumentSelector
{
    public static List<ParagraphInfo> Select(List<ParagraphInfo> all, DataQuery query)
    {
        return query.Mode switch
        {
            "search"  => SearchParagraphs(all, query.SearchText),
            "range"   => RangeParagraphs(all, query.FromIndex, query.ToIndex),
            "section" => SectionParagraphs(all, query.SectionName),
            _         => all // "all"
        };
    }

    // Поиск по тексту — берём совпадающие абзацы + ±3 соседних для контекста
    private static List<ParagraphInfo> SearchParagraphs(List<ParagraphInfo> all, string? searchText)
    {
        if (string.IsNullOrWhiteSpace(searchText)) return all;

        var matches = all
            .Where(p => p.Text.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            .Select(p => p.Index)
            .ToHashSet();

        if (matches.Count == 0) return all; // Ничего не нашли — вернём всё

        // Расширяем контекст: ±3 абзаца вокруг каждого совпадения
        var indices = new HashSet<int>();
        foreach (var i in matches)
            for (int k = Math.Max(0, i - 3); k <= Math.Min(all.Count - 1, i + 3); k++)
                indices.Add(k);

        return all.Where(p => indices.Contains(p.Index)).OrderBy(p => p.Index).ToList();
    }

    // Диапазон абзацев
    private static List<ParagraphInfo> RangeParagraphs(List<ParagraphInfo> all, int? from, int? to)
    {
        int start = from ?? 0;
        int end   = to   ?? all.Count - 1;
        return all.Where(p => p.Index >= start && p.Index <= end).ToList();
    }

    // Секция по названию стиля заголовка
    private static List<ParagraphInfo> SectionParagraphs(List<ParagraphInfo> all, string? sectionName)
    {
        if (string.IsNullOrWhiteSpace(sectionName)) return all;

        // Ищем заголовок секции
        int? sectionStart = null;
        int? sectionEnd   = null;

        for (int i = 0; i < all.Count; i++)
        {
            var p = all[i];
            bool isHeading = p.Style.Contains("Heading") || p.Style.Contains("Заголовок") ||
                             p.Style.Contains("heading");

            if (isHeading && p.Text.Contains(sectionName, StringComparison.OrdinalIgnoreCase))
            {
                sectionStart = i;
                continue;
            }

            // Конец секции — следующий заголовок того же или более высокого уровня
            if (sectionStart.HasValue && isHeading && !p.Text.Contains(sectionName, StringComparison.OrdinalIgnoreCase))
            {
                sectionEnd = i - 1;
                break;
            }
        }

        if (!sectionStart.HasValue) return all; // Секция не найдена

        int start = sectionStart.Value;
        int end   = sectionEnd   ?? all.Count - 1;
        return all.Where(p => p.Index >= start && p.Index <= end).ToList();
    }

    // Краткое оглавление для фазы 1 (очень мало токенов)
    public static string BuildSummary(List<ParagraphInfo> all)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"Всего абзацев: {all.Count}");
        sb.AppendLine("Структура (заголовки):");

        var headings = all
            .Where(p => p.Style.Contains("Heading") || p.Style.Contains("Заголовок") ||
                        p.Style.Contains("heading"))
            .Take(20)
            .ToList();

        if (headings.Count == 0)
        {
            // Нет заголовков — показываем первые абзацы
            sb.AppendLine("(нет явных заголовков)");
            foreach (var p in all.Take(10))
                sb.AppendLine($"  [{p.Index}] {p.Text[..Math.Min(p.Text.Length, 60)]}");
        }
        else
        {
            foreach (var h in headings)
                sb.AppendLine($"  [{h.Index}] ({h.Style}) {h.Text[..Math.Min(h.Text.Length, 80)]}");
        }

        return sb.ToString();
    }
}
