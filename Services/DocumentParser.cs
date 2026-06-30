using System.IO;
using System.Text.Json;
using AIDocAssistant.Models;

namespace AIDocAssistant.Services;

public class DocumentParser
{
    private static readonly HashSet<string> HeadingStyles = new(StringComparer.OrdinalIgnoreCase)
    {
        "Heading 1", "Heading 2", "Heading 3",
        "Заголовок 1", "Заголовок 2", "Заголовок 3"
    };

    public DocumentModel Parse(dynamic doc, List<ParagraphInfo> paragraphs)
    {
        return new DocumentModel
        {
            Document        = (string)doc.Name,
            TotalParagraphs = paragraphs.Count,
            Sections        = GroupIntoSections(paragraphs)
        };
    }

    public DocumentModel ParseMetadataOnly(dynamic doc, List<ParagraphInfo> paragraphs)
    {
        var stripped = paragraphs.Select(p => new ParagraphInfo
        {
            Index           = p.Index,
            PageNumber      = p.PageNumber,
            SectionNumber   = p.SectionNumber,
            Text            = string.IsNullOrWhiteSpace(p.Text) ? "" :
                              p.Text.Length > 30 ? p.Text[..30] + "…" : p.Text,
            Style           = p.Style,
            Font            = p.Font,
            Size            = p.Size,
            LineSpacing     = p.LineSpacing,
            LineSpacingRule = p.LineSpacingRule,
            Alignment       = p.Alignment,
            FirstLineIndent = p.FirstLineIndent,
            LeftIndent      = p.LeftIndent,
            RightIndent     = p.RightIndent,
            SpaceBefore     = p.SpaceBefore,
            SpaceAfter      = p.SpaceAfter,
            IsEmpty         = p.IsEmpty,
            IsInTable       = p.IsInTable,
            IsTableEndMarker= p.IsTableEndMarker
        }).ToList();

        return new DocumentModel
        {
            Document        = (string)doc.Name,
            TotalParagraphs = paragraphs.Count,
            Sections        = GroupIntoSections(stripped)
        };
    }

    public IEnumerable<DocumentModel> ParseChunked(dynamic doc,
        List<ParagraphInfo> paragraphs, int chunkSize = 80, bool metadataOnly = true)
    {
        if (paragraphs.Count <= chunkSize)
        {
            yield return metadataOnly
                ? ParseMetadataOnly(doc, paragraphs)
                : Parse(doc, paragraphs);
            yield break;
        }

        int chunkNum = 0;
        for (int i = 0; i < paragraphs.Count; i += chunkSize)
        {
            var chunk = paragraphs.Skip(i).Take(chunkSize).ToList();
            var model = metadataOnly
                ? ParseMetadataOnly(doc, chunk)
                : Parse(doc, chunk);

            model.Document = $"{(string)doc.Name} [часть {++chunkNum} / {(int)Math.Ceiling((double)paragraphs.Count / chunkSize)}]";
            yield return model;
        }
    }

    public string ToJson(DocumentModel model)
    {
        var options = new JsonSerializerOptions
        {
            WriteIndented       = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };
        return JsonSerializer.Serialize(model, options);
    }

    // Оценка кол-ва токенов (грубо: 1 токен ≈ 4 символа)
    public static int EstimateTokens(string json) => json.Length / 4;

    private static List<SectionModel> GroupIntoSections(List<ParagraphInfo> paragraphs)
    {
        var sections = new List<SectionModel>();
        var current  = new SectionModel { Title = "Начало документа" };

        foreach (var para in paragraphs)
        {
            if (HeadingStyles.Contains(para.Style) && !string.IsNullOrWhiteSpace(para.Text))
            {
                if (current.Paragraphs.Count > 0)
                    sections.Add(current);
                current = new SectionModel { Title = para.Text };
            }
            else
            {
                current.Paragraphs.Add(para);
            }
        }

        if (current.Paragraphs.Count > 0 || sections.Count == 0)
            sections.Add(current);

        return sections;
    }
}
