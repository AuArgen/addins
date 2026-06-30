namespace AIDocAssistant.Models;

public class ParagraphInfo
{
    public int Index { get; set; }
    public int PageNumber { get; set; }
    public int SectionNumber { get; set; }
    public string Text { get; set; } = string.Empty;
    public string Style { get; set; } = string.Empty;
    public string Font { get; set; } = string.Empty;
    public float Size { get; set; }
    public float LineSpacing { get; set; }
    public int LineSpacingRule { get; set; }
    public string Alignment { get; set; } = string.Empty;
    public float FirstLineIndent { get; set; }
    public float LeftIndent { get; set; }
    public float RightIndent { get; set; }
    public float SpaceBefore { get; set; }
    public float SpaceAfter { get; set; }
    // "auto" = от стиля, иначе: "red","yellow","blue","black",...
    public string Color { get; set; } = "auto";
    public bool IsEmpty { get; set; }
    public bool IsInTable { get; set; }
    public bool IsTableEndMarker { get; set; }
}

public class SectionModel
{
    public string Title { get; set; } = string.Empty;
    public List<ParagraphInfo> Paragraphs { get; set; } = [];
}

public class DocumentModel
{
    public string Document { get; set; } = string.Empty;
    public int TotalParagraphs { get; set; }
    public List<SectionModel> Sections { get; set; } = [];
}
