namespace AIDocAssistant.Models;

public enum OfficeFileKind
{
    Unknown,
    Word,
    Excel,
    PowerPoint
}

public class OfficeWorkspaceFile
{
    public long Id { get; set; }
    public OfficeFileKind Kind { get; set; }
    public string Name { get; set; } = "";
    public string FullPath { get; set; } = "";
    public DateTime OpenedAt { get; set; } = DateTime.Now;
    public bool IsSelected { get; set; }

    public string DisplayName => $"[{Id}] {Kind}: {Name}";
    public string ChatLabel => IsSelected ? $"✓ {DisplayName}" : DisplayName;
}

public class OfficeDocumentSnapshot
{
    public long FileId { get; set; }
    public OfficeFileKind Kind { get; set; }
    public string Name { get; set; } = "";
    public string FullPath { get; set; } = "";
    public string Summary { get; set; } = "";
}

public class OfficeApplyResult
{
    public int Applied { get; set; }
    public int Skipped { get; set; }
    public string Details { get; set; } = "";
}
