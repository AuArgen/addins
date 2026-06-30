namespace AIDocAssistant.Models;

public class ChatSessionInfo
{
    public long Id { get; set; }
    public string Title { get; set; } = "";
    public string DocumentName { get; set; } = "";
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public DateTime UpdatedAt { get; set; } = DateTime.Now;

    public string UpdatedAtStr => UpdatedAt.ToString("dd.MM HH:mm");
    public string DisplayTitle => string.IsNullOrWhiteSpace(Title) ? "Новый чат" : Title;
}

public class StoredChatMessage
{
    public long Id { get; set; }
    public long SessionId { get; set; }
    public DateTime Timestamp { get; set; } = DateTime.Now;
    public string Role { get; set; } = "";
    public string Content { get; set; } = "";
}
