using System.Text.Json.Serialization;

namespace AIDocAssistant.Models;

public class AgentStep
{
    [JsonPropertyName("message")]    public string  Message    { get; set; } = "";
    [JsonPropertyName("tool")]       public string? Tool       { get; set; }
    [JsonPropertyName("query")]      public string? Query      { get; set; }  // search_text
    [JsonPropertyName("from")]       public int?    From       { get; set; }  // read_range
    [JsonPropertyName("to")]         public int?    To         { get; set; }  // read_range
    [JsonPropertyName("page")]       public int?    Page       { get; set; }  // read_page
    [JsonPropertyName("count")]      public int?    Count      { get; set; }  // read_tail
    [JsonPropertyName("file_id")]    public long?   FileId     { get; set; }  // read_workspace_file
    [JsonPropertyName("source_file_id")] public long? SourceFileId { get; set; } // compare/copy
    [JsonPropertyName("target_file_id")] public long? TargetFileId { get; set; }
    [JsonPropertyName("done")]       public bool    Done       { get; set; }
    [JsonPropertyName("operations")] public List<ChangeOperation>? Operations { get; set; }

    [JsonIgnore] public string RawJson { get; set; } = "";
}
