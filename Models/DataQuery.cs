using System.Text.Json.Serialization;

namespace AIDocAssistant.Models;

// Фаза 1'ден кайткан — AI кайсы маалымат керек экенин айтат
public class DataQuery
{
    [JsonPropertyName("mode")]
    public string Mode { get; set; } = "all"; // search | range | section | all

    [JsonPropertyName("search_text")]
    public string? SearchText { get; set; }   // mode=search болгондо

    [JsonPropertyName("full_text")]
    public bool FullText { get; set; } = true; // толук текстпи же metadata гана

    [JsonPropertyName("from_index")]
    public int? FromIndex { get; set; }        // mode=range болгондо

    [JsonPropertyName("to_index")]
    public int? ToIndex { get; set; }

    [JsonPropertyName("section_name")]
    public string? SectionName { get; set; }   // mode=section болгондо

    [JsonPropertyName("reason")]
    public string? Reason { get; set; }

    // Эгер AI Phase 2'сиз эле жооп бере алса — бул жерге операцияларды жазат
    // App Phase 2'ни өткөрүп жиберет жана сразу Word'го применяет
    [JsonPropertyName("direct_operations")]
    public List<ChangeOperation>? DirectOperations { get; set; }
}
