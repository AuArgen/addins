using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIDocAssistant.Models;

namespace AIDocAssistant.Services;

public class AIService
{
    private readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(120) };
    private AppSettings _settings     = new();
    private string _systemPrompt      = string.Empty;
    private string _apiKey            = string.Empty;

    public void Configure(AppSettings settings, string apiKey, string systemPrompt)
    {
        _settings     = settings;
        _apiKey       = apiKey;
        _systemPrompt = systemPrompt;
    }

    // ── Агентский шаг — главный метод ────────────────────────────────────────

    public async Task<AgentStep> AgentStepAsync(
        List<(string role, string content)> history,
        CancellationToken ct = default)
    {
        string raw = _settings.Provider == AiProvider.Anthropic
            ? await SendHistoryAnthropicAsync(history, ct)
            : await SendHistoryOpenAiAsync(history, ct);

        return ParseAgentStep(raw);
    }

    // ── Низкоуровневые HTTP ───────────────────────────────────────────────────

    private async Task<string> SendHistoryOpenAiAsync(
        List<(string role, string content)> history,
        CancellationToken ct)
    {
        var messages = new List<object>
        {
            new { role = "system", content = _systemPrompt }
        };
        foreach (var (role, content) in history)
            messages.Add(new { role, content });

        var body = new
        {
            model       = _settings.Model,
            messages,
            max_tokens  = 2000,
            temperature = 0.1
        };

        var req = new HttpRequestMessage(HttpMethod.Post, AppSettings.Endpoint(_settings.Provider))
        {
            Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json")
        };
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);

        var resp = await _http.SendAsync(req, ct);
        resp.EnsureSuccessStatusCode();
        var json = JsonNode.Parse(await resp.Content.ReadAsStringAsync(ct));
        return json?["choices"]?[0]?["message"]?["content"]?.GetValue<string>() ?? "{}";
    }

    private async Task<string> SendHistoryAnthropicAsync(
        List<(string role, string content)> history,
        CancellationToken ct)
    {
        var messages = history.Select(h => new { role = h.role, content = h.content }).ToArray();

        var body = new
        {
            model      = _settings.Model,
            max_tokens = 2000,
            system     = _systemPrompt,
            messages
        };

        var req = new HttpRequestMessage(HttpMethod.Post, AppSettings.Endpoint(_settings.Provider))
        {
            Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json")
        };
        req.Headers.Add("x-api-key", _apiKey);
        req.Headers.Add("anthropic-version", "2023-06-01");

        var resp = await _http.SendAsync(req, ct);
        resp.EnsureSuccessStatusCode();
        var json = JsonNode.Parse(await resp.Content.ReadAsStringAsync(ct));
        return json?["content"]?[0]?["text"]?.GetValue<string>() ?? "{}";
    }

    // ── Парсинг ───────────────────────────────────────────────────────────────

    private static AgentStep ParseAgentStep(string raw)
    {
        try
        {
            var start = raw.IndexOf('{');
            var end   = raw.LastIndexOf('}');
            if (start >= 0 && end > start)
            {
                var jsonStr = raw[start..(end + 1)];
                var opts    = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var step    = JsonSerializer.Deserialize<AgentStep>(jsonStr, opts);
                if (step is not null)
                {
                    step.RawJson = jsonStr;
                    return step;
                }
            }
        }
        catch { }

        // Если AI вернул не JSON — показываем как сообщение и завершаем
        var msg = raw.Length > 500 ? raw[..500] + "…" : raw;
        return new AgentStep
        {
            Message = msg,
            Done    = true,
            RawJson = $"{{\"message\":{JsonSerializer.Serialize(msg)},\"done\":true}}"
        };
    }

    // ── Оставлено для совместимости (ParserChangeSet используется в HistoryViewModel) ──

    public static List<ChangeOperation> ParseChangeSet(string raw)
    {
        var start = raw.IndexOf('[');
        var end   = raw.LastIndexOf(']');
        if (start < 0 || end <= start) return [];
        try
        {
            var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            return JsonSerializer.Deserialize<List<ChangeOperation>>(raw[start..(end + 1)], opts) ?? [];
        }
        catch { return []; }
    }
}
