namespace AIDocAssistant.Models;

public enum AiProvider
{
    OpenAI,
    OpenRouter,
    Anthropic,
    DeepSeek
}

public class AppSettings
{
    public AiProvider Provider { get; set; } = AiProvider.DeepSeek;
    public string Model { get; set; } = "deepseek-chat";
    public int MaxTokens { get; set; } = 4000;
    public bool AutoBackup { get; set; } = true;
    public bool AutoAnalyzeOnConnect { get; set; } = false;
    public int ChunkSize { get; set; } = 80;  // абзацтар чанкке

    public static string DefaultModel(AiProvider provider) => provider switch
    {
        AiProvider.OpenAI      => "gpt-4o-mini",
        AiProvider.OpenRouter  => "deepseek/deepseek-chat",
        AiProvider.Anthropic   => "claude-haiku-4-5",
        AiProvider.DeepSeek    => "deepseek-chat",
        _                      => "deepseek-chat"
    };

    public static string Endpoint(AiProvider provider) => provider switch
    {
        AiProvider.OpenAI      => "https://api.openai.com/v1/chat/completions",
        AiProvider.OpenRouter  => "https://openrouter.ai/api/v1/chat/completions",
        AiProvider.Anthropic   => "https://api.anthropic.com/v1/messages",
        AiProvider.DeepSeek    => "https://api.deepseek.com/v1/chat/completions",
        _                      => "https://api.deepseek.com/v1/chat/completions"
    };
}
