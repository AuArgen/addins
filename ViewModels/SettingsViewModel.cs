using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using AIDocAssistant.Models;
using AIDocAssistant.Services;

namespace AIDocAssistant.ViewModels;

public partial class SettingsViewModel : ObservableObject
{
    public event Action<AppSettings, string, string>? Saved;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(AvailableModels))]
    private AiProvider _selectedProvider = AiProvider.DeepSeek;

    [ObservableProperty]
    private string _apiKey = string.Empty;

    [ObservableProperty]
    private string _selectedModel = "deepseek-chat";

    [ObservableProperty]
    private int _maxTokens = 4000;

    [ObservableProperty]
    private bool _autoBackup = true;

    [ObservableProperty]
    private bool _autoAnalyzeOnConnect = false;

    [ObservableProperty]
    private int _chunkSize = 80;

    public List<AiProvider> Providers { get; } =
        [AiProvider.DeepSeek, AiProvider.OpenAI, AiProvider.OpenRouter, AiProvider.Anthropic];

    public List<string> AvailableModels => SelectedProvider switch
    {
        AiProvider.DeepSeek   => ["deepseek-chat", "deepseek-reasoner"],
        AiProvider.OpenAI     => ["gpt-4o-mini", "gpt-4o", "gpt-4-turbo"],
        AiProvider.OpenRouter => ["deepseek/deepseek-chat", "deepseek/deepseek-r1",
                                  "google/gemini-flash-1.5", "google/gemini-pro-1.5",
                                  "openai/gpt-4o-mini"],
        AiProvider.Anthropic  => ["claude-haiku-4-5", "claude-sonnet-4-6", "claude-opus-4-8"],
        _                     => []
    };

    partial void OnSelectedProviderChanged(AiProvider value)
    {
        SelectedModel = AppSettings.DefaultModel(value);
        var saved = CredentialService.Load(value.ToString());
        ApiKey = saved ?? string.Empty;
    }

    public SettingsViewModel() : this(new AppSettings()) { }

    public SettingsViewModel(AppSettings current)
    {
        _selectedProvider      = current.Provider;
        _selectedModel         = current.Model;
        _maxTokens             = current.MaxTokens;
        _autoBackup            = current.AutoBackup;
        _autoAnalyzeOnConnect  = current.AutoAnalyzeOnConnect;
        _chunkSize             = current.ChunkSize;

        var saved = CredentialService.Load(current.Provider.ToString());
        _apiKey = saved ?? string.Empty;
    }

    [RelayCommand]
    private void Save()
    {
        if (!string.IsNullOrWhiteSpace(ApiKey))
            CredentialService.Save(SelectedProvider.ToString(), ApiKey);

        var settings = new AppSettings
        {
            Provider             = SelectedProvider,
            Model                = SelectedModel,
            MaxTokens            = MaxTokens,
            AutoBackup           = AutoBackup,
            AutoAnalyzeOnConnect = AutoAnalyzeOnConnect,
            ChunkSize            = ChunkSize
        };

        var promptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Prompts", "system_prompt.txt");
        var prompt     = File.Exists(promptPath) ? File.ReadAllText(promptPath) : string.Empty;

        Saved?.Invoke(settings, ApiKey, prompt);
    }
}
