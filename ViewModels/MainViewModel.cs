using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using AIDocAssistant.Models;
using AIDocAssistant.Services;
using System.Text;
using Microsoft.Win32;

namespace AIDocAssistant.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly WordInteropService _word    = new();
    private readonly DocumentParser     _parser  = new();
    private readonly AIService          _ai      = new();
    private readonly BackupService      _backup  = new();
    private readonly HistoryService     _history = new();
    private readonly ChatService        _chat    = new();
    private readonly OfficeWorkspaceService _office = new();
    private readonly DispatcherTimer    _timer   = new();

    private dynamic?               _activeDoc;
    private List<ChangeOperation>  _pendingChanges = [];
    private AppSettings            _settings       = new();
    private CancellationTokenSource? _cts;
    private bool _isRefreshing;
    private bool _isLoadingChatSession;
    private long _lastHistoryId; // ID последней записи в БД → обновляем applied_count после Apply
    private string _lastUserCommand = string.Empty;
    private List<(string role, string content)> _lastAgentHistory = [];

    public MainViewModel()
    {
        LoadSettings();
        ReloadChatSessions();
        _timer.Interval = TimeSpan.FromSeconds(5);
        _timer.Tick    += OnTimerTick;
        _timer.Start();
        // Первая проверка сразу при старте
        TryAutoConnect();
    }

    // ──────────────────── Состояние ────────────────────

    [ObservableProperty]
    private bool _isConnected;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ConnectStatusColor))]
    private string _connectStatus = "Не подключено";

    public string ConnectStatusColor => IsConnected ? "#10B981" : "#EF4444";

    [ObservableProperty] private ObservableCollection<string> _openDocuments = [];
    [ObservableProperty] private string? _selectedDocument;
    [ObservableProperty] private string _lastRefreshText = "—";
    [ObservableProperty] private ObservableCollection<ChatSessionInfo> _chatSessions = [];
    [ObservableProperty] private ChatSessionInfo? _selectedChatSession;
    [ObservableProperty] private ObservableCollection<OfficeWorkspaceFile> _workspaceFiles = [];
    [ObservableProperty] private OfficeWorkspaceFile? _selectedWorkspaceFile;
    [ObservableProperty] private string _chatResourceSummary = "Ресурсы чата не выбраны";
    [ObservableProperty] private ObservableCollection<ChatMessage> _chatMessages = [];
    [ObservableProperty] private ObservableCollection<ChangeOperationItem> _pendingItems = [];
    [ObservableProperty] private string _userInput = string.Empty;
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private bool _hasChanges;
    [ObservableProperty] private string _busyText = string.Empty;
    [ObservableProperty] private bool _autoAnalyzeEnabled;

    // Прогресс (0–100)
    [ObservableProperty] private double _progress;

    // Кнопка «Отменить» видна только когда можно отменить
    [ObservableProperty] private bool _canCancel;

    // Статистика документа (в сайдбаре)
    [ObservableProperty] private string _docStats = "—";

    // Количество ожидающих операций
    [ObservableProperty] private int _pendingCount;

    // ──────────────────── Подключение ────────────────────

    // ──────────────────── Авто-обновление списка файлов ────────────────────

    private void OnTimerTick(object? sender, EventArgs e) => TryAutoConnect();

    private void TryAutoConnect(bool force = false)
    {
        if (IsBusy) return;

        try
        {
            // Попытка подключиться если ещё нет
            if (!IsConnected)
            {
                if (!_word.TryConnect()) { ConnectStatus = "Word не запущен"; return; }

                IsConnected   = true;
                ConnectStatus = "Подключено";
                AddSystemMessage("✓ Word обнаружен, подключено автоматически.");
            }

            var names   = _word.GetOpenDocumentNames();
            var current = OpenDocuments.ToList();

            bool changed = force ||
                           names.Count != current.Count ||
                           names.Any(n => !current.Contains(n));

            if (changed)
            {
                var prev = SelectedDocument;
                _isRefreshing = true; // блокируем SelectDocument во время Clear/Add
                try
                {
                    OpenDocuments.Clear();
                    foreach (var n in names) OpenDocuments.Add(n);
                    SelectedDocument = names.Contains(prev ?? "") ? prev : (names.Count > 0 ? names[0] : null);
                }
                finally
                {
                    _isRefreshing = false;
                }

                // Если документ сменился — обновляем ссылку
                if (SelectedDocument != prev)
                    _activeDoc = SelectedDocument is not null ? _word.GetDocumentByName(SelectedDocument) : null;

                // Предупреждения о backup-файлах
                foreach (var name in OpenDocuments)
                    if (BackupService.IsBackupFile(name) && !current.Contains(name))
                        AddSystemMessage($"⚠ «{name}» — backup-файл. Откройте оригинал.");
            }

            // Если Word закрыт между тиками
            if (names.Count == 0 && IsConnected && OpenDocuments.Count == 0)
            {
                IsConnected      = false;
                ConnectStatus    = "Word не запущен";
                _activeDoc       = null;
            }

            LastRefreshText = $"{DateTime.Now:HH:mm:ss}";
        }
        catch
        {
            // Word закрыт или COM упал
            IsConnected   = false;
            ConnectStatus = "Word не запущен";
            _activeDoc    = null;
            OpenDocuments.Clear();
        }
    }

    [RelayCommand]
    private void OpenHistory()
    {
        try
        {
            var vm  = new HistoryViewModel(_history);
            var win = new Views.HistoryWindow { DataContext = vm };
            win.Show();
        }
        catch (Exception ex)
        {
            AddSystemMessage($"✗ История: {ex.Message}");
        }
    }

    [RelayCommand]
    private void RefreshDocs() => TryAutoConnect(force: true);

    [RelayCommand]
    private void OpenDocumentFile()
    {
        if (IsBusy) return;

        var dlg = new OpenFileDialog
        {
            Title = "Открыть Office файл",
            Filter = "Office файлы (*.docx;*.doc;*.docm;*.xlsx;*.xls;*.xlsm;*.pptx;*.ppt;*.pptm)|*.docx;*.doc;*.docm;*.xlsx;*.xls;*.xlsm;*.pptx;*.ppt;*.pptm|Word (*.docx;*.doc;*.docm)|*.docx;*.doc;*.docm|Excel (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|PowerPoint (*.pptx;*.ppt;*.pptm)|*.pptx;*.ppt;*.pptm|Все файлы (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            var workspaceFile = _office.AddFile(dlg.FileName);
            workspaceFile.IsSelected = true;
            ReloadWorkspaceFiles(workspaceFile.Id);
            SaveSelectedChatResources(silent: true);

            if (workspaceFile.Kind == OfficeFileKind.Word)
            {
                var doc = _word.OpenDocumentFromFile(dlg.FileName, visible: true);
                var docName = (string)doc.Name;

                IsConnected = true;
                ConnectStatus = "Подключено";
                _activeDoc = doc;

                TryAutoConnect(force: true);

                if (!OpenDocuments.Contains(docName))
                    OpenDocuments.Add(docName);

                SelectedDocument = docName;
                _activeDoc = doc;
                AddSystemMessage($"✓ Word документ открыт и добавлен в workspace: [{workspaceFile.Id}] {docName}");
            }
            else
            {
                AddSystemMessage($"✓ Файл добавлен в workspace: [{workspaceFile.Id}] {workspaceFile.Kind} {workspaceFile.Name}");
            }
        }
        catch (Exception ex)
        {
            AddSystemMessage($"✗ Не удалось открыть Office файл: {ex.Message}");
        }
    }

    [RelayCommand]
    private void SelectDocument()
    {
        if (_isRefreshing || SelectedDocument is null) return;
        _activeDoc = _word.GetDocumentByName(SelectedDocument);
        if (_activeDoc is not null)
            AddSystemMessage($"Документ выбран: {SelectedDocument}");
    }

    partial void OnSelectedChatSessionChanged(ChatSessionInfo? value)
    {
        if (_isLoadingChatSession) return;
        if (value is not null)
            LoadChatSession(value);
    }

    [RelayCommand]
    private void NewChat()
    {
        _isLoadingChatSession = true;
        try
        {
            SelectedChatSession = null;
            ChatMessages.Clear();
            ClearPendingChanges();
            _lastUserCommand = string.Empty;
            _lastAgentHistory.Clear();
            _lastHistoryId = 0;
            foreach (var file in _office.Files)
                file.IsSelected = false;
            WorkspaceFiles = new ObservableCollection<OfficeWorkspaceFile>(_office.Files);
            UpdateChatResourceSummary();
        }
        finally
        {
            _isLoadingChatSession = false;
        }
    }

    private void ReloadChatSessions()
    {
        ChatSessions = new ObservableCollection<ChatSessionInfo>(_chat.GetSessions());
    }

    private void ReloadWorkspaceFiles(long? selectedId = null)
    {
        var selectedPaths = SelectedChatSession is null
            ? WorkspaceFiles.Where(f => f.IsSelected).Select(f => f.FullPath).ToHashSet(StringComparer.OrdinalIgnoreCase)
            : _chat.GetResourcePaths(SelectedChatSession.Id).ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var file in _office.Files)
            file.IsSelected = selectedPaths.Contains(file.FullPath) || file.Id == selectedId;

        WorkspaceFiles = new ObservableCollection<OfficeWorkspaceFile>(_office.Files);
        SelectedWorkspaceFile = selectedId.HasValue
            ? WorkspaceFiles.FirstOrDefault(f => f.Id == selectedId.Value)
            : WorkspaceFiles.FirstOrDefault();
        UpdateChatResourceSummary();
    }

    [RelayCommand]
    private void SaveSelectedChatResources() => SaveSelectedChatResources(silent: false);

    [RelayCommand]
    private void SelectAllWorkspaceResources()
    {
        foreach (var file in _office.Files)
            file.IsSelected = true;
        WorkspaceFiles = new ObservableCollection<OfficeWorkspaceFile>(_office.Files);
        SaveSelectedChatResources(silent: false);
    }

    [RelayCommand]
    private void ClearWorkspaceResources()
    {
        foreach (var file in _office.Files)
            file.IsSelected = false;
        WorkspaceFiles = new ObservableCollection<OfficeWorkspaceFile>(_office.Files);
        SaveSelectedChatResources(silent: false);
    }

    private void SaveSelectedChatResources(bool silent)
    {
        if (SelectedChatSession is null)
        {
            UpdateChatResourceSummary();
            return;
        }

        PersistCurrentResourceSelection();
        if (!silent)
            AddSystemMessage($"Ресурсы чата сохранены: {_office.Files.Count(f => f.IsSelected)} файл(ов).");
    }

    private void PersistCurrentResourceSelection()
    {
        if (SelectedChatSession is null) return;

        var paths = _office.Files
            .Where(f => f.IsSelected)
            .Select(f => f.FullPath)
            .ToList();
        _chat.SetResourcePaths(SelectedChatSession.Id, paths);
        UpdateChatResourceSummary();
    }

    private void ApplyChatResourceSelection(long sessionId)
    {
        var paths = _chat.GetResourcePaths(sessionId).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var file in _office.Files)
            file.IsSelected = paths.Contains(file.FullPath);
        WorkspaceFiles = new ObservableCollection<OfficeWorkspaceFile>(_office.Files);
        UpdateChatResourceSummary();
    }

    private List<long> ActiveWorkspaceFileIds() =>
        _office.Files.Where(f => f.IsSelected).Select(f => f.Id).ToList();

    private void UpdateChatResourceSummary()
    {
        var count = _office.Files.Count(f => f.IsSelected);
        ChatResourceSummary = count == 0
            ? "Ресурсы чата не выбраны"
            : $"Выбрано ресурсов для этого чата: {count}";
    }

    private void LoadChatSession(ChatSessionInfo session)
    {
        _isLoadingChatSession = true;
        try
        {
            ChatMessages.Clear();
            foreach (var msg in _chat.GetMessages(session.Id))
                ChatMessages.Add(new ChatMessage(msg.Content, ParseMessageRole(msg.Role)));

            ClearPendingChanges();
            _lastUserCommand = string.Empty;
            _lastAgentHistory.Clear();
            _lastHistoryId = 0;
            ApplyChatResourceSelection(session.Id);
        }
        finally
        {
            _isLoadingChatSession = false;
        }
    }

    private void EnsureCurrentChatSession(string firstMessage)
    {
        if (SelectedChatSession is not null) return;

        var session = _chat.CreateSession(
            MakeChatTitle(firstMessage),
            SelectedDocument ?? string.Empty);

        _isLoadingChatSession = true;
        try
        {
            ChatSessions.Insert(0, session);
            SelectedChatSession = session;
            SaveSelectedChatResources(silent: true);
        }
        finally
        {
            _isLoadingChatSession = false;
        }
    }

    private static string MakeChatTitle(string text)
    {
        var title = text.Trim().Replace("\r", " ").Replace("\n", " ");
        if (title.Length > 42) title = title[..42] + "...";
        return string.IsNullOrWhiteSpace(title) ? "Новый чат" : title;
    }

    private static MessageRole ParseMessageRole(string role) =>
        Enum.TryParse<MessageRole>(role, ignoreCase: true, out var parsed)
            ? parsed
            : MessageRole.System;

    // ──────────────────── Отправка ────────────────────

    [RelayCommand]
    private async Task SendMessageAsync()
    {
        if (string.IsNullOrWhiteSpace(UserInput) || IsBusy) return;
        var command = UserInput.Trim();
        UserInput = string.Empty;
        AddUserMessage(command);
        await RunAgentLoopAsync(command);
    }

    [RelayCommand]
    private void CancelOperation()
    {
        _cts?.Cancel();
        CanCancel = false;
        AddSystemMessage("⏹ Операция отменена.");
    }

    [RelayCommand]
    private void ClearChat() => NewChat();

    // ──────────────────── Быстрые команды ────────────────────

    [RelayCommand]
    private async Task QuickCheckAllAsync()   => await RunAgentLoopAsync(QuickCheckAllPrompt);

    [RelayCommand]
    private async Task QuickFixFontsAsync()   => await RunAgentLoopAsync(
        "Установи шрифт Times New Roman 14pt для основного текста. " +
        "Используй modify_style для стиля Normal (или Обычный). Сразу done:true.");

    [RelayCommand]
    private async Task QuickFixIndentsAsync() => await RunAgentLoopAsync(
        "Установи отступ первой строки 1.25 см для всего документа. " +
        "Используй set_all_indent с first_line_cm:1.25. Сразу done:true.");

    [RelayCommand]
    private async Task QuickFixSpacingAsync() => await RunAgentLoopAsync(
        "Установи межстрочный интервал 1.5 для всего документа. " +
        "Используй set_all_spacing с line_spacing:1.5. Сразу done:true.");

    [RelayCommand]
    private async Task QuickFixAlignAsync()   => await RunAgentLoopAsync(
        "Установи выравнивание по ширине (justify) для всего документа. " +
        "Используй set_all_alignment с alignment:justify. Сразу done:true.");

    // ──────────────────── Применение ────────────────────

    [RelayCommand]
    private void ShowPreview() => HasChanges = _pendingChanges.Count > 0;

    [RelayCommand]
    private async Task ApplyPendingChangesAsync()
    {
        var selected = PendingItems
            .Where(i => i.IsSelected)
            .Select(i => i.Operation)
            .ToList();

        ClearPendingChanges();
        if (selected.Count > 0)
            await ApplyChangesAsync(selected);
    }

    [RelayCommand]
    private void CancelPendingChanges()
    {
        ClearPendingChanges();
        AddSystemMessage("Операции отменены.");
    }

    private void SetPendingChanges(List<ChangeOperation> operations)
    {
        _pendingChanges = operations;
        PendingItems = new ObservableCollection<ChangeOperationItem>(
            operations.Select(op => new ChangeOperationItem(op)));
        PendingCount = operations.Count;
        HasChanges = operations.Count > 0;
    }

    private void ClearPendingChanges()
    {
        _pendingChanges.Clear();
        PendingItems.Clear();
        PendingCount = 0;
        HasChanges = false;
    }

    [RelayCommand]
    private void OpenSettings()
    {
        var vm  = new SettingsViewModel(_settings);
        var win = new Views.SettingsWindow { DataContext = vm };
        vm.Saved += (settings, key, prompt) =>
        {
            _settings          = settings;
            AutoAnalyzeEnabled = settings.AutoAnalyzeOnConnect;
            SaveSettingsToFile(settings);
            ConfigureAI();
            win.Close();
            AddSystemMessage($"✓ Настройки: {settings.Provider} / {settings.Model}");
        };
        win.ShowDialog();
    }

    // ──────────────────── Агентский цикл ────────────────────

    private async Task RunAgentLoopAsync(string command)
    {
        SaveSelectedChatResources(silent: true);
        var activeResourceIds = ActiveWorkspaceFileIds();

        if (false && _activeDoc is null && activeResourceIds.Count == 0)
        {
            AddSystemMessage("Документ не выбран.");
            return;
        }
        if (IsBusy) return;

        _cts      = new CancellationTokenSource();
        var ct    = _cts.Token;
        IsBusy    = true;
        CanCancel = true;
        ClearPendingChanges();
        Progress     = 0;

        List<ParagraphInfo>? paragraphs = null;
        var doc     = _activeDoc;
        var allOps  = new List<ChangeOperation>();
        var history = new List<(string role, string content)>();

        // Ленивая загрузка документа — только когда AI вызвал инструмент
        async Task<List<ParagraphInfo>> EnsureParagraphsAsync()
        {
            if (paragraphs is not null) return paragraphs;
            if (doc is null)
            {
                paragraphs = [];
                return paragraphs;
            }
            SetProgress(20, "Читаю документ...");
            paragraphs = await Task.Run(() => (List<ParagraphInfo>)_word.ReadParagraphs(doc), ct);
            var tokens = DocumentParser.EstimateTokens(string.Join(" ", paragraphs.Select(p => p.Text)));
            var pages = paragraphs.Count == 0 ? 0 : paragraphs.Max(p => p.PageNumber);
            DocStats = $"{paragraphs.Count} абз · {pages} стр · ~{tokens / 1000}K токенов";
            return paragraphs;
        }

        try
        {
            SetProgress(5, "Отправляю...");

            var wordContext = doc is null
                ? "WORD_CONTEXT: active Word document is not selected. Only Office workspace read/list/compare tools are available."
                : _word.GetWordContext(doc);
            history.Add(("user",
                $"{wordContext}\n" +
                $"Документ: {SelectedDocument ?? "—"}\n" +
                $"OFFICE_WORKSPACE:\n{_office.ListFiles(activeResourceIds)}\n" +
                $"CHAT_TRANSCRIPT:\n{BuildRecentChatTranscript()}\n\n" +
                $"Задача: {command}"));

            const int maxSteps = 10;
            for (int step = 0; step < maxSteps; step++)
            {
                ct.ThrowIfCancellationRequested();
                SetProgress(10 + (double)step / maxSteps * 85,
                    step == 0 ? "Думаю..." : $"Шаг {step + 1}...");

                var aiStep = await _ai.AgentStepAsync(history, ct);
                var assistantStepContent = aiStep.RawJson.Length > 0
                    ? aiStep.RawJson
                    : JsonSerializer.Serialize(aiStep);

                // Показываем ответ AI в чате
                if (!string.IsNullOrEmpty(aiStep.Message))
                    AddAiMessage(aiStep.Message);

                // AI завершил задачу
                if (aiStep.Done)
                {
                    history.Add(("assistant", assistantStepContent));
                    SetProgress(100, string.Empty);
                    allOps = aiStep.Operations ?? [];
                    if (allOps.Count > 0)
                        SetPendingChanges(allOps);
                    break;
                }

                // AI хочет использовать инструмент
                if (!string.IsNullOrEmpty(aiStep.Tool))
                {
                    var paras = IsWorkspaceTool(aiStep.Tool)
                        ? new List<ParagraphInfo>()
                        : await EnsureParagraphsAsync();
                    var toolResult = ExecuteTool(aiStep, paras);
                    history.Add(("assistant", assistantStepContent));
                    // Напоминание о задаче при каждом ответе инструмента — AI не "забывает" задачу
                    history.Add(("user",
                        $"[{aiStep.Tool}] результат:\n{toolResult}\n\n" +
                        $"───\nЗадача: {command}\n" +
                        $"Теперь сгенерируй operations и верни done:true."));
                }
                else
                {
                    // Ни done, ни tool — принудительно завершаем
                    break;
                }
            }
        }
        catch (OperationCanceledException)
        {
            SetProgress(0, string.Empty);
        }
        catch (Exception ex)
        {
            AddSystemMessage($"✗ {ex.Message}");
            SetProgress(0, string.Empty);
        }
        finally
        {
            IsBusy    = false;
            CanCancel = false;
            _cts?.Dispose();
            _cts = null;
        }

        _lastUserCommand = command;
        _lastAgentHistory = history.ToList();
        _lastHistoryId = SaveHistory(command, allOps);
    }

    // ──────────────────── Инструменты AI ────────────────────

    private string ExecuteTool(AgentStep step, List<ParagraphInfo> paragraphs) =>
        step.Tool switch
        {
            "search_text"  => ToolSearch(step.Query ?? "", paragraphs),
            "read_range"   => ToolReadRange(step.From ?? 0, step.To ?? 30, paragraphs),
            "read_tail"    => ToolReadTail(step.Count ?? 25, paragraphs),
            "read_page"    => ToolReadPage(step.Page ?? 1, paragraphs),
            "get_document_map" => ToolGetDocumentMap(paragraphs),
            "scan_formatting" => ToolScanFormatting(paragraphs),
            "get_headings" => ToolGetHeadings(paragraphs),
            "scan_colors"  => ToolScanColors(paragraphs),
            "list_workspace_files" => _office.ListFiles(ActiveWorkspaceFileIds()),
            "read_workspace_file" => ToolReadWorkspaceFile(step.FileId),
            "search_workspace_files" => _office.SearchFiles(step.Query ?? "", ActiveWorkspaceFileIds()),
            "compare_workspace_files" => ToolCompareWorkspaceFiles(step.SourceFileId, step.TargetFileId),
            _              => $"Инструмент «{step.Tool}» не найден."
        };

    private static bool IsWorkspaceTool(string? tool) =>
        tool is "list_workspace_files" or "read_workspace_file" or "search_workspace_files" or "compare_workspace_files";

    private string ToolReadWorkspaceFile(long? fileId)
    {
        if (!fileId.HasValue)
            return "Нужен file_id. Сначала вызови list_workspace_files.";

        if (!ActiveWorkspaceFileIds().Contains(fileId.Value))
            return $"file_id={fileId.Value} is not selected for this chat. User must select it in chat resources.";

        try
        {
            return _office.ReadFile(fileId.Value).Summary;
        }
        catch (Exception ex)
        {
            return $"Не удалось прочитать workspace file_id={fileId.Value}: {ex.Message}";
        }
    }

    private string ToolCompareWorkspaceFiles(long? sourceFileId, long? targetFileId)
    {
        if (!sourceFileId.HasValue || !targetFileId.HasValue)
            return "Нужны source_file_id и target_file_id. Сначала вызови list_workspace_files.";

        var ids = ActiveWorkspaceFileIds();
        if (!ids.Contains(sourceFileId.Value) || !ids.Contains(targetFileId.Value))
            return "One or both file ids are not selected for this chat. User must select them in chat resources.";

        try
        {
            return _office.CompareFiles(sourceFileId.Value, targetFileId.Value);
        }
        catch (Exception ex)
        {
            return $"Не удалось сравнить файлы workspace: {ex.Message}";
        }
    }

    private static string ToolSearch(string query, List<ParagraphInfo> paragraphs)
    {
        if (string.IsNullOrEmpty(query)) return "Нет поискового запроса.";
        var matches = paragraphs
            .Where(p => p.Text.Contains(query, StringComparison.OrdinalIgnoreCase))
            .Take(30).ToList();
        if (matches.Count == 0) return $"По запросу «{query}» ничего не найдено.";
        var sb = new StringBuilder();
        sb.AppendLine($"Найдено {matches.Count} абзацев:");
        foreach (var p in matches)
            sb.AppendLine(FormatParagraphLine(p, 200));
        return sb.ToString();
    }

    private static string ToolReadRange(int from, int to, List<ParagraphInfo> paragraphs)
    {
        // Ограничиваем диапазон до 30 абзацев, текст до 120 символов
        var range = paragraphs.Where(p => p.Index >= from && p.Index <= to).Take(30).ToList();
        if (range.Count == 0) return "Диапазон пуст.";
        var sb = new StringBuilder();
        sb.AppendLine($"Абзацев в диапазоне [{from}–{to}]: {range.Count} (всего в документе: {paragraphs.Count})");
        foreach (var p in range)
            sb.AppendLine(FormatParagraphLine(p, 160));
        return sb.ToString();
    }

    private static string ToolReadTail(int count, List<ParagraphInfo> paragraphs)
    {
        var take = Math.Clamp(count, 1, 80);
        var start = Math.Max(0, paragraphs.Count - take);
        return ToolReadRange(start, paragraphs.Count - 1, paragraphs);
    }

    private static string ToolReadPage(int page, List<ParagraphInfo> paragraphs)
    {
        var pageParas = paragraphs
            .Where(p => p.PageNumber == page)
            .Take(120)
            .ToList();

        if (pageParas.Count == 0)
            return $"Страница {page} не найдена. Доступные страницы: {PageRangeText(paragraphs)}.";

        var sb = new StringBuilder();
        sb.AppendLine($"Страница {page}: {pageParas.Count} абзацев (показано до 120)");
        foreach (var p in pageParas)
            sb.AppendLine(FormatParagraphLine(p, 180));
        return sb.ToString();
    }

    private static string ToolGetDocumentMap(List<ParagraphInfo> paragraphs)
    {
        var sb = new StringBuilder();
        var maxPage = paragraphs.Count == 0 ? 0 : paragraphs.Max(p => p.PageNumber);
        sb.AppendLine($"DOCUMENT_MAP: paragraphs={paragraphs.Count}, pages={maxPage}");
        sb.AppendLine("Legend: T=inside table, M=table end marker, E=empty real paragraph");

        foreach (var page in paragraphs.GroupBy(p => p.PageNumber).OrderBy(g => g.Key))
        {
            var real = page.Count(p => !p.IsTableEndMarker);
            var markers = page.Count(p => p.IsTableEndMarker);
            var tables = page.Count(p => p.IsInTable);
            var first = page.First().Index;
            var last = page.Last().Index;
            var sample = page.FirstOrDefault(p => !p.IsEmpty && !p.IsTableEndMarker)?.Text ?? "";
            if (sample.Length > 70) sample = sample[..70] + "…";

            sb.AppendLine($"page={page.Key}: p[{first}..{last}], real={real}, tableParas={tables}, markers={markers}, sample=\"{sample}\"");
        }

        sb.AppendLine("Tail:");
        foreach (var p in paragraphs.TakeLast(12))
            sb.AppendLine(FormatParagraphLine(p, 120));
        return sb.ToString();
    }

    private static string ToolGetHeadings(List<ParagraphInfo> paragraphs)
    {
        var heads = paragraphs
            .Where(p => p.Style.Contains("Heading", StringComparison.OrdinalIgnoreCase) ||
                        p.Style.Contains("Заголов", StringComparison.OrdinalIgnoreCase) ||
                        p.Style.Contains("Title",   StringComparison.OrdinalIgnoreCase))
            .Take(60)
            .Select(p => $"[{p.Index}] {p.Style} color={p.Color}: {p.Text}");
        var result = string.Join("\n", heads);
        return string.IsNullOrEmpty(result)
            ? $"Заголовков не найдено. Всего абзацев: {paragraphs.Count}."
            : $"Заголовки ({paragraphs.Count} абз. всего):\n{result}";
    }

    private static string ToolScanColors(List<ParagraphInfo> paragraphs)
    {
        var groups = paragraphs
            .Where(p => !string.IsNullOrWhiteSpace(p.Text))
            .GroupBy(p => p.Color)
            .OrderByDescending(g => g.Count())
            .Take(15);
        var sb = new StringBuilder();
        sb.AppendLine($"Цвета в документе (всего {paragraphs.Count} абзацев):");
        foreach (var g in groups)
        {
            var sample = g.First();
            var sampleText = sample.Text.Length > 60 ? sample.Text[..60] + "…" : sample.Text;
            sb.AppendLine($"  {g.Key}: {g.Count()} абз. | пример[{sample.Index}]: \"{sampleText}\"");
        }
        return sb.ToString();
    }

    private static string ToolScanFormatting(List<ParagraphInfo> paragraphs)
    {
        var real = paragraphs
            .Where(p => !p.IsTableEndMarker && !string.IsNullOrWhiteSpace(p.Text))
            .ToList();

        var sb = new StringBuilder();
        sb.AppendLine($"FORMAT_SCAN: realTextParagraphs={real.Count}, allParagraphs={paragraphs.Count}");
        sb.AppendLine("Line spacing note: lineRule is more important than rawSp. rawSp near 1.5 with rule=Exactly/AtLeast can visually collapse text.");
        AppendGroups(sb, "styles", real.GroupBy(p => p.Style).OrderByDescending(g => g.Count()).Take(12), g => g.Key);
        AppendGroups(sb, "fonts", real.GroupBy(p => $"{p.Font} {p.Size:0.##}pt").OrderByDescending(g => g.Count()).Take(12), g => g.Key);
        AppendGroups(sb, "lineSpacing", real.GroupBy(p => $"{LineRuleName(p.LineSpacingRule)} raw={p.LineSpacing:0.##}").OrderByDescending(g => g.Count()).Take(12), g => g.Key);
        AppendGroups(sb, "firstLineIndentCm", real.GroupBy(p => Math.Round(p.FirstLineIndent, 2)).OrderByDescending(g => g.Count()).Take(12), g => g.Key.ToString("0.##"));
        AppendGroups(sb, "leftIndentCm", real.GroupBy(p => Math.Round(p.LeftIndent, 2)).OrderByDescending(g => g.Count()).Take(12), g => g.Key.ToString("0.##"));
        AppendGroups(sb, "spaceBeforeAfter", real.GroupBy(p => $"before={p.SpaceBefore:0.##}, after={p.SpaceAfter:0.##}").OrderByDescending(g => g.Count()).Take(12), g => g.Key);

        var suspicious = real
            .Where(p => (p.LineSpacingRule is 3 or 4 && p.LineSpacing < 10) ||
                        (p.LineSpacingRule == 5 && p.LineSpacing < 12) ||
                        (p.Size > 0 && p.LineSpacingRule is 3 or 4 && p.LineSpacing < p.Size))
            .Take(40)
            .ToList();

        sb.AppendLine($"suspiciousLineSpacing={suspicious.Count}");
        foreach (var p in suspicious)
            sb.AppendLine(FormatParagraphLine(p, 140));

        sb.AppendLine("Tail sample:");
        foreach (var p in paragraphs.TakeLast(15))
            sb.AppendLine(FormatParagraphLine(p, 120));

        return sb.ToString();
    }

    private static void AppendGroups<TKey>(
        StringBuilder sb,
        string title,
        IEnumerable<IGrouping<TKey, ParagraphInfo>> groups,
        Func<IGrouping<TKey, ParagraphInfo>, string> label)
    {
        sb.AppendLine(title + ":");
        foreach (var group in groups)
        {
            var sample = group.First();
            sb.AppendLine($"  {label(group)}: {group.Count()} абз. sample=[{sample.Index}] page={sample.PageNumber}");
        }
    }

    private static string FormatParagraphLine(ParagraphInfo p, int maxText)
    {
        var flags = new List<string>();
        if (p.IsInTable) flags.Add("T");
        if (p.IsTableEndMarker) flags.Add("M");
        if (p.IsEmpty && !p.IsTableEndMarker) flags.Add("E");
        var flagText = flags.Count == 0 ? "-" : string.Join("", flags);
        var text = p.Text.Length > maxText ? p.Text[..maxText] + "…" : p.Text;
        return $"[{p.Index}] page={p.PageNumber} sec={p.SectionNumber} flags={flagText} {p.Style} | {p.Font} {p.Size}pt | color={p.Color} | lineRule={LineRuleName(p.LineSpacingRule)} rawSp={p.LineSpacing:F1} first={p.FirstLineIndent:F2}cm left={p.LeftIndent:F2}cm right={p.RightIndent:F2}cm before={p.SpaceBefore:F1} after={p.SpaceAfter:F1} | \"{text}\"";
    }

    private static string LineRuleName(int rule) => rule switch
    {
        0 => "Single",
        1 => "1.5",
        2 => "Double",
        3 => "AtLeast",
        4 => "Exactly",
        5 => "Multiple",
        _ => $"Unknown({rule})"
    };

    private static string PageRangeText(List<ParagraphInfo> paragraphs)
    {
        var pages = paragraphs
            .Select(p => p.PageNumber)
            .Where(p => p > 0)
            .Distinct()
            .Order()
            .ToList();
        return pages.Count == 0 ? "нет данных" : string.Join(", ", pages);
    }

    private async Task ApplyChangesAsync(IEnumerable<ChangeOperation> ops)
    {
        var opList = ops.ToList();
        if (opList.Count == 0) return;
        var workspaceOps = opList.Where(OfficeWorkspaceService.IsWorkspaceOperation).ToList();

        if (_activeDoc is null && workspaceOps.Count != opList.Count)
        {
            AddSystemMessage("Word document is not selected for Word operations.");
            return;
        }

        if (workspaceOps.Count == opList.Count)
        {
            bool retryOffice = false;
            string officeReport = string.Empty;
            try
            {
                var officeResult = await Task.Run(() => _office.ApplyChanges(workspaceOps));
                PersistCurrentResourceSelection();
                WorkspaceFiles = new ObservableCollection<OfficeWorkspaceFile>(_office.Files);
                UpdateChatResourceSummary();
                if (_lastHistoryId > 0)
                    try { _history.UpdateApplied(_lastHistoryId, officeResult.Applied); } catch { }

                if (officeResult.Applied == 0 || officeResult.Skipped > 0)
                {
                    AddSystemMessage($"⚠ Office workspace: applied {officeResult.Applied}, skipped {officeResult.Skipped}.");
                    retryOffice = true;
                    officeReport = BuildApplyReport(officeResult.Applied, officeResult.Skipped, officeResult.Applied, officeResult.Details);
                }
                else
                {
                    AddSystemMessage($"✓ Office workspace: applied {officeResult.Applied} operations.");
                }
            }
            catch (Exception ex)
            {
                AddSystemMessage($"✗ Office workspace error: {ex.Message}");
                retryOffice = true;
                officeReport = $"Office workspace exception: {ex.GetType().Name}: {ex.Message}";
            }
            finally
            {
                ClearPendingChanges();
            }

            if (retryOffice)
                await ContinueAgentAfterApplyFailureAsync(officeReport, workspaceOps);
            return;
        }

        OfficeApplyResult? appliedWorkspace = null;
        if (workspaceOps.Count > 0)
        {
            appliedWorkspace = await Task.Run(() => _office.ApplyChanges(workspaceOps));
            PersistCurrentResourceSelection();
            WorkspaceFiles = new ObservableCollection<OfficeWorkspaceFile>(_office.Files);
            UpdateChatResourceSummary();
            opList = opList.Where(op => !OfficeWorkspaceService.IsWorkspaceOperation(op)).ToList();
        }

        bool askAgentToContinue = false;
        string applyReport = string.Empty;

        try
        {
            var docName = (string)_activeDoc!.Name;
            if (BackupService.IsBackupFile(docName))
            {
                AddSystemMessage("⚠ Вы работаете с backup-файлом. Откройте оригинал.");
                return;
            }

            // Backup — некритичен: ошибка не останавливает применение
            if (_settings.AutoBackup)
            {
                try
                {
                    var backupPath = _backup.CreateBackup(_activeDoc);
                    AddSystemMessage($"Backup: {Path.GetFileName(backupPath)}");
                }
                catch (Exception ex)
                {
                    AddSystemMessage($"⚠ Backup не создан ({ex.Message[..Math.Min(ex.Message.Length, 60)]}), применяю без него...");
                }
            }

            var result       = (ValueTuple<int, int, int>)_word.ApplyChanges(_activeDoc, opList);
            int applied      = result.Item1;
            int skipped      = result.Item2;
            int textMatches  = result.Item3;
            if (appliedWorkspace is not null)
            {
                applied += appliedWorkspace.Applied;
                skipped += appliedWorkspace.Skipped;
                textMatches += appliedWorkspace.Applied;
            }

            _word.SaveDocument(_activeDoc);
            _backup.LogAction($"Применено {applied}, пропущено {skipped}, совпадений {textMatches}");
            // Обновляем applied_count в истории (SaveHistory вызывается ДО Apply)
            if (_lastHistoryId > 0)
                try { _history.UpdateApplied(_lastHistoryId, applied); } catch { }

            // Понятное сообщение пользователю
            if (applied == 0)
            {
                AddSystemMessage("⚠ Ни одна операция не была применена.");
                askAgentToContinue = true;
            }
            else if (skipped > 0)
            {
                var details = WordInteropService.LastApplyError;
                if (appliedWorkspace is not null && !string.IsNullOrWhiteSpace(appliedWorkspace.Details))
                    details = string.IsNullOrWhiteSpace(details)
                        ? appliedWorkspace.Details
                        : details + "\n" + appliedWorkspace.Details;
                if (!string.IsNullOrEmpty(details))
                    _backup.LogAction("Детали ошибок:\n" + details);

                AddSystemMessage($"⚠ Применено {applied}, пропущено {skipped}. Передаю результат AI, чтобы он продолжил точнее.");
                askAgentToContinue = true;
            }
            else if (textMatches == 0 && skipped == 0)
            {
                // Операции set_text_* — текст не найден
                AddSystemMessage($"⚠ Применено {applied} операций, но текст не найден в документе (0 вхождений). Проверьте написание.");
                askAgentToContinue = true;
            }
            else
            {
                var summary = $"✓ Применено: {applied} операций · {textMatches} изменений в тексте";
                AddSystemMessage(summary);
            }

            if (askAgentToContinue)
            {
                var details = WordInteropService.LastApplyError;
                if (appliedWorkspace is not null && !string.IsNullOrWhiteSpace(appliedWorkspace.Details))
                    details = string.IsNullOrWhiteSpace(details)
                        ? appliedWorkspace.Details
                        : details + "\n" + appliedWorkspace.Details;
                applyReport = BuildApplyReport(applied, skipped, textMatches, details);
            }
        }
        catch (Exception ex)
        {
            AddSystemMessage($"✗ Ошибка: {ex.Message}");
            askAgentToContinue = true;
            applyReport = $"Исключение при применении операций: {ex.GetType().Name}: {ex.Message}";
        }
        finally
        {
            // Всегда очищаем очередь — даже при ошибке, чтобы не применять дважды
            ClearPendingChanges();
        }

        if (askAgentToContinue)
            await ContinueAgentAfterApplyFailureAsync(applyReport, opList);
    }

    private async Task ContinueAgentAfterApplyFailureAsync(string applyReport, List<ChangeOperation> failedOps)
    {
        if (_activeDoc is null || _lastAgentHistory.Count == 0 || string.IsNullOrWhiteSpace(_lastUserCommand))
            return;

        _cts      = new CancellationTokenSource();
        var ct    = _cts.Token;
        IsBusy    = true;
        CanCancel = true;
        Progress  = 0;

        List<ParagraphInfo>? paragraphs = null;
        var doc     = _activeDoc;
        var allOps  = new List<ChangeOperation>();
        var history = _lastAgentHistory.ToList();

        async Task<List<ParagraphInfo>> EnsureParagraphsAsync()
        {
            if (paragraphs is not null) return paragraphs;
            SetProgress(20, "Кайра документти окуп жатам...");
            paragraphs = await Task.Run(() => (List<ParagraphInfo>)_word.ReadParagraphs(doc), ct);
            var tokens = DocumentParser.EstimateTokens(string.Join(" ", paragraphs.Select(p => p.Text)));
            var pages = paragraphs.Count == 0 ? 0 : paragraphs.Max(p => p.PageNumber);
            DocStats = $"{paragraphs.Count} абз · {pages} стр · ~{tokens / 1000}K токенов";
            return paragraphs;
        }

        try
        {
            AddSystemMessage("↻ Операция дала ошибку/0 результат. Передаю контекст AI, чтобы он продолжил.");
            SetProgress(5, "AI'га результатты берип жатам...");

            var failedJson = JsonSerializer.Serialize(failedOps, new JsonSerializerOptions { WriteIndented = true });
            history.Add(("user",
                $"{_word.GetWordContext(doc)}\n" +
                $"APPLY_RESULT:\n{applyReport}\n\n" +
                $"FAILED_OR_PARTIAL_OPERATIONS_JSON:\n{failedJson}\n\n" +
                $"CHAT_TRANSCRIPT:\n{BuildRecentChatTranscript()}\n\n" +
                $"Исходная задача клиента: {_lastUserCommand}\n" +
                "Продолжи с этого места. Не повторяй удачные действия. Если нужна информация из документа — вызови tool. " +
                "Если можешь исправить сразу — верни новые operations и done:true."));

            const int maxSteps = 6;
            for (int step = 0; step < maxSteps; step++)
            {
                ct.ThrowIfCancellationRequested();
                SetProgress(10 + (double)step / maxSteps * 85,
                    step == 0 ? "AI ойлонуп жатат..." : $"Кайра аракет {step + 1}...");

                var aiStep = await _ai.AgentStepAsync(history, ct);
                var assistantStepContent = aiStep.RawJson.Length > 0
                    ? aiStep.RawJson
                    : JsonSerializer.Serialize(aiStep);

                if (!string.IsNullOrEmpty(aiStep.Message))
                    AddAiMessage(aiStep.Message);

                if (aiStep.Done)
                {
                    history.Add(("assistant", assistantStepContent));
                    SetProgress(100, string.Empty);
                    allOps = aiStep.Operations ?? [];
                    if (allOps.Count > 0)
                    {
                        SetPendingChanges(allOps);
                        AddSystemMessage("✓ AI оңдолгон операцияларды даярдады. Чаттын ичинен операцияларды текшерип тастыктаңыз.");
                    }
                    break;
                }

                if (!string.IsNullOrEmpty(aiStep.Tool))
                {
                    var paras      = await EnsureParagraphsAsync();
                    var toolResult = ExecuteTool(aiStep, paras);
                    history.Add(("assistant", assistantStepContent));
                    history.Add(("user",
                        $"[{aiStep.Tool}] результат:\n{toolResult}\n\n" +
                        $"───\nИсходная задача: {_lastUserCommand}\n" +
                        $"Предыдущая ошибка применения:\n{applyReport}\n" +
                        "Теперь сгенерируй исправленные operations и верни done:true."));
                }
                else
                {
                    break;
                }
            }
        }
        catch (OperationCanceledException)
        {
            SetProgress(0, string.Empty);
        }
        catch (Exception ex)
        {
            AddSystemMessage($"✗ AI retry: {ex.Message}");
            SetProgress(0, string.Empty);
        }
        finally
        {
            IsBusy    = false;
            CanCancel = false;
            _cts?.Dispose();
            _cts = null;
        }

        _lastAgentHistory = history.ToList();
        if (allOps.Count > 0)
            _lastHistoryId = SaveHistory(_lastUserCommand + " [retry]", allOps);
    }

    private static string BuildApplyReport(int applied, int skipped, int textMatches, string details)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"applied={applied}");
        sb.AppendLine($"skipped={skipped}");
        sb.AppendLine($"textMatches={textMatches}");
        if (!string.IsNullOrWhiteSpace(details))
        {
            sb.AppendLine("details:");
            sb.AppendLine(details);
        }
        return sb.ToString();
    }

    private string BuildRecentChatTranscript(int maxMessages = 20)
    {
        var sb = new StringBuilder();
        foreach (var msg in ChatMessages.TakeLast(maxMessages))
            sb.AppendLine($"{msg.Role}: {msg.Text}");
        return sb.ToString();
    }

    // ──────────────────── Настройки ────────────────────

    private void LoadSettings()
    {
        var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
        if (File.Exists(path))
        {
            try
            {
                var loaded = System.Text.Json.JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(path));
                if (loaded is not null)
                {
                    _settings          = loaded;
                    AutoAnalyzeEnabled = loaded.AutoAnalyzeOnConnect;
                }
            }
            catch { }
        }
        ConfigureAI();
    }

    private void ConfigureAI()
    {
        var key        = CredentialService.Load(_settings.Provider.ToString()) ?? string.Empty;
        var promptsDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Prompts");
        var prompt     = ReadFile(Path.Combine(promptsDir, "system_prompt.txt"));
        _ai.Configure(_settings, key, prompt);
    }

    private static void SaveSettingsToFile(AppSettings s)
    {
        var path    = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
        var options = new System.Text.Json.JsonSerializerOptions { WriteIndented = true };
        File.WriteAllText(path, System.Text.Json.JsonSerializer.Serialize(s, options));
    }

    private static string ReadFile(string path) =>
        File.Exists(path) ? File.ReadAllText(path) : string.Empty;

    // ──────────────────── Helpers ────────────────────

    private void SetProgress(double value, string text)
    {
        Progress = value;
        if (!string.IsNullOrEmpty(text)) BusyText = text;
    }

    private const string QuickCheckAllPrompt =
        "Полная проверка и исправление оформления документа по стандарту ГОСТ/учебных работ:\n" +
        "• Шрифт основного текста: Times New Roman, 14pt\n" +
        "• Отступ первой строки: 1.25 см\n" +
        "• Межстрочный интервал: 1.5\n" +
        "• Выравнивание: по ширине (justify)\n\n" +
        "Используй modify_style для стиля 'Normal' (или 'Обычный') — это исправит всё тело документа за один шаг. " +
        "Заголовки (Heading 1, 2, 3) не трогай — у них свой стиль. " +
        "После применения modify_style сразу done:true.";

    private long SaveHistory(string command, List<ChangeOperation> ops)
    {
        try
        {
            return _history.AddAndGetId(new HistoryEntry
            {
                Timestamp       = DateTime.Now,
                DocumentName    = SelectedDocument ?? "—",
                Command         = command,
                Phase1Plan      = string.Empty,
                OperationsCount = ops.Count,
                AppliedCount    = 0,
                WasDirect       = false,
                OperationsJson  = JsonSerializer.Serialize(ops)
            });
        }
        catch { return 0; }
    }

    private void AddUserMessage(string text)
    {
        EnsureCurrentChatSession(text);
        AddChatMessage(text, MessageRole.User);
    }

    private void AddAiMessage(string text) => AddChatMessage(text, MessageRole.Assistant);

    private void AddSystemMessage(string text) => AddChatMessage(text, MessageRole.System);

    private void AddChatMessage(string text, MessageRole role)
    {
        ChatMessages.Add(new ChatMessage(text, role));

        if (_isLoadingChatSession || SelectedChatSession is null) return;

        try
        {
            _chat.AddMessage(SelectedChatSession.Id, role.ToString(), text);
            SelectedChatSession.UpdatedAt = DateTime.Now;
            ReloadChatSessionsKeepingSelection();
        }
        catch
        {
            // Чат в UI важнее: ошибки сохранения не должны ломать редактирование документа.
        }
    }

    private void ReloadChatSessionsKeepingSelection()
    {
        var selectedId = SelectedChatSession?.Id;
        var sessions = _chat.GetSessions();

        _isLoadingChatSession = true;
        try
        {
            ChatSessions = new ObservableCollection<ChatSessionInfo>(sessions);
            SelectedChatSession = selectedId.HasValue
                ? ChatSessions.FirstOrDefault(s => s.Id == selectedId.Value)
                : null;
        }
        finally
        {
            _isLoadingChatSession = false;
        }
    }
}

public record ChatMessage(string Text, MessageRole Role);
public enum MessageRole { User, Assistant, System }
