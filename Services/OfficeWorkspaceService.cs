using System.IO;
using System.Text;
using AIDocAssistant.Models;

namespace AIDocAssistant.Services;

public class OfficeWorkspaceService
{
    private readonly List<OfficeWorkspaceFile> _files = [];
    private long _nextId = 1;

    public IReadOnlyList<OfficeWorkspaceFile> Files => _files;

    public OfficeWorkspaceFile AddFile(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("Файл не найден.", path);

        var fullPath = Path.GetFullPath(path);
        var existing = _files.FirstOrDefault(f =>
            string.Equals(f.FullPath, fullPath, StringComparison.OrdinalIgnoreCase));
        if (existing is not null) return existing;

        var file = new OfficeWorkspaceFile
        {
            Id = _nextId++,
            Kind = DetectKind(fullPath),
            Name = Path.GetFileName(fullPath),
            FullPath = fullPath,
            OpenedAt = DateTime.Now
        };
        _files.Add(file);
        return file;
    }

    public OfficeWorkspaceFile? GetFile(long id) => _files.FirstOrDefault(f => f.Id == id);

    public string ListFiles()
    {
        if (_files.Count == 0)
            return "Office workspace пуст. Пользователь ещё не открыл файлы-контекст.";

        var sb = new StringBuilder();
        sb.AppendLine("OFFICE_WORKSPACE_FILES:");
        foreach (var file in _files)
            sb.AppendLine($"[{file.Id}] kind={file.Kind} name=\"{file.Name}\" path=\"{file.FullPath}\"");
        return sb.ToString();
    }

    public string ListFiles(IEnumerable<long> fileIds)
    {
        var ids = fileIds.ToHashSet();
        var files = _files.Where(f => ids.Contains(f.Id)).ToList();
        if (files.Count == 0)
            return "CHAT_RESOURCES empty. User has not selected files for this chat.";

        var sb = new StringBuilder();
        sb.AppendLine("CHAT_RESOURCES:");
        foreach (var file in files)
            sb.AppendLine($"[{file.Id}] kind={file.Kind} name=\"{file.Name}\" path=\"{file.FullPath}\"");
        return sb.ToString();
    }

    public OfficeDocumentSnapshot ReadFile(long id)
    {
        var file = GetFile(id) ?? throw new InvalidOperationException($"Файл workspace id={id} не найден.");
        return file.Kind switch
        {
            OfficeFileKind.Word       => ReadWord(file),
            OfficeFileKind.Excel      => ReadExcel(file),
            OfficeFileKind.PowerPoint => ReadPowerPoint(file),
            _ => new OfficeDocumentSnapshot
            {
                FileId = file.Id,
                Kind = file.Kind,
                Name = file.Name,
                FullPath = file.FullPath,
                Summary = "Неподдерживаемый тип файла."
            }
        };
    }

    public string CompareFiles(long sourceId, long targetId)
    {
        var source = ReadFile(sourceId);
        var target = ReadFile(targetId);

        var sb = new StringBuilder();
        sb.AppendLine("OFFICE_COMPARE_CONTEXT:");
        sb.AppendLine("SOURCE:");
        sb.AppendLine(source.Summary);
        sb.AppendLine();
        sb.AppendLine("TARGET:");
        sb.AppendLine(target.Summary);
        return sb.ToString();
    }

    public string SearchFiles(string query) => SearchFiles(query, _files.Select(f => f.Id));

    public string SearchFiles(string query, IEnumerable<long> fileIds)
    {
        if (string.IsNullOrWhiteSpace(query))
            return "Search query is empty.";
        var allowedIds = fileIds.ToHashSet();
        var files = _files.Where(f => allowedIds.Contains(f.Id)).ToList();
        if (files.Count == 0)
            return "Office workspace is empty.";

        var sb = new StringBuilder();
        sb.AppendLine($"OFFICE_WORKSPACE_SEARCH query=\"{query}\"");
        int total = 0;

        foreach (var file in files)
        {
            string summary;
            try { summary = ReadFile(file.Id).Summary; }
            catch (Exception ex)
            {
                sb.AppendLine($"[{file.Id}] {file.Name}: read error {ex.Message}");
                continue;
            }

            var lines = summary
                .Split(["\r\n", "\n"], StringSplitOptions.None)
                .Where(line => line.Contains(query, StringComparison.OrdinalIgnoreCase))
                .Take(20)
                .ToList();

            if (lines.Count == 0) continue;
            total += lines.Count;
            sb.AppendLine($"[{file.Id}] {file.Kind} {file.Name}: {lines.Count} matches");
            foreach (var line in lines)
                sb.AppendLine("  " + Trim(line, 260));
        }

        if (total == 0)
            sb.AppendLine("No matches.");

        return sb.ToString();
    }

    public static bool IsWorkspaceOperation(ChangeOperation op) => op.Type switch
    {
        "create_office_file" => true,
        "excel_set_cell" or "excel_replace_text" or "excel_add_sheet" or "excel_autofit" => true,
        "word_file_replace_text" or "word_file_append_text" => true,
        "powerpoint_replace_text" or "powerpoint_add_slide" or "powerpoint_set_shape_text" => true,
        _ => false
    };

    public OfficeApplyResult ApplyChanges(IEnumerable<ChangeOperation> operations)
    {
        var result = new OfficeApplyResult();
        var details = new StringBuilder();

        foreach (var op in operations)
        {
            try
            {
                var applied = ApplySingle(op);
                if (applied > 0) result.Applied += applied;
                else
                {
                    result.Skipped++;
                    details.AppendLine($"{op.Type}: skipped");
                }
            }
            catch (Exception ex)
            {
                result.Skipped++;
                details.AppendLine($"{op.Type}: {ex.GetType().Name}: {ex.Message}");
            }
        }

        result.Details = details.ToString();
        return result;
    }

    private int ApplySingle(ChangeOperation op)
    {
        if (op.Type == "create_office_file")
            return CreateOfficeFile(op);

        if (!op.FileId.HasValue)
            throw new InvalidOperationException("file_id is required for Office workspace operation.");

        var file = GetFile(op.FileId.Value)
            ?? throw new InvalidOperationException($"Workspace file id={op.FileId.Value} not found.");

        return op.Type switch
        {
            "excel_set_cell" or "excel_replace_text" or "excel_add_sheet" or "excel_autofit" =>
                ApplyExcel(file, op),
            "word_file_replace_text" or "word_file_append_text" =>
                ApplyWordFile(file, op),
            "powerpoint_replace_text" or "powerpoint_add_slide" or "powerpoint_set_shape_text" =>
                ApplyPowerPoint(file, op),
            _ => 0
        };
    }

    private int CreateOfficeFile(ChangeOperation op)
    {
        var kind = ParseKind(op.Kind);
        if (kind == OfficeFileKind.Unknown)
            throw new InvalidOperationException("kind must be Word, Excel, or PowerPoint.");

        var path = ResolveCreatePath(op, kind);
        Directory.CreateDirectory(Path.GetDirectoryName(path) ?? Environment.CurrentDirectory);

        switch (kind)
        {
            case OfficeFileKind.Word:
                CreateWordFile(path, op);
                break;
            case OfficeFileKind.Excel:
                CreateExcelFile(path, op);
                break;
            case OfficeFileKind.PowerPoint:
                CreatePowerPointFile(path, op);
                break;
        }

        AddFile(path).IsSelected = true;
        return 1;
    }

    private static OfficeFileKind ParseKind(string? kind) => kind?.Trim().ToLowerInvariant() switch
    {
        "word" or "doc" or "docx" => OfficeFileKind.Word,
        "excel" or "xls" or "xlsx" => OfficeFileKind.Excel,
        "powerpoint" or "ppt" or "pptx" => OfficeFileKind.PowerPoint,
        _ => OfficeFileKind.Unknown
    };

    private static string ResolveCreatePath(ChangeOperation op, OfficeFileKind kind)
    {
        if (!string.IsNullOrWhiteSpace(op.Path))
            return Path.GetFullPath(op.Path);

        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "AIDocAssistant");
        var ext = kind switch
        {
            OfficeFileKind.Word => ".docx",
            OfficeFileKind.Excel => ".xlsx",
            OfficeFileKind.PowerPoint => ".pptx",
            _ => ".dat"
        };
        var name = string.IsNullOrWhiteSpace(op.Name)
            ? $"new_{kind}_{DateTime.Now:yyyyMMdd_HHmmss}{ext}"
            : Path.ChangeExtension(SanitizeFileName(op.Name), ext);
        return Path.Combine(dir, name);
    }

    private static string SanitizeFileName(string name)
    {
        foreach (var ch in Path.GetInvalidFileNameChars())
            name = name.Replace(ch, '_');
        return name;
    }

    private static int ApplyExcel(OfficeWorkspaceFile file, ChangeOperation op)
    {
        if (file.Kind != OfficeFileKind.Excel)
            throw new InvalidOperationException($"File [{file.Id}] is {file.Kind}, not Excel.");

        dynamic? app = null;
        dynamic? wb = null;
        try
        {
            app = CreateCom("Excel.Application");
            app.Visible = false;
            app.DisplayAlerts = false;
            wb = app.Workbooks.Open(file.FullPath, ReadOnly: false);

            switch (op.Type)
            {
                case "excel_set_cell":
                {
                    if (string.IsNullOrWhiteSpace(op.Cell))
                        throw new InvalidOperationException("cell is required.");
                    dynamic ws = GetWorksheet(wb, op.Sheet);
                    ws.Range[op.Cell].Value2 = op.Value ?? op.Text ?? "";
                    wb.Save();
                    return 1;
                }

                case "excel_replace_text":
                {
                    var from = op.SearchText ?? op.From;
                    if (string.IsNullOrEmpty(from)) return 0;
                    int count = 0;
                    foreach (var ws in wb.Worksheets)
                    {
                        dynamic used = ws.UsedRange;
                        int rows = SafeInt(() => (int)used.Rows.Count);
                        int cols = SafeInt(() => (int)used.Columns.Count);
                        for (int r = 1; r <= rows; r++)
                        {
                            for (int c = 1; c <= cols; c++)
                            {
                                string current = Safe(() => Convert.ToString(used.Cells[r, c].Value2) ?? "", "");
                                if (current.Contains(from, StringComparison.OrdinalIgnoreCase))
                                {
                                    used.Cells[r, c].Value2 = current.Replace(from, op.To ?? "", StringComparison.OrdinalIgnoreCase);
                                    count++;
                                }
                            }
                        }
                    }
                    if (count > 0) wb.Save();
                    return count;
                }

                case "excel_add_sheet":
                {
                    dynamic ws = wb.Worksheets.Add();
                    if (!string.IsNullOrWhiteSpace(op.Sheet))
                        ws.Name = op.Sheet;
                    wb.Save();
                    return 1;
                }

                case "excel_autofit":
                {
                    dynamic ws = GetWorksheet(wb, op.Sheet);
                    ws.UsedRange.Columns.AutoFit();
                    wb.Save();
                    return 1;
                }
            }

            return 0;
        }
        finally
        {
            try { wb?.Close(false); } catch { }
            try { app?.Quit(); } catch { }
        }
    }

    private static void CreateWordFile(string path, ChangeOperation op)
    {
        dynamic? app = null;
        dynamic? doc = null;
        try
        {
            app = CreateCom("Word.Application");
            app.Visible = false;
            doc = app.Documents.Add();
            if (!string.IsNullOrWhiteSpace(op.Title))
                doc.Content.InsertAfter(op.Title + "\n\n");
            if (!string.IsNullOrWhiteSpace(op.Text ?? op.Value))
                doc.Content.InsertAfter(op.Text ?? op.Value);
            doc.SaveAs2(path);
        }
        finally
        {
            try { doc?.Close(false); } catch { }
            try { app?.Quit(false); } catch { }
        }
    }

    private static void CreateExcelFile(string path, ChangeOperation op)
    {
        dynamic? app = null;
        dynamic? wb = null;
        try
        {
            app = CreateCom("Excel.Application");
            app.Visible = false;
            app.DisplayAlerts = false;
            wb = app.Workbooks.Add();
            if (!string.IsNullOrWhiteSpace(op.Sheet))
                wb.Worksheets[1].Name = op.Sheet;
            if (!string.IsNullOrWhiteSpace(op.Value ?? op.Text))
                wb.Worksheets[1].Range["A1"].Value2 = op.Value ?? op.Text;
            wb.SaveAs(path);
        }
        finally
        {
            try { wb?.Close(false); } catch { }
            try { app?.Quit(); } catch { }
        }
    }

    private static void CreatePowerPointFile(string path, ChangeOperation op)
    {
        dynamic? app = null;
        dynamic? pres = null;
        try
        {
            app = CreateCom("PowerPoint.Application");
            pres = app.Presentations.Add(0);
            dynamic slide = pres.Slides.Add(1, 12); // ppLayoutBlank
            if (!string.IsNullOrWhiteSpace(op.Title))
            {
                dynamic titleBox = slide.Shapes.AddTextbox(1, 40, 30, 640, 60);
                titleBox.TextFrame.TextRange.Text = op.Title;
                titleBox.TextFrame.TextRange.Font.Size = 28;
            }
            if (!string.IsNullOrWhiteSpace(op.Text ?? op.Value))
            {
                dynamic bodyBox = slide.Shapes.AddTextbox(1, 40, 110, 640, 360);
                bodyBox.TextFrame.TextRange.Text = op.Text ?? op.Value;
                bodyBox.TextFrame.TextRange.Font.Size = 18;
            }
            pres.SaveAs(path);
        }
        finally
        {
            try { pres?.Close(); } catch { }
            try { app?.Quit(); } catch { }
        }
    }

    private static int ApplyWordFile(OfficeWorkspaceFile file, ChangeOperation op)
    {
        if (file.Kind != OfficeFileKind.Word)
            throw new InvalidOperationException($"File [{file.Id}] is {file.Kind}, not Word.");

        dynamic? app = null;
        dynamic? doc = null;
        try
        {
            app = CreateCom("Word.Application");
            app.Visible = false;
            doc = app.Documents.Open(FileName: file.FullPath, ReadOnly: false, AddToRecentFiles: false, Visible: false);

            switch (op.Type)
            {
                case "word_file_replace_text":
                {
                    var from = op.SearchText ?? op.From;
                    if (string.IsNullOrEmpty(from)) return 0;
                    dynamic find = doc.Content.Find;
                    find.ClearFormatting();
                    find.Text = from;
                    find.Replacement.ClearFormatting();
                    find.Replacement.Text = op.To ?? "";
                    find.MatchCase = op.MatchCase ?? false;
                    find.MatchWholeWord = op.WholeWord ?? false;
                    find.Execute(Replace: 2);
                    doc.Save();
                    return 1;
                }

                case "word_file_append_text":
                {
                    var text = op.Text ?? op.Value;
                    if (string.IsNullOrEmpty(text)) return 0;
                    doc.Content.InsertAfter("\n" + text);
                    doc.Save();
                    return 1;
                }
            }

            return 0;
        }
        finally
        {
            try { doc?.Close(false); } catch { }
            try { app?.Quit(false); } catch { }
        }
    }

    private static int ApplyPowerPoint(OfficeWorkspaceFile file, ChangeOperation op)
    {
        if (file.Kind != OfficeFileKind.PowerPoint)
            throw new InvalidOperationException($"File [{file.Id}] is {file.Kind}, not PowerPoint.");

        dynamic? app = null;
        dynamic? pres = null;
        try
        {
            app = CreateCom("PowerPoint.Application");
            pres = app.Presentations.Open(file.FullPath, ReadOnly: 0, Untitled: 0, WithWindow: 0);

            switch (op.Type)
            {
                case "powerpoint_replace_text":
                {
                    var from = op.SearchText ?? op.From;
                    if (string.IsNullOrEmpty(from)) return 0;
                    int count = 0;
                    foreach (var slide in pres.Slides)
                    {
                        foreach (var shape in slide.Shapes)
                        {
                            try
                            {
                                if ((bool)shape.HasTextFrame && (bool)shape.TextFrame.HasText)
                                {
                                    string current = (string)shape.TextFrame.TextRange.Text;
                                    if (current.Contains(from, StringComparison.OrdinalIgnoreCase))
                                    {
                                        shape.TextFrame.TextRange.Text = current.Replace(from, op.To ?? "", StringComparison.OrdinalIgnoreCase);
                                        count++;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    if (count > 0) pres.Save();
                    return count;
                }

                case "powerpoint_add_slide":
                {
                    int index = (int)pres.Slides.Count + 1;
                    dynamic slide = pres.Slides.Add(index, 12); // ppLayoutBlank
                    if (!string.IsNullOrWhiteSpace(op.Title))
                    {
                        dynamic titleBox = slide.Shapes.AddTextbox(1, 40, 30, 640, 60);
                        titleBox.TextFrame.TextRange.Text = op.Title;
                        titleBox.TextFrame.TextRange.Font.Size = 28;
                    }
                    var body = op.Text ?? op.Value;
                    if (!string.IsNullOrWhiteSpace(body))
                    {
                        dynamic bodyBox = slide.Shapes.AddTextbox(1, 40, 110, 640, 360);
                        bodyBox.TextFrame.TextRange.Text = body;
                        bodyBox.TextFrame.TextRange.Font.Size = 18;
                    }
                    pres.Save();
                    return 1;
                }

                case "powerpoint_set_shape_text":
                {
                    if (!op.SlideIndex.HasValue || !op.ShapeIndex.HasValue)
                        throw new InvalidOperationException("slide_index and shape_index are required.");
                    dynamic shape = pres.Slides[op.SlideIndex.Value].Shapes[op.ShapeIndex.Value];
                    shape.TextFrame.TextRange.Text = op.Text ?? op.Value ?? "";
                    pres.Save();
                    return 1;
                }
            }

            return 0;
        }
        finally
        {
            try { pres?.Close(); } catch { }
            try { app?.Quit(); } catch { }
        }
    }

    private static dynamic GetWorksheet(dynamic wb, string? sheet)
    {
        if (string.IsNullOrWhiteSpace(sheet))
            return wb.Worksheets[1];

        foreach (var ws in wb.Worksheets)
        {
            string name = Safe(() => (string)ws.Name, "");
            if (string.Equals(name, sheet, StringComparison.OrdinalIgnoreCase))
                return ws;
        }

        throw new InvalidOperationException($"Worksheet \"{sheet}\" not found.");
    }

    private static OfficeFileKind DetectKind(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext switch
        {
            ".doc" or ".docx" or ".docm" => OfficeFileKind.Word,
            ".xls" or ".xlsx" or ".xlsm" or ".csv" => OfficeFileKind.Excel,
            ".ppt" or ".pptx" or ".pptm" => OfficeFileKind.PowerPoint,
            _ => OfficeFileKind.Unknown
        };
    }

    private static OfficeDocumentSnapshot ReadWord(OfficeWorkspaceFile file)
    {
        dynamic? app = null;
        dynamic? doc = null;
        try
        {
            app = CreateCom("Word.Application");
            app.Visible = false;
            doc = app.Documents.Open(FileName: file.FullPath, ReadOnly: true, AddToRecentFiles: false, Visible: false);

            var sb = new StringBuilder();
            sb.AppendLine($"WORD_FILE id={file.Id} name=\"{file.Name}\"");
            sb.AppendLine($"pages={Safe(() => doc.ComputeStatistics(2).ToString(), "unknown")} paragraphs={Safe(() => doc.Paragraphs.Count.ToString(), "unknown")} tables={Safe(() => doc.Tables.Count.ToString(), "unknown")}");
            sb.AppendLine("paragraphs sample:");

            int count = 0;
            foreach (var para in doc.Paragraphs)
            {
                if (count >= 80) break;
                string text = Safe(() => ((string)para.Range.Text).TrimEnd('\r', '\n', '\a'), "");
                if (string.IsNullOrWhiteSpace(text) && count > 20) continue;
                string style = Safe(() => (string)para.Style.NameLocal, "");
                int page = SafeInt(() => (int)para.Range.Information[3]);
                bool inTable = SafeBool(() => (bool)para.Range.Information[12]);
                sb.AppendLine($"[{count}] page={page} table={inTable} style=\"{style}\" text=\"{Trim(text, 220)}\"");
                count++;
            }

            return Snapshot(file, sb.ToString());
        }
        finally
        {
            try { doc?.Close(false); } catch { }
            try { app?.Quit(false); } catch { }
        }
    }

    private static OfficeDocumentSnapshot ReadExcel(OfficeWorkspaceFile file)
    {
        dynamic? app = null;
        dynamic? wb = null;
        try
        {
            app = CreateCom("Excel.Application");
            app.Visible = false;
            app.DisplayAlerts = false;
            wb = app.Workbooks.Open(file.FullPath, ReadOnly: true);

            var sb = new StringBuilder();
            sb.AppendLine($"EXCEL_FILE id={file.Id} name=\"{file.Name}\" sheets={wb.Worksheets.Count}");

            foreach (var ws in wb.Worksheets)
            {
                string sheetName = Safe(() => (string)ws.Name, "Sheet");
                dynamic used = ws.UsedRange;
                int rows = SafeInt(() => (int)used.Rows.Count);
                int cols = SafeInt(() => (int)used.Columns.Count);
                sb.AppendLine($"sheet=\"{sheetName}\" usedRows={rows} usedCols={cols}");

                int maxRows = Math.Min(rows, 25);
                int maxCols = Math.Min(cols, 12);
                for (int r = 1; r <= maxRows; r++)
                {
                    var vals = new List<string>();
                    for (int c = 1; c <= maxCols; c++)
                    {
                        string value = Safe(() => Convert.ToString(used.Cells[r, c].Text) ?? "", "");
                        vals.Add(Trim(value.Replace("\t", " "), 80));
                    }
                    if (vals.Any(v => !string.IsNullOrWhiteSpace(v)))
                        sb.AppendLine($"  r{r}: " + string.Join(" | ", vals));
                }
            }

            return Snapshot(file, sb.ToString());
        }
        finally
        {
            try { wb?.Close(false); } catch { }
            try { app?.Quit(); } catch { }
        }
    }

    private static OfficeDocumentSnapshot ReadPowerPoint(OfficeWorkspaceFile file)
    {
        dynamic? app = null;
        dynamic? pres = null;
        try
        {
            app = CreateCom("PowerPoint.Application");
            pres = app.Presentations.Open(file.FullPath, ReadOnly: -1, Untitled: 0, WithWindow: 0);

            var sb = new StringBuilder();
            sb.AppendLine($"POWERPOINT_FILE id={file.Id} name=\"{file.Name}\" slides={pres.Slides.Count}");

            int slideLimit = Math.Min((int)pres.Slides.Count, 40);
            for (int i = 1; i <= slideLimit; i++)
            {
                dynamic slide = pres.Slides[i];
                sb.AppendLine($"slide={i} shapes={slide.Shapes.Count}");
                int shapeLimit = Math.Min((int)slide.Shapes.Count, 30);
                for (int s = 1; s <= shapeLimit; s++)
                {
                    dynamic shape = slide.Shapes[s];
                    string text = "";
                    try
                    {
                        if ((bool)shape.HasTextFrame && (bool)shape.TextFrame.HasText)
                            text = (string)shape.TextFrame.TextRange.Text;
                    }
                    catch { }

                    if (!string.IsNullOrWhiteSpace(text))
                        sb.AppendLine($"  shape={s} text=\"{Trim(text.Replace("\r", " ").Replace("\n", " "), 240)}\"");
                }
            }

            return Snapshot(file, sb.ToString());
        }
        finally
        {
            try { pres?.Close(); } catch { }
            try { app?.Quit(); } catch { }
        }
    }

    private static dynamic CreateCom(string progId)
    {
        var type = Type.GetTypeFromProgID(progId)
            ?? throw new InvalidOperationException($"{progId} недоступен. Проверьте установку Microsoft Office.");
        return Activator.CreateInstance(type)
            ?? throw new InvalidOperationException($"Не удалось запустить {progId}.");
    }

    private static OfficeDocumentSnapshot Snapshot(OfficeWorkspaceFile file, string summary) => new()
    {
        FileId = file.Id,
        Kind = file.Kind,
        Name = file.Name,
        FullPath = file.FullPath,
        Summary = summary
    };

    private static string Trim(string value, int max) =>
        value.Length <= max ? value : value[..max] + "...";

    private static string Safe(Func<string> read, string fallback)
    {
        try { return read(); }
        catch { return fallback; }
    }

    private static int SafeInt(Func<int> read)
    {
        try { return read(); }
        catch { return 0; }
    }

    private static bool SafeBool(Func<bool> read)
    {
        try { return read(); }
        catch { return false; }
    }
}
