using System.IO;
using System.Runtime.InteropServices;
using AIDocAssistant.Models;

namespace AIDocAssistant.Services;

// Word COM через dynamic — не зависит от версии PIA
public class WordInteropService : IDisposable
{
    private dynamic? _wordApp;
    private bool _disposed;

    public static string LastConnectError { get; private set; } = string.Empty;

    // --- P/Invoke ---

    [DllImport("oleaut32.dll", EntryPoint = "GetActiveObject", PreserveSig = true)]
    private static extern int GetActiveObjectNative(ref Guid rclsid, IntPtr pvReserved, out IntPtr ppunk);

    [DllImport("ole32.dll", EntryPoint = "CLSIDFromProgID", CharSet = CharSet.Unicode, PreserveSig = true)]
    private static extern int CLSIDFromProgIDNative(string lpszProgID, out Guid pclsid);

    // --- Подключение ---

    public bool TryConnect()
    {
        var log  = new System.Text.StringBuilder();
        var diag = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AIDocAssistant_diag.txt");

        log.AppendLine($"=== TryConnect {DateTime.Now} ===");
        try
        {
            // Шаг 1
            log.AppendLine("Step1: CLSIDFromProgID...");
            int hr = CLSIDFromProgIDNative("Word.Application", out Guid clsid);
            log.AppendLine($"  HR=0x{hr:X8}  CLSID={clsid}");
            if (hr != 0) { LastConnectError = $"CLSIDFromProgID HR=0x{hr:X8}"; return false; }

            // Шаг 2
            log.AppendLine("Step2: GetActiveObject...");
            hr = GetActiveObjectNative(ref clsid, IntPtr.Zero, out IntPtr pUnk);
            log.AppendLine($"  HR=0x{hr:X8}  pUnk={pUnk}");
            if (hr != 0) { LastConnectError = $"GetActiveObject HR=0x{hr:X8}"; return false; }

            // Шаг 3
            log.AppendLine("Step3: GetObjectForIUnknown...");
            var obj = Marshal.GetObjectForIUnknown(pUnk);
            Marshal.Release(pUnk);
            log.AppendLine($"  Type={obj?.GetType()?.FullName}");

            // Шаг 4: проверяем ActiveDocument
            log.AppendLine("Step4: ActiveDocument.Name...");
            _wordApp = obj;
            var docName = (string)_wordApp!.ActiveDocument.Name;
            log.AppendLine($"  ActiveDoc={docName}");

            LastConnectError = string.Empty;
            log.AppendLine("SUCCESS");
            return true;
        }
        catch (Exception ex)
        {
            log.AppendLine($"EXCEPTION: {ex.GetType().Name}: {ex.Message}");
            LastConnectError = ex.Message;
            _wordApp = null;
            return false;
        }
        finally
        {
            File.AppendAllText(diag, log.ToString() + "\n");
        }
    }

    public bool IsConnected => _wordApp is not null;

    public dynamic OpenDocumentFromFile(string path, bool visible = true)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("Документ не найден.", path);

        EnsureWordForFileOpen(visible);
        _wordApp!.Visible = visible;

        dynamic doc = _wordApp.Documents.Open(
            FileName: path,
            ReadOnly: false,
            AddToRecentFiles: false,
            Visible: visible);

        doc.Activate();
        LastConnectError = string.Empty;
        return doc;
    }

    public List<string> GetOpenDocumentNames()
    {
        EnsureConnected();
        var names = new List<string>();
        foreach (var doc in _wordApp!.Documents)
            names.Add((string)doc.Name);
        return names;
    }

    public dynamic? GetActiveDocument()
    {
        EnsureConnected();
        try { return _wordApp!.ActiveDocument; }
        catch { return null; }
    }

    public dynamic? GetDocumentByName(string name)
    {
        EnsureConnected();
        foreach (var doc in _wordApp!.Documents)
            if ((string)doc.Name == name) return doc;
        return null;
    }

    private void EnsureWordForFileOpen(bool visible)
    {
        if (_wordApp is not null) return;

        if (TryAttachRunningWord())
        {
            try { _wordApp!.Visible = visible; } catch { }
            return;
        }

        var type = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException("Microsoft Word не установлен или ProgID Word.Application недоступен.");

        _wordApp = Activator.CreateInstance(type)
            ?? throw new InvalidOperationException("Не удалось запустить Microsoft Word.");

        _wordApp.Visible = visible;
    }

    private bool TryAttachRunningWord()
    {
        try
        {
            int hr = CLSIDFromProgIDNative("Word.Application", out Guid clsid);
            if (hr != 0) return false;

            hr = GetActiveObjectNative(ref clsid, IntPtr.Zero, out IntPtr pUnk);
            if (hr != 0) return false;

            _wordApp = Marshal.GetObjectForIUnknown(pUnk);
            Marshal.Release(pUnk);
            return _wordApp is not null;
        }
        catch
        {
            _wordApp = null;
            return false;
        }
    }

    public string GetWordContext(dynamic? doc)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("WORD_CONTEXT:");
        sb.AppendLine($"- WordVersion: {Safe(() => (string)_wordApp!.Version, "unknown")}");
        sb.AppendLine($"- WordBuild: {Safe(() => (string)_wordApp!.Build, "unknown")}");
        sb.AppendLine($"- UI language id: {Safe(() => _wordApp!.LanguageSettings.LanguageID[2].ToString(), "unknown")}");

        if (doc is null)
        {
            sb.AppendLine("- Document: none");
            return sb.ToString();
        }

        sb.AppendLine($"- DocumentName: {Safe(() => (string)doc.Name, "unknown")}");
        sb.AppendLine($"- FullName: {Safe(() => (string)doc.FullName, "unknown")}");
        sb.AppendLine($"- ReadOnly: {Safe(() => doc.ReadOnly.ToString(), "unknown")}");
        sb.AppendLine($"- ProtectionType: {Safe(() => doc.ProtectionType.ToString(), "unknown")}");
        sb.AppendLine($"- Paragraphs: {Safe(() => doc.Paragraphs.Count.ToString(), "unknown")}");
        sb.AppendLine($"- Pages: {Safe(() => doc.ComputeStatistics(2).ToString(), "unknown")}"); // wdStatisticPages
        sb.AppendLine($"- Tables: {Safe(() => doc.Tables.Count.ToString(), "unknown")}");
        sb.AppendLine($"- Sections: {Safe(() => doc.Sections.Count.ToString(), "unknown")}");
        sb.AppendLine($"- TablesOfContents: {Safe(() => doc.TablesOfContents.Count.ToString(), "unknown")}");
        sb.AppendLine($"- HasSelection: {Safe(() => _wordApp!.Selection is not null ? "true" : "false", "unknown")}");
        return sb.ToString();
    }

    // --- Чтение ---

    public List<ParagraphInfo> ReadParagraphs(dynamic doc)
    {
        var result      = new List<ParagraphInfo>();
        int index       = 0;
        // Кэш цветов стилей — чтобы не вызывать COM для каждого абзаца
        var styleColors = new Dictionary<string, int>();

        foreach (var para in doc.Paragraphs)
        {
            string rawText;
            try { rawText = (string)para.Range.Text; }
            catch { rawText = string.Empty; }

            var text = rawText.TrimEnd('\r', '\n', '\a');
            bool isInTable = SafeBool(() => (bool)para.Range.Information[12]); // wdWithInTable
            bool isEmpty = string.IsNullOrWhiteSpace(text);
            bool isTableEndMarker = isInTable && isEmpty && rawText.Contains('\a');

            string styleName;
            try { styleName = (string)para.Style.NameLocal; }
            catch { styleName = "Normal"; }

            // 1. Явный цвет символа
            int rawColor;
            try { rawColor = (int)para.Range.Font.Color; } catch { rawColor = -16777216; }

            // 2. Если auto — пробуем прочитать цвет из стиля (Heading 1 → золотой и т.п.)
            if (rawColor == unchecked((int)0xFF000000) || rawColor == -16777216)
            {
                if (!styleColors.TryGetValue(styleName, out int sc))
                {
                    try { sc = (int)doc.Styles[styleName].Font.Color; }
                    catch { sc = -16777216; }
                    styleColors[styleName] = sc;
                }
                if (sc != -16777216 && sc != unchecked((int)0xFF000000))
                    rawColor = sc;
            }

            var info = new ParagraphInfo
            {
                Index           = index++,
                PageNumber      = SafeInt(() => (int)para.Range.Information[3]), // wdActiveEndPageNumber
                SectionNumber   = SafeInt(() => (int)para.Range.Information[2]), // wdActiveEndSectionNumber
                Text            = text,
                Style           = styleName,
                Font            = (string)para.Range.Font.Name,
                Size            = (float)para.Range.Font.Size,
                LineSpacing     = (float)para.LineSpacing,
                LineSpacingRule = SafeInt(() => (int)para.LineSpacingRule),
                Alignment       = AlignmentName((int)para.Alignment),
                FirstLineIndent = CmFromPoints((float)para.FirstLineIndent),
                LeftIndent      = CmFromPoints((float)para.LeftIndent),
                RightIndent     = CmFromPoints((float)para.RightIndent),
                SpaceBefore     = (float)para.SpaceBefore,
                SpaceAfter      = (float)para.SpaceAfter,
                Color           = ColorName(rawColor),
                IsEmpty         = isEmpty,
                IsInTable       = isInTable,
                IsTableEndMarker= isTableEndMarker
            };
            result.Add(info);
        }
        return result;
    }

    // --- Применение операций ---

    public static string LastApplyError { get; private set; } = string.Empty;

    // returns (appliedOps, skippedOps, textMatchesTotal)
    public (int applied, int skipped, int textMatches) ApplyChanges(dynamic doc, IEnumerable<ChangeOperation> ops)
    {
        int applied = 0, skipped = 0, textMatches = 0;
        var errors  = new System.Text.StringBuilder();

        var paragraphs = new List<dynamic>();
        foreach (var p in doc.Paragraphs) paragraphs.Add(p);

        foreach (var op in ops)
        {
            try
            {
                int matches = ApplySingle(doc, paragraphs, op);
                if (matches > 0)
                {
                    applied++;
                    textMatches += matches;
                }
                else
                {
                    skipped++;
                    errors.AppendLine($"  [{op.Type} п{op.ParagraphIndex}]: операция не изменила документ");
                }
            }
            catch (Exception ex)
            {
                skipped++;
                errors.AppendLine($"  [{op.Type} п{op.ParagraphIndex}]: {ex.Message}");
            }
        }

        LastApplyError = errors.ToString();
        return (applied, skipped, textMatches);
    }

    // returns count of text matches (for set_text_* / replace_text ops); 1 for paragraph ops
    private static int ApplySingle(dynamic doc, List<dynamic> paragraphs, ChangeOperation op)
    {
        dynamic? para = op.ParagraphIndex.HasValue && op.ParagraphIndex.Value < paragraphs.Count
            ? paragraphs[op.ParagraphIndex.Value] : null;

        switch (op.Type)
        {
            // ── По абзацу ──────────────────────────────────────────────
            case "set_font":
            case "set_font_name":
                if (para is null) return 0;
                if (!string.IsNullOrEmpty(op.Font)) para.Range.Font.Name = op.Font;
                if (op.Size.HasValue) para.Range.Font.Size = op.Size.Value;
                return 1;

            case "set_font_size":
                if (para is null) return 0;
                if (op.Size.HasValue) para.Range.Font.Size = op.Size.Value;
                return 1;

            case "set_color":
                if (para is null) return 0;
                para.Range.Font.Color = ParseColor(op.Color);
                return 1;

            case "set_bold":
                if (para is null) return 0;
                para.Range.Font.Bold = op.Bold == true ? 1 : 0;
                return 1;

            case "set_italic":
                if (para is null) return 0;
                para.Range.Font.Italic = op.Italic == true ? 1 : 0;
                return 1;

            case "set_underline":
                if (para is null) return 0;
                para.Range.Font.Underline = op.Underline == true ? 1 : 0;
                return 1;

            case "set_alignment":
                if (para is null) return 0;
                para.Alignment = ParseAlignment(op.Alignment);
                return 1;

            case "set_spacing":
            case "set_line_spacing":
                if (para is null) return 0;
                if (op.LineSpacing.HasValue) ApplyLineSpacing(para, op.LineSpacing.Value);
                if (op.SpaceBefore.HasValue) para.SpaceBefore = op.SpaceBefore.Value;
                if (op.SpaceAfter.HasValue)  para.SpaceAfter  = op.SpaceAfter.Value;
                return 1;

            case "set_indent":
                if (para is null) return 0;
                // FirstLine / left_indent_cm (legacy) / first_line_cm — все три варианта
                var fi = op.FirstLineCm ?? op.FirstLine;
                if (fi.HasValue) para.FirstLineIndent = PointsFromCm(fi.Value);
                // Left / left_cm
                var li = op.LeftCm ?? op.Left;
                if (li.HasValue) para.LeftIndent = PointsFromCm(li.Value);
                // Right / right_cm
                var ri = op.RightCm ?? op.Right;
                if (ri.HasValue) para.RightIndent = PointsFromCm(ri.Value);
                return 1;

            // ── По тексту (весь документ) ──────────────────────────────
            case "replace_text":
            {
                var needle = op.SearchText ?? op.From;
                if (string.IsNullOrEmpty(needle)) return 0;
                var f = doc.Content.Find;
                f.ClearFormatting();
                f.Text = needle;
                f.Replacement.ClearFormatting();
                f.Replacement.Text  = op.To ?? string.Empty;
                f.MatchCase         = op.MatchCase ?? false;
                f.MatchWholeWord    = op.WholeWord ?? false;
                f.Execute(Replace: 2); // wdReplaceAll
                return 1;
            }

            case "set_text_color":
            {
                var needle = op.SearchText ?? op.From;
                if (string.IsNullOrEmpty(needle)) return 0;
                return FindAndApply(doc, needle, TextFmt.Color, ParseColor(op.Color));
            }

            case "set_text_bold":
            {
                var needle = op.SearchText ?? op.From;
                if (string.IsNullOrEmpty(needle)) return 0;
                return FindAndApply(doc, needle, TextFmt.Bold, op.Bold == true ? 1 : 0);
            }

            case "set_text_italic":
            {
                var needle = op.SearchText ?? op.From;
                if (string.IsNullOrEmpty(needle)) return 0;
                return FindAndApply(doc, needle, TextFmt.Italic, op.Italic == true ? 1 : 0);
            }

            case "set_text_underline":
            {
                var needle = op.SearchText ?? op.From;
                if (string.IsNullOrEmpty(needle)) return 0;
                return FindAndApply(doc, needle, TextFmt.Underline, op.Underline == true ? 1 : 0);
            }

            case "set_text_size":
            {
                var needle = op.SearchText ?? op.From;
                if (string.IsNullOrEmpty(needle) || !op.Size.HasValue) return 0;
                return FindAndApply(doc, needle, TextFmt.Size, op.Size.Value);
            }

            // ── Прочие ─────────────────────────────────────────────────
            case "set_style":
                if (para is null || string.IsNullOrEmpty(op.Style)) return 0;
                para.Style = op.Style;
                return 1;

            case "add_comment":
                if (para is null || string.IsNullOrEmpty(op.Text)) return 0;
                doc.Comments.Add(para.Range, op.Text);
                return 1;

            case "insert_text":
                if (string.IsNullOrEmpty(op.Text)) return 0;
                if (op.ParagraphIndex == -1)
                {
                    doc.Content.InsertAfter("\n" + op.Text);
                    return 1;
                }
                if (para is null) return 0;
                if (op.Position == "before") para.Range.InsertBefore(op.Text + "\n");
                else                          para.Range.InsertAfter("\n" + op.Text);
                return 1;

            case "delete_paragraph":
                if (para is null) return 0;
                para.Range.Delete();
                return 1;

            case "update_toc":
                foreach (var toc in doc.TablesOfContents) toc.Update();
                return 1;

            case "insert_toc":
            case "insert_table_of_contents": // AI sometimes uses this name
            {
                dynamic firstParaRng = doc.Paragraphs[1].Range;
                firstParaRng.Collapse(1); // wdCollapseStart = 1
                // UseHeadingStyles=true, UpperLevel=1, LowerLevel=3
                doc.TablesOfContents.Add(firstParaRng, true, 1, 3);
                return 1;
            }

            case "insert_page_break_end":
            {
                dynamic rng = doc.Content;
                rng.Collapse(0); // wdCollapseEnd
                rng.InsertBreak(7); // wdPageBreak
                return 1;
            }

            case "insert_text_end":
            {
                if (string.IsNullOrEmpty(op.Text)) return 0;
                doc.Content.InsertAfter("\n" + op.Text);
                return 1;
            }

            // ── Изменить форматирование стиля (влияет на все абзацы с этим стилем) ──
            case "modify_style":
            {
                if (string.IsNullOrEmpty(op.Style)) return 0;
                dynamic style = doc.Styles[op.Style];
                if (!string.IsNullOrEmpty(op.Font))  style.Font.Name = op.Font;
                if (op.Size.HasValue)                 style.Font.Size = op.Size.Value;
                if (!string.IsNullOrEmpty(op.Alignment))
                    style.ParagraphFormat.Alignment = ParseAlignment(op.Alignment);
                if (op.LineSpacing.HasValue)
                    ApplyLineSpacing(style.ParagraphFormat, op.LineSpacing.Value);
                if (op.FirstLineCm.HasValue)
                    style.ParagraphFormat.FirstLineIndent = PointsFromCm(op.FirstLineCm.Value);
                if (op.LeftCm.HasValue)
                    style.ParagraphFormat.LeftIndent = PointsFromCm(op.LeftCm.Value);
                if (op.SpaceBefore.HasValue)
                    style.ParagraphFormat.SpaceBefore = op.SpaceBefore.Value;
                if (op.SpaceAfter.HasValue)
                    style.ParagraphFormat.SpaceAfter  = op.SpaceAfter.Value;
                if (op.Bold.HasValue)   style.Font.Bold      = op.Bold   == true ? 1 : 0;
                if (op.Italic.HasValue) style.Font.Italic    = op.Italic == true ? 1 : 0;
                return 1;
            }

            // ── Замена цвета по всему документу (Word Find-by-Format) ──
            case "replace_color":
            {
                if (string.IsNullOrEmpty(op.FromColor)) return 0;
                int fromClr = ParseColor(op.FromColor);
                int toClr   = ParseColor(op.ToColor);
                dynamic rng = doc.Content;
                rng.Find.ClearFormatting();
                rng.Find.Text            = "";
                rng.Find.Format          = true; // ← поиск по форматированию, не только по тексту
                rng.Find.Font.Color      = fromClr;
                rng.Find.Wrap            = 1; // wdFindContinue
                rng.Find.Replacement.ClearFormatting();
                rng.Find.Replacement.Text            = "";
                rng.Find.Replacement.Font.Color      = toClr;
                rng.Find.Execute(Replace: 2); // wdReplaceAll
                // Сброс флага Format после операции
                rng.Find.ClearFormatting();
                rng.Find.Format = false;
                return 1;
            }

            // ── Поля страницы ───────────────────────────────────────────
            case "set_page_margin":
            {
                var ps = doc.PageSetup;
                if (op.LeftCm.HasValue)   ps.LeftMargin   = PointsFromCm(op.LeftCm.Value);
                if (op.RightCm.HasValue)  ps.RightMargin  = PointsFromCm(op.RightCm.Value);
                if (op.TopCm.HasValue)    ps.TopMargin    = PointsFromCm(op.TopCm.Value);
                if (op.BottomCm.HasValue) ps.BottomMargin = PointsFromCm(op.BottomCm.Value);
                return 1;
            }

            // ── Применить ко ВСЕМ абзацам ───────────────────────────────
            case "set_all_indent":
            {
                foreach (var p in paragraphs)
                {
                    if (op.LeftCm.HasValue)      p.LeftIndent      = PointsFromCm(op.LeftCm.Value);
                    if (op.FirstLineCm.HasValue)  p.FirstLineIndent = PointsFromCm(op.FirstLineCm.Value);
                }
                return paragraphs.Count;
            }

            case "set_all_spacing":
            {
                foreach (var p in paragraphs)
                {
                    if (op.LineSpacing.HasValue)  ApplyLineSpacing(p, op.LineSpacing.Value);
                    if (op.SpaceBefore.HasValue)  p.SpaceBefore = op.SpaceBefore.Value;
                    if (op.SpaceAfter.HasValue)   p.SpaceAfter  = op.SpaceAfter.Value;
                }
                return paragraphs.Count;
            }

            case "set_all_font":
            {
                foreach (var p in paragraphs)
                {
                    if (!string.IsNullOrEmpty(op.Font)) p.Range.Font.Name = op.Font;
                    if (op.Size.HasValue)               p.Range.Font.Size = op.Size.Value;
                }
                return paragraphs.Count;
            }

            case "set_all_alignment":
            {
                int align = ParseAlignment(op.Alignment);
                foreach (var p in paragraphs)
                    p.Alignment = align;
                return paragraphs.Count;
            }

            default:
                return 0;
        }
    }

    // --- Backup ---

    public string CreateBackup(dynamic doc)
    {
        var dir   = Path.GetDirectoryName((string)doc.FullName) ?? string.Empty;
        var name  = Path.GetFileNameWithoutExtension((string)doc.Name);
        var stamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm");
        var path  = Path.Combine(dir, $"{name}_backup_{stamp}.docx");
        doc.SaveCopyAs(path); // Копия — оригинал остаётся активным
        return path;
    }

    public void SaveDocument(dynamic doc) => doc.Save();

    // --- Helpers ---
    // WdParagraphAlignment: Left=0, Right=2, Center=1, Justify=3, Mixed=-1

    private static string AlignmentName(int a) => a switch
    {
        0 => "Left", 1 => "Center", 2 => "Right", 3 => "Justify", _ => "Left"
    };

    private static int ParseAlignment(string? name) => name?.ToLower() switch
    {
        "center"  => 1,
        "right"   => 2,
        "justify" => 3,
        _         => 0
    };

    private static void ApplyLineSpacing(dynamic paragraphOrFormat, float lineSpacing)
    {
        if (Math.Abs(lineSpacing - 1.0f) < 0.01f)
        {
            paragraphOrFormat.LineSpacingRule = 0; // wdLineSpaceSingle
            return;
        }

        if (Math.Abs(lineSpacing - 1.5f) < 0.01f)
        {
            paragraphOrFormat.LineSpacingRule = 1; // wdLineSpace1pt5
            return;
        }

        if (Math.Abs(lineSpacing - 2.0f) < 0.01f)
        {
            paragraphOrFormat.LineSpacingRule = 2; // wdLineSpaceDouble
            return;
        }

        paragraphOrFormat.LineSpacingRule = 5; // wdLineSpaceMultiple
        paragraphOrFormat.LineSpacing = lineSpacing * 12f; // Word's LinesToPoints approximation
    }

    private enum TextFmt { Color, Bold, Italic, Underline, Size }

    // Находит все вхождения needle, применяет форматирование, возвращает кол-во совпадений
    private static int FindAndApply(dynamic doc, string needle, TextFmt fmt, object value)
    {
        dynamic rng = doc.Content;
        rng.Find.ClearFormatting();
        rng.Find.Text      = needle;
        rng.Find.MatchCase = false;
        rng.Find.Wrap      = 1; // wdFindContinue

        int count = 0;
        while ((bool)rng.Find.Execute() && count < 500)
        {
            switch (fmt)
            {
                case TextFmt.Color:     rng.Font.Color     = (int)value;   break;
                case TextFmt.Bold:      rng.Font.Bold      = (int)value;   break;
                case TextFmt.Italic:    rng.Font.Italic    = (int)value;   break;
                case TextFmt.Underline: rng.Font.Underline = (int)value;   break;
                case TextFmt.Size:      rng.Font.Size      = (float)value; break;
            }
            rng.Collapse(0); // wdCollapseEnd — сдвигаем после найденного, иначе повтор
            count++;
        }
        return count;
    }

    // WdColor values (BGR формат Word)
    private static int ParseColor(string? color) => color?.ToLower() switch
    {
        "auto"     => -16777216, // wdColorAutomatic — наследует от стиля
        "red"      => 255,       // wdColorRed
        "blue"     => 16711680,  // wdColorBlue
        "green"    => 32768,     // wdColorGreen
        "black"    => 0,         // wdColorBlack
        "white"    => 16777215,  // wdColorWhite
        "yellow"   => 65535,     // wdColorYellow
        "orange"   => 26367,     // wdColorOrange
        "purple"   => 8388736,   // wdColorViolet
        "cyan"     => 16776960,  // wdColorCyan
        "magenta"  => 16711935,  // wdColorPink
        "darkred"  => 128,       // wdColorDarkRed
        "darkblue" => 8388608,   // wdColorDarkBlue
        "darkgreen"=> 32896,     // wdColorDarkGreen
        "gray"     => 8421504,   // wdColorGray50
        _          => 0
    };

    // Word color int → читаемое имя. Формат: R + G*256 + B*65536
    private static string ColorName(int c) => c switch
    {
        -16777216  => "auto",   // wdColorAutomatic = unchecked((int)0xFF000000)
        0          => "black",
        255        => "red",
        65535      => "yellow",
        16711680   => "blue",
        32768      => "green",
        16777215   => "white",
        26367      => "orange",
        8388736    => "purple",
        8421504    => "gray",
        16711935   => "magenta",
        16776960   => "cyan",
        128        => "darkred",
        8388608    => "darkblue",
        // Общие gold/amber цвета заголовков Word
        49407      => "gold",      // RGB(255,192,0)
        52479      => "gold",      // RGB(255,204,0)
        33023      => "amber",     // RGB(255,128,0)
        // Неизвестный → RGB hex для диагностики
        _ => $"#{c & 0xFF:X2}{(c >> 8) & 0xFF:X2}{(c >> 16) & 0xFF:X2}"
    };

    private static float CmFromPoints(float pt) => pt / 28.3465f;
    private static float PointsFromCm(float cm) => cm * 28.3465f;

    private void EnsureConnected()
    {
        if (_wordApp is null)
            throw new InvalidOperationException("Нет подключения к Word.");
    }

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

    public void Dispose()
    {
        if (_disposed) return;
        if (_wordApp is not null)
        {
            Marshal.ReleaseComObject(_wordApp);
            _wordApp = null;
        }
        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
