using System.Text.Json.Serialization;

namespace AIDocAssistant.Models;

public class ChangeOperation
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = string.Empty;

    // set_font / абзац-операции
    [JsonPropertyName("paragraphIndex")]
    public int? ParagraphIndex { get; set; }

    // snake_case алиас — AI "paragraph_index" жиберет
    [JsonPropertyName("paragraph_index")]
    public int? ParagraphIndexSnake { get => ParagraphIndex; set => ParagraphIndex = value; }

    [JsonPropertyName("font")]
    public string? Font { get; set; }

    // snake_case алиас — AI "font_name" жиберет
    [JsonPropertyName("font_name")]
    public string? FontName { get => Font; set => Font = value; }

    [JsonPropertyName("size")]
    public float? Size { get; set; }

    // snake_case алиасы — AI "font_size" жиберет
    [JsonPropertyName("font_size")]
    public float? FontSize { get => Size; set => Size = value; }

    // set_alignment
    [JsonPropertyName("alignment")]
    public string? Alignment { get; set; }

    // set_spacing / set_line_spacing
    [JsonPropertyName("lineSpacing")]
    public float? LineSpacing { get; set; }

    // snake_case алиас — AI "line_spacing" жиберет
    [JsonPropertyName("line_spacing")]
    public float? LineSpacingSnake { get => LineSpacing; set => LineSpacing = value; }

    [JsonPropertyName("spaceBefore")]
    public float? SpaceBefore { get; set; }

    [JsonPropertyName("spaceAfter")]
    public float? SpaceAfter { get; set; }

    // set_indent
    [JsonPropertyName("firstLine")]
    public float? FirstLine { get; set; }

    [JsonPropertyName("left")]
    public float? Left { get; set; }

    // snake_case алиас — AI "left_indent_cm" жиберет
    [JsonPropertyName("left_indent_cm")]
    public float? LeftIndentCm { get => FirstLine; set => FirstLine = value; }

    [JsonPropertyName("right")]
    public float? Right { get; set; }

    // replace_text / set_text_color / set_text_bold …
    [JsonPropertyName("from")]
    public string? From { get; set; }

    [JsonPropertyName("to")]
    public string? To { get; set; }

    // snake_case алиас — AI "replace_with" жиберет
    [JsonPropertyName("replace_with")]
    public string? ReplaceWith { get => To; set => To = value; }

    // Для set_text_color / set_text_bold / set_text_italic / set_text_underline
    [JsonPropertyName("search_text")]
    public string? SearchText { get; set; }

    [JsonPropertyName("matchCase")]
    public bool? MatchCase { get; set; }

    [JsonPropertyName("wholeWord")]
    public bool? WholeWord { get; set; }

    // set_color / set_highlight
    [JsonPropertyName("color")]
    public string? Color { get; set; }

    // set_bold / set_italic / set_underline
    [JsonPropertyName("bold")]
    public bool? Bold { get; set; }

    [JsonPropertyName("italic")]
    public bool? Italic { get; set; }

    [JsonPropertyName("underline")]
    public bool? Underline { get; set; }

    // set_style / add_comment / insert_text
    [JsonPropertyName("style")]
    public string? Style { get; set; }

    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("position")]
    public string? Position { get; set; }

    // set_page_margin — поля страницы (в см)
    [JsonPropertyName("left_cm")]
    public float? LeftCm { get; set; }

    [JsonPropertyName("right_cm")]
    public float? RightCm { get; set; }

    [JsonPropertyName("top_cm")]
    public float? TopCm { get; set; }

    [JsonPropertyName("bottom_cm")]
    public float? BottomCm { get; set; }

    // set_all_indent / set_all_spacing / set_all_font — для ВСЕХ абзацев
    [JsonPropertyName("first_line_cm")]
    public float? FirstLineCm { get; set; }

    // replace_color — найти всё с цветом from_color, изменить на to_color
    [JsonPropertyName("from_color")]
    public string? FromColor { get; set; }

    [JsonPropertyName("to_color")]
    public string? ToColor { get; set; }

    [JsonPropertyName("file_id")]
    public long? FileId { get; set; }

    [JsonPropertyName("source_file_id")]
    public long? SourceFileId { get; set; }

    [JsonPropertyName("target_file_id")]
    public long? TargetFileId { get; set; }

    [JsonPropertyName("sheet")]
    public string? Sheet { get; set; }

    [JsonPropertyName("cell")]
    public string? Cell { get; set; }

    [JsonPropertyName("range")]
    public string? Range { get; set; }

    [JsonPropertyName("row")]
    public int? Row { get; set; }

    [JsonPropertyName("column")]
    public int? Column { get; set; }

    [JsonPropertyName("value")]
    public string? Value { get; set; }

    [JsonPropertyName("target_sheet")]
    public string? TargetSheet { get; set; }

    [JsonPropertyName("target_cell")]
    public string? TargetCell { get; set; }

    [JsonPropertyName("slide_index")]
    public int? SlideIndex { get; set; }

    [JsonPropertyName("shape_index")]
    public int? ShapeIndex { get; set; }

    [JsonPropertyName("title")]
    public string? Title { get; set; }

    [JsonPropertyName("kind")]
    public string? Kind { get; set; }

    [JsonPropertyName("name")]
    public string? Name { get; set; }

    [JsonPropertyName("path")]
    public string? Path { get; set; }

    // Отображение в Preview
    public string DisplayText => Type switch
    {
        "excel_set_cell"    => $"Excel file [{FileId}] {Sheet}!{Cell}: {Value ?? Text}",
        "excel_replace_text"=> $"Excel file [{FileId}]: replace \"{SearchText ?? From}\" -> \"{To}\"",
        "excel_add_sheet"   => $"Excel file [{FileId}]: add sheet \"{Sheet}\"",
        "excel_autofit"     => $"Excel file [{FileId}] {Sheet}: autofit columns",
        "word_file_replace_text" => $"Word file [{FileId}]: replace \"{SearchText ?? From}\" -> \"{To}\"",
        "word_file_append_text"  => $"Word file [{FileId}]: append text \"{Short(Text ?? Value)}\"",
        "powerpoint_replace_text" => $"PowerPoint file [{FileId}]: replace \"{SearchText ?? From}\" -> \"{To}\"",
        "powerpoint_add_slide"    => $"PowerPoint file [{FileId}]: add slide \"{Title ?? Short(Text)}\"",
        "powerpoint_set_shape_text" => $"PowerPoint file [{FileId}] slide {SlideIndex}, shape {ShapeIndex}: set text",
        "create_office_file" => $"Create {Kind} file: {Name ?? Path}",
        "set_font"          => $"Абзац {ParagraphIndex}: шрифт {Font} {Size}pt",
        "set_color"         => $"Абзац {ParagraphIndex}: цвет текста {Color}",
        "set_bold"          => $"Абзац {ParagraphIndex}: жирный {(Bold == true ? "вкл" : "выкл")}",
        "set_italic"        => $"Абзац {ParagraphIndex}: курсив {(Italic == true ? "вкл" : "выкл")}",
        "set_underline"     => $"Абзац {ParagraphIndex}: подчёркивание {(Underline == true ? "вкл" : "выкл")}",
        "set_text_color"    => $"Найти «{SearchText ?? From}» во всём документе → цвет {Color}",
        "set_text_bold"     => $"Найти «{SearchText ?? From}» во всём документе → жирный {(Bold == true ? "вкл" : "выкл")}",
        "set_text_italic"   => $"Найти «{SearchText ?? From}» во всём документе → курсив {(Italic == true ? "вкл" : "выкл")}",
        "set_text_underline"=> $"Найти «{SearchText ?? From}» во всём документе → подчёркивание",
        "set_text_size"     => $"Найти «{SearchText ?? From}» во всём документе → размер {Size}pt",
        "set_alignment"    => $"Абзац {ParagraphIndex}: выравнивание {Alignment}",
        "set_spacing"      => $"Абзац {ParagraphIndex}: межстрочный {LineSpacing}",
        "set_indent"       => $"Абзац {ParagraphIndex}: отступ первой строки {FirstLine} см",
        "replace_text"     => $"Замена: «{SearchText ?? From}» → «{To}»",
        "set_style"        => $"Абзац {ParagraphIndex}: стиль {Style}",
        "add_comment"      => $"Абзац {ParagraphIndex}: комментарий",
        "insert_text"      => ParagraphIndex == -1
            ? $"В конец документа: вставка текста"
            : $"Абзац {ParagraphIndex}: вставка текста ({Position})",
        "delete_paragraph"  => $"Абзац {ParagraphIndex}: удалить",
        "insert_toc" or "insert_table_of_contents" => "Вставить оглавление в начало",
        "update_toc"        => "Обновить оглавление",
        "insert_page_break_end" => "В конец документа: вставить новый лист",
        "insert_text_end"    => $"В конец документа: вставить текст «{Short(Text)}»",
        "modify_style"      => $"Стиль «{Style}»: шрифт={Font} {Size}pt выравн.={Alignment} отступ={FirstLineCm}см интервал={LineSpacing}",
        "set_page_margin"   => $"Поля страницы: лево={LeftCm}см право={RightCm}см верх={TopCm}см низ={BottomCm}см",
        "set_all_indent"    => $"Все абзацы: лев.отступ={LeftCm}см, первая строка={FirstLineCm}см",
        "set_all_spacing"   => $"Все абзацы: межстрочный={LineSpacing}",
        "set_all_font"      => $"Все абзацы: шрифт={Font} {Size}pt",
        "set_all_alignment" => $"Все абзацы: выравнивание {Alignment}",
        "replace_color"     => $"Заменить цвет {FromColor} → {ToColor} (весь документ)",
        _                   => $"{Type} (абзац {ParagraphIndex})"
    };

    public string ScriptText => Type switch
    {
        "excel_set_cell" =>
            $"Excel.Workbooks.Open(file[{FileId}]).Worksheets[{Q(Sheet)}].Range[{Q(Cell)}].Value2 = {Q(Value ?? Text)};",

        "excel_replace_text" =>
            $"Excel file[{FileId}]: ReplaceAll({Q(SearchText ?? From)}, {Q(To)});",

        "excel_add_sheet" =>
            $"Excel file[{FileId}]: Worksheets.Add().Name = {Q(Sheet)};",

        "excel_autofit" =>
            $"Excel file[{FileId}]: Worksheets[{Q(Sheet)}].UsedRange.Columns.AutoFit();",

        "word_file_replace_text" =>
            $"Word file[{FileId}]: Find.ReplaceAll({Q(SearchText ?? From)}, {Q(To)});",

        "word_file_append_text" =>
            $"Word file[{FileId}]: Content.InsertAfter({Q(Text ?? Value)});",

        "powerpoint_replace_text" =>
            $"PowerPoint file[{FileId}]: ReplaceAll({Q(SearchText ?? From)}, {Q(To)});",

        "powerpoint_add_slide" =>
            $"PowerPoint file[{FileId}]: Slides.Add(title={Q(Title)}, text={Q(Text ?? Value)});",

        "powerpoint_set_shape_text" =>
            $"PowerPoint file[{FileId}].Slides[{SlideIndex}].Shapes[{ShapeIndex}].TextFrame.TextRange.Text = {Q(Text ?? Value)};",

        "create_office_file" =>
            $"CreateOfficeFile(kind={Q(Kind)}, name={Q(Name)}, path={Q(Path)});",

        "set_font" or "set_font_name" =>
            $"doc.Paragraphs[{WordParagraphIndex}].Range.Font.Name = {Q(Font)};\n" +
            IfHas(Size, v => $"doc.Paragraphs[{WordParagraphIndex}].Range.Font.Size = {v};"),

        "set_font_size" =>
            $"doc.Paragraphs[{WordParagraphIndex}].Range.Font.Size = {Size?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "null"};",

        "set_color" =>
            $"doc.Paragraphs[{WordParagraphIndex}].Range.Font.Color = {Q(Color)};",

        "set_bold" =>
            $"doc.Paragraphs[{WordParagraphIndex}].Range.Font.Bold = {(Bold == true ? 1 : 0)};",

        "set_italic" =>
            $"doc.Paragraphs[{WordParagraphIndex}].Range.Font.Italic = {(Italic == true ? 1 : 0)};",

        "set_underline" =>
            $"doc.Paragraphs[{WordParagraphIndex}].Range.Font.Underline = {(Underline == true ? 1 : 0)};",

        "set_alignment" =>
            $"doc.Paragraphs[{WordParagraphIndex}].Alignment = {Q(Alignment)};",

        "set_spacing" or "set_line_spacing" =>
            $"ApplyLineSpacing(doc.Paragraphs[{WordParagraphIndex}], {Val(LineSpacing)});\n" +
            IfHas(SpaceBefore, v => $"doc.Paragraphs[{WordParagraphIndex}].SpaceBefore = {v};") +
            IfHas(SpaceAfter,  v => $"doc.Paragraphs[{WordParagraphIndex}].SpaceAfter = {v};"),

        "set_indent" =>
            $"doc.Paragraphs[{WordParagraphIndex}].FirstLineIndent = CmToPoints({Val(FirstLineCm ?? FirstLine)});\n" +
            IfHas(LeftCm ?? Left,   v => $"doc.Paragraphs[{WordParagraphIndex}].LeftIndent = CmToPoints({v});") +
            IfHas(RightCm ?? Right, v => $"doc.Paragraphs[{WordParagraphIndex}].RightIndent = CmToPoints({v});"),

        "replace_text" =>
            $"Find.ReplaceAll(searchText: {Q(SearchText ?? From)}, replaceWith: {Q(To)});",

        "set_text_color" =>
            $"Find.All(searchText: {Q(SearchText ?? From)}).ForEach(range => range.Font.Color = {Q(Color)});",

        "set_text_bold" =>
            $"Find.All(searchText: {Q(SearchText ?? From)}).ForEach(range => range.Font.Bold = {(Bold == true ? 1 : 0)});",

        "set_text_italic" =>
            $"Find.All(searchText: {Q(SearchText ?? From)}).ForEach(range => range.Font.Italic = {(Italic == true ? 1 : 0)});",

        "set_text_underline" =>
            $"Find.All(searchText: {Q(SearchText ?? From)}).ForEach(range => range.Font.Underline = {(Underline == true ? 1 : 0)});",

        "set_text_size" =>
            $"Find.All(searchText: {Q(SearchText ?? From)}).ForEach(range => range.Font.Size = {Val(Size)});",

        "set_style" =>
            $"doc.Paragraphs[{WordParagraphIndex}].Style = {Q(Style)};",

        "add_comment" =>
            $"doc.Comments.Add(doc.Paragraphs[{WordParagraphIndex}].Range, {Q(Text)});",

        "insert_text" =>
            ParagraphIndex == -1
                ? $"doc.Content.InsertAfter(\"\\n\" + {Q(Text)});"
                : Position == "before"
                    ? $"doc.Paragraphs[{WordParagraphIndex}].Range.InsertBefore({Q(Text)} + \"\\n\");"
                    : $"doc.Paragraphs[{WordParagraphIndex}].Range.InsertAfter(\"\\n\" + {Q(Text)});",

        "delete_paragraph" =>
            $"doc.Paragraphs[{WordParagraphIndex}].Range.Delete();",

        "insert_toc" or "insert_table_of_contents" =>
            "var range = doc.Paragraphs[1].Range;\nrange.Collapse(wdCollapseStart);\ndoc.TablesOfContents.Add(range, UseHeadingStyles: true, UpperHeadingLevel: 1, LowerHeadingLevel: 3);",

        "update_toc" =>
            "foreach (var toc in doc.TablesOfContents) toc.Update();",

        "insert_page_break_end" =>
            "var range = doc.Content;\nrange.Collapse(wdCollapseEnd);\nrange.InsertBreak(wdPageBreak);",

        "insert_text_end" =>
            $"doc.Content.InsertAfter(\"\\n\" + {Q(Text)});",

        "modify_style" =>
            $"var style = doc.Styles[{Q(Style)}];\n" +
            IfHas(Font, v => $"style.Font.Name = {Q(v)};") +
            IfHas(Size, v => $"style.Font.Size = {v};") +
            IfHas(Alignment, v => $"style.ParagraphFormat.Alignment = {Q(v)};") +
            IfHas(LineSpacing, v => $"ApplyLineSpacing(style.ParagraphFormat, {v});") +
            IfHas(FirstLineCm, v => $"style.ParagraphFormat.FirstLineIndent = CmToPoints({v});"),

        "replace_color" =>
            $"Find.ByFormat(fontColor: {Q(FromColor)}).ReplaceAll(replacementFontColor: {Q(ToColor)});",

        "set_page_margin" =>
            "var ps = doc.PageSetup;\n" +
            IfHas(LeftCm,   v => $"ps.LeftMargin = CmToPoints({v});") +
            IfHas(RightCm,  v => $"ps.RightMargin = CmToPoints({v});") +
            IfHas(TopCm,    v => $"ps.TopMargin = CmToPoints({v});") +
            IfHas(BottomCm, v => $"ps.BottomMargin = CmToPoints({v});"),

        "set_all_indent" =>
            "foreach (var p in doc.Paragraphs) {\n" +
            IfHas(LeftCm, v => $"  p.LeftIndent = CmToPoints({v});") +
            IfHas(FirstLineCm, v => $"  p.FirstLineIndent = CmToPoints({v});") +
            "}",

        "set_all_spacing" =>
            "foreach (var p in doc.Paragraphs) {\n" +
            IfHas(LineSpacing, v => $"  ApplyLineSpacing(p, {v});") +
            IfHas(SpaceBefore, v => $"  p.SpaceBefore = {v};") +
            IfHas(SpaceAfter, v => $"  p.SpaceAfter = {v};") +
            "}",

        "set_all_font" =>
            "foreach (var p in doc.Paragraphs) {\n" +
            IfHas(Font, v => $"  p.Range.Font.Name = {Q(v)};") +
            IfHas(Size, v => $"  p.Range.Font.Size = {v};") +
            "}",

        "set_all_alignment" =>
            $"foreach (var p in doc.Paragraphs) p.Alignment = {Q(Alignment)};",

        _ => $"// Неизвестная операция: {Type}\n// Приложение пропустит или применит её только если исполнитель WordInteropService знает этот type."
    };

    public string GroupKey => Type switch
    {
        "set_font" or "set_bold" or "set_italic" or "set_underline" => "Шрифт",
        "set_color"              => "Цвет абзаца",
        "set_text_color" or "set_text_bold" or "set_text_italic" or "set_text_underline" or "set_text_size"
                                 => "По тексту (весь документ)",
        "set_alignment"          => "Выравнивание",
        "set_spacing"            => "Межстрочный интервал",
        "set_indent"             => "Отступы",
        "replace_text"           => "Замена текста",
        "set_style"              => "Стили",
        "add_comment"            => "Комментарии",
        "insert_text"            => "Вставка текста",
        "delete_paragraph"       => "Удаление абзацев",
        "update_toc" or "insert_toc" or "insert_table_of_contents" => "Оглавление",
        "insert_page_break_end" or "insert_text_end" => "Конец документа",
        "excel_set_cell" or "excel_replace_text" or "excel_add_sheet" or "excel_autofit" => "Excel workspace",
        "word_file_replace_text" or "word_file_append_text" => "Word workspace",
        "powerpoint_replace_text" or "powerpoint_add_slide" or "powerpoint_set_shape_text" => "PowerPoint workspace",
        "create_office_file" => "Create Office file",
        _                        => "Прочее"
    };

    private int WordParagraphIndex => (ParagraphIndex ?? 0) + 1;

    private static string Q(string? value) =>
        value is null ? "null" : "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n") + "\"";

    private static string Val(float? value) =>
        value?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "null";

    private static string Short(string? value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        var text = value.Replace("\r", " ").Replace("\n", " ");
        return text.Length > 50 ? text[..50] + "..." : text;
    }

    private static string IfHas(float? value, Func<string, string> line) =>
        value.HasValue ? line(value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)) + "\n" : string.Empty;

    private static string IfHas(string? value, Func<string, string> line) =>
        string.IsNullOrEmpty(value) ? string.Empty : line(value) + "\n";
}
