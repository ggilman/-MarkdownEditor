using Microsoft.Win32;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;

namespace MarkdownEditor;

/// <summary>
/// Main window providing markdown editing, live preview, and export functionality.
/// </summary>
public partial class MainWindow : Window
{
    // UI Constants
    private const string DefaultPreviewTitle = "Markdown Preview";
    private const int MaxFileNameLength = 200;

    // Preview Styling Constants
    private const string PreviewFontFamily = "Segoe UI,Arial,sans-serif";
    private const string CodeFontFamily = "Consolas";
    private const string CodeBackgroundColor = "#f3f3f3";
    private const string CodeHeaderBackground = "#333";
    private const string CodeHeaderForeground = "#f2f2f2";
    private const string CodeCopyButtonBackground = "#4a4a4a";
    private const string CodeCopyButtonHoverBackground = "#5a5a5a";
    private const string CodeBorderColor = "#d8d8d8";
    private const string BlockquoteBorderColor = "#999";
    private const string BlockquoteTextColor = "#444";
    private const string TableBorderColor = "#ccc";
    private const string BodyTextColor = "#222";
    private const string CodeBorderGrayColor = "#666";

    // KaTeX CDN Constants
    private const string KatexVersion = "0.16.9";
    private const string KatexCssUrl = "https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css";
    private const string KatexCssIntegrity = "sha384-n8MVd4RsNIU0tAv4ct0nTaAbDJwPJzDEaqSD1odI+WdtXRGWt2kTvGFasHpSy3SV";
    private const string KatexJsUrl = "https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.js";
    private const string KatexJsIntegrity = "sha384-XjKyOOlGwcjNTAIQHIpgOno0Hl1YQqzUOEleOLALmuqehneUG+vnGctmUb0ZY0l8";
    private const string KatexAutoRenderUrl = "https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/contrib/auto-render.min.js";
    private const string KatexAutoRenderIntegrity = "sha384-+VBxd3r6XgURycqtZ117nYw44OOcIax56Z4dCRWbxyPt0Koah1uHoK0o4+/RRE05";

    // Copy button text
    private const string CopyButtonText = "Copy";
    private const string CopyButtonCopiedText = "Copied";
    private const string CopyButtonFailedText = "Failed";
    private const int CopyButtonFeedbackDelayMs = 1000;

    // Markdown formatting markers
    private const string BoldMarker = "**";
    private const string ItalicMarker = "*";
    private const string StrikethroughMarker = "~~";
    private const string InlineCodeMarker = "`";
    private const string CodeBlockMarker = "```";
    private const string BulletListPrefix = "- ";
    private const string QuotePrefix = "> ";
    private const string HorizontalRuleMarker = "---";

    // Placeholder text
    private const string BoldPlaceholder = "bold text";
    private const string ItalicPlaceholder = "italic text";
    private const string StrikethroughPlaceholder = "strikethrough text";
    private const string CodePlaceholder = "code";
    private const string LinkTextPlaceholder = "link text";
    private const string LinkUrlPlaceholder = "https://example.com";
    private const string ImageAltPlaceholder = "alt text";
    private const string ImageUrlPlaceholder = "https://example.com/image.png";
    private const string ListItemPlaceholder = "list item";
    private const string TextPlaceholder = "text";

    // Status messages
    private const string StatusReady = "Ready";
    private const string StatusEditorClear = "Editor already clear";
    private const string StatusNewDocument = "New document";
    private const string StatusSaveFailed = "Save failed";
    private const string StatusImportFailed = "Import failed";
    private const string StatusExportFailed = "Export failed";
    private const string StatusBoldApplied = "Applied bold formatting";
    private const string StatusItalicApplied = "Applied italic formatting";
    private const string StatusStrikethroughApplied = "Applied strikethrough formatting";
    private const string StatusInlineCodeApplied = "Applied inline code formatting";
    private const string StatusCodeBlockInserted = "Inserted code block";
    private const string StatusHorizontalRuleInserted = "Inserted horizontal rule";
    private const string StatusBulletListInserted = "Inserted bullet list markers";
    private const string StatusNumberedListInserted = "Inserted numbered list markers";
    private const string StatusQuoteInserted = "Inserted quote markers";
    private const string StatusLinkInserted = "Inserted markdown link";
    private const string StatusImageInserted = "Inserted markdown image";

    // File filter constants
    private const string MarkdownFileFilter = "Markdown files (*.md)|*.md|All files (*.*)|*.*";
    private const string WordFileFilter = "Word document (*.docx)|*.docx";
    private const string PdfFileFilter = "PDF document (*.pdf)|*.pdf";
    private const string MarkdownExtension = ".md";
    private const string WordExtension = ".docx";
    private const string PdfExtension = ".pdf";
    private const string DefaultDocumentName = "document";

    private const string DefaultMarkdown = "# Markdown Editor Example\n\nUse this sample to verify preview and Word export formatting.\n\n## Heading Levels\n\n### H3 Example\n\n#### H4 Example\n\n##### H5 Example\n\n###### H6 Example\n\n## Text Styles\n\nThis line uses **bold**, *italic*, ~~strikethrough~~, and `inline code`.\n\nA markdown link: [Markdown Guide](https://www.markdownguide.org)\n\nAn image placeholder: ![Sample image](https://example.com/image.png)\n\n---\n\n## Lists\n\n- Bullet one\n- Bullet two\n  - Nested bullet\n\n1. First step\n2. Second step\n   1. Nested numbered step\n\n- [x] Task complete\n- [ ] Task pending\n\n## Quote\n\n> Preview pane is read-only and this quote should be styled in Word export.\n\n## Code Block\n\n```csharp\nusing System;\n\nConsole.WriteLine(\"Hello markdown\");\n```\n\n## Table\n\n| Feature | Preview | Word Export |\n| --- | --- | --- |\n| Headings | Yes | Yes |\n| Tables | Yes | Yes |\n| Checkboxes | Yes | Yes |\n| Code Language Label | Yes | Yes |";
    private int _selectedHeadingLevel = 2;
    private bool _isWebView2Initialized = false;
    private string? _currentFileName = null;

    public MainWindow()
    {
        InitializeComponent();
        EditorTextBox.Text = DefaultMarkdown;
        UpdateViewVisibility();
        InitializeWebView2Async();
    }

    /// <summary>
    /// Initializes the WebView2 control for preview rendering.
    /// Shows an error message if WebView2 Runtime is not installed.
    /// </summary>
    private async void InitializeWebView2Async()
    {
        try
        {
            await PreviewBrowser.EnsureCoreWebView2Async();
            _isWebView2Initialized = true;

            PreviewBrowser.CoreWebView2.NewWindowRequested += CoreWebView2_NewWindowRequested;

            RenderPreview();
        }
        catch (Exception ex)
        {
            string htmlDocument = "<!doctype html><html><head><meta charset=\"utf-8\"></head><body>" +
                "<div style=\"padding: 20px; font-family: sans-serif; color: #333;\">" +
                "<h2>Preview Unavailable</h2>" +
                "<p>The formatted preview requires the Microsoft Edge WebView2 Runtime, which doesn't seem to be installed on this computer.</p>" +
                "<p><a href=\"https://developer.microsoft.com/en-us/microsoft-edge/webview2/\">Download WebView2 Runtime</a></p>" +
                "<p style=\"color: #666; font-size: 12px; margin-top: 20px;\">Error details: " + System.Net.WebUtility.HtmlEncode(ex.Message) + "</p>" +
                "</div></body></html>";
            
            try { PreviewBrowser.NavigateToString(htmlDocument); } catch { /* ignore if completely broken */ }
            _isWebView2Initialized = false;
        }
    }

    private void EditorTextBox_TextChanged(object? sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        RenderPreview();
    }

    private void EditorTextBox_PreviewKeyDown(object? sender, KeyEventArgs e)
    {
        ModifierKeys modifiers = Keyboard.Modifiers;

        if (modifiers == ModifierKeys.Control)
        {
            switch (e.Key)
            {
                case Key.B:
                    BoldButton_Click(sender, new RoutedEventArgs());
                    e.Handled = true;
                    return;
                case Key.I:
                    ItalicButton_Click(sender, new RoutedEventArgs());
                    e.Handled = true;
                    return;
                case Key.K:
                    LinkButton_Click(sender, new RoutedEventArgs());
                    e.Handled = true;
                    return;
                case Key.D2:
                    InsertHeading(2);
                    e.Handled = true;
                    return;
            }
        }

        if (modifiers == (ModifierKeys.Control | ModifierKeys.Shift))
        {
            switch (e.Key)
            {
                case Key.C:
                    CodeBlockButton_Click(sender, new RoutedEventArgs());
                    e.Handled = true;
                    return;
                case Key.D8:
                    BulletListButton_Click(sender, new RoutedEventArgs());
                    e.Handled = true;
                    return;
                case Key.D7:
                    NumberedListButton_Click(sender, new RoutedEventArgs());
                    e.Handled = true;
                    return;
                case Key.Q:
                    QuoteButton_Click(sender, new RoutedEventArgs());
                    e.Handled = true;
                    return;
                case Key.K:
                    InlineCodeButton_Click(sender, new RoutedEventArgs());
                    e.Handled = true;
                    return;
            }
        }
    }

    /// <summary>
    /// Renders the current markdown text as HTML in the preview pane.
    /// Includes KaTeX for LaTeX math rendering and custom code block styling.
    /// </summary>
    private void RenderPreview()
    {
        string markdown = EditorTextBox.Text;
        string htmlBody = EnhanceCodeBlocksHtml(MarkdownConverter.ToHtml(markdown));

        string documentTitle = string.IsNullOrEmpty(_currentFileName) ? DefaultPreviewTitle : _currentFileName;
        string htmlDocument = "<!doctype html><html><head><meta charset=\"utf-8\">" +
            $"<title>{System.Net.WebUtility.HtmlEncode(documentTitle)}</title>" +
            $"<link rel=\"stylesheet\" href=\"{KatexCssUrl}\" integrity=\"{KatexCssIntegrity}\" crossorigin=\"anonymous\">" +
            $"<script defer src=\"{KatexJsUrl}\" integrity=\"{KatexJsIntegrity}\" crossorigin=\"anonymous\"></script>" +
            $"<script defer src=\"{KatexAutoRenderUrl}\" integrity=\"{KatexAutoRenderIntegrity}\" crossorigin=\"anonymous\"></script>" +
            "<style>" +
            "@page{margin:0.5in;size:auto;}" +
            "@media print{" +
            "@page{margin:0.5in;}" +
            "body{margin:0;padding:12px;}" +
            ".code-copy-btn{display:none !important;}" +
            "}" +
            $"body{{font-family:{PreviewFontFamily};padding:12px;line-height:1.5;color:{BodyTextColor};}}" +
            $".code-block{{margin:1em 0;border-radius:6px;overflow:hidden;border:1px solid {CodeBorderColor};background:{CodeBackgroundColor};}}" +
            $".code-header{{position:relative;background:{CodeHeaderBackground};color:{CodeHeaderForeground};padding:7px 72px 7px 10px;min-height:30px;" +
            "box-sizing:border-box;white-space:normal;}" +
            $".code-lang-badge{{display:inline-block;font-size:12px;line-height:1.2;color:{CodeHeaderForeground};}}" +
            $".code-block pre{{margin:0;background:{CodeBackgroundColor};overflow:auto;}}" +
            ".code-block pre > code{display:block;padding:10px;}" +
            $".code-copy-btn{{position:absolute;right:10px;top:6px;border:1px solid {CodeBorderGrayColor};background:{CodeCopyButtonBackground};color:#fff;" +
            "border-radius:4px;padding:2px 8px;font-size:12px;line-height:1.2;cursor:pointer;opacity:0;visibility:hidden;" +
            "transition:opacity .15s ease;}" +
            ".code-block:hover .code-copy-btn,.code-block:focus-within .code-copy-btn{opacity:1;visibility:visible;}" +
            $".code-copy-btn:hover{{background:{CodeCopyButtonHoverBackground};}}" +
            $"code{{background:{CodeBackgroundColor};padding:2px 4px;border-radius:4px;}}" +
            $"blockquote{{border-left:4px solid {BlockquoteBorderColor};padding-left:10px;color:{BlockquoteTextColor};margin-left:0;}}" +
            $"table{{border-collapse:collapse;}}th,td{{border:1px solid {TableBorderColor};padding:6px;}}" +
            "</style>" +
            "<script>(function(){" +
            "function getText(el){if(!el){return '';}if(typeof el.textContent==='string'){return el.textContent;}if(typeof el.innerText==='string'){return el.innerText;}return '';}" +
            "function copyText(text){" +
            "var ta=document.createElement('textarea');" +
            "ta.value=text;document.body.appendChild(ta);ta.select();" +
            "var ok=false;" +
            "try{ok=document.execCommand('copy');}catch(e){ok=false;}" +
            "document.body.removeChild(ta);" +
            "return ok;" +
            "}" +
            "window.copyCode=function(button){" +
            "var header=button.parentNode;var block=header?header.parentNode:null;" +
            "var pres=block?block.getElementsByTagName('pre'):null;" +
            "var pre=pres&&pres.length?pres[0]:null;" +
            "var codes=pre?pre.getElementsByTagName('code'):null;" +
            "var text=getText(codes&&codes.length?codes[0]:pre);" +
            $"if(copyText(text)){{button.innerText='{CopyButtonCopiedText}';setTimeout(function(){{button.innerText='{CopyButtonText}';}},{CopyButtonFeedbackDelayMs});}}" +
            $"else{{button.innerText='{CopyButtonFailedText}';setTimeout(function(){{button.innerText='{CopyButtonText}';}},{CopyButtonFeedbackDelayMs});}}" +
            "return false;" +
            "};" +
            "document.addEventListener('DOMContentLoaded',function(){" +
            "if(window.renderMathInElement){" +
            "renderMathInElement(document.body,{" +
            "delimiters:[" +
            "{left:'$$',right:'$$',display:true}," +
            "{left:'$',right:'$',display:false}," +
            "{left:'\\\\[',right:'\\\\]',display:true}," +
            "{left:'\\\\(',right:'\\\\)',display:false}" +
            "]," +
            "throwOnError:false" +
            "});" +
            "}" +
            "});" +
            "})();</script></head><body>" +
            htmlBody + "</body></html>";

        if (_isWebView2Initialized)
        {
            PreviewBrowser.NavigateToString(htmlDocument);
        }
    }

    /// <summary>
    /// Enhances code blocks in HTML with language badges and copy buttons.
    /// </summary>
    private static string EnhanceCodeBlocksHtml(string htmlBody)
    {
        if (string.IsNullOrEmpty(htmlBody))
        {
            return htmlBody;
        }

        return Regex.Replace(
            htmlBody,
            "<pre><code(?<attrs>[^>]*)>(?<content>[\\s\\S]*?)</code></pre>",
            match =>
            {
                string attrs = match.Groups["attrs"].Value;
                string content = match.Groups["content"].Value;
                string language = ExtractLanguage(attrs);
                string badgeText = string.IsNullOrWhiteSpace(language) ? CodePlaceholder : language;
                string badge = $"<div class=\"code-header\"><span class=\"code-lang-badge\">{System.Net.WebUtility.HtmlEncode(badgeText)}</span>" +
                               $"<button type=\"button\" class=\"code-copy-btn\" onclick=\"return copyCode(this);\">{CopyButtonText}</button></div>";

                  return "<div class=\"code-block\">" +
                       badge +
                      $"<pre><code{attrs}>{content}</code></pre>" +
                      "</div>";
            },
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
        );
    }

    private static string ExtractLanguage(string codeAttributes)
    {
        if (string.IsNullOrWhiteSpace(codeAttributes))
        {
            return string.Empty;
        }

        Match classMatch = Regex.Match(codeAttributes, "class=\\\"(?<class>[^\\\"]+)\\\"", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        if (!classMatch.Success)
        {
            return string.Empty;
        }

        string classList = classMatch.Groups["class"].Value;
        foreach (string className in classList.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (className.StartsWith("language-", StringComparison.OrdinalIgnoreCase) && className.Length > 9)
            {
                return className.Substring(9);
            }
        }

        return string.Empty;
    }

    private void NewButton_Click(object? sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(EditorTextBox.Text))
        {
            EditorTextBox.Focus();
            StatusText.Text = StatusEditorClear;
            return;
        }

        MessageBoxResult result = MessageBox.Show(
            this,
            "Clear the current markdown and start a new document?",
            "New Document",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        EditorTextBox.Clear();
        EditorTextBox.Focus();
        _currentFileName = null;
        StatusText.Text = StatusNewDocument;
    }

    private void SaveButton_Click(object? sender, RoutedEventArgs e)
    {
        SaveFileDialog dialog = new()
        {
            Filter = MarkdownFileFilter,
            DefaultExt = MarkdownExtension,
            FileName = $"{DefaultDocumentName}{MarkdownExtension}",
            Title = "Save Markdown"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            File.WriteAllText(dialog.FileName, EditorTextBox.Text ?? string.Empty);
            StatusText.Text = $"Saved: {Path.GetFileName(dialog.FileName)}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Save failed", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusText.Text = StatusSaveFailed;
        }
    }

    private void ImportButton_Click(object? sender, RoutedEventArgs e)
    {
        OpenFileDialog dialog = new()
        {
            Filter = MarkdownFileFilter,
            Title = "Import Markdown"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            EditorTextBox.Text = File.ReadAllText(dialog.FileName);
            _currentFileName = Path.GetFileName(dialog.FileName);
            StatusText.Text = $"Imported: {_currentFileName}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Import failed", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusText.Text = StatusImportFailed;
        }
    }

    private void ExportButton_Click(object? sender, RoutedEventArgs e)
    {
        string suggestedFileName = GetSuggestedFileName(EditorTextBox.Text) ?? $"{DefaultDocumentName}{WordExtension}";

        SaveFileDialog dialog = new()
        {
            Filter = WordFileFilter,
            DefaultExt = WordExtension,
            FileName = suggestedFileName,
            Title = "Export to Word"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            bool autoNumberHeadings = AutoNumberHeadingsMenuItem.IsChecked;
            bool useTitleStyle = UseTitleStyleMenuItem.IsChecked;
            DocxExporter.ExportFromMarkdown(EditorTextBox.Text, dialog.FileName, autoNumberHeadings, useTitleStyle);
            StatusText.Text = $"Exported: {Path.GetFileName(dialog.FileName)}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Export failed", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusText.Text = StatusExportFailed;
        }
    }

    private async void ExportPdfButton_Click(object? sender, RoutedEventArgs e)
    {
        if (!_isWebView2Initialized)
        {
            MessageBox.Show(this, "Preview is not ready. Please wait a moment and try again.", "PDF Export", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        string suggestedFileName = GetSuggestedFileName(EditorTextBox.Text);
        if (suggestedFileName != null && suggestedFileName.EndsWith(WordExtension, StringComparison.OrdinalIgnoreCase))
        {
            suggestedFileName = Path.ChangeExtension(suggestedFileName, PdfExtension);
        }
        else
        {
            suggestedFileName = $"{DefaultDocumentName}{PdfExtension}";
        }

        SaveFileDialog dialog = new()
        {
            Filter = PdfFileFilter,
            DefaultExt = PdfExtension,
            FileName = suggestedFileName,
            Title = "Export to PDF"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            var settings = PreviewBrowser.CoreWebView2.Environment.CreatePrintSettings();
            settings.ShouldPrintHeaderAndFooter = false;
            settings.MarginTop = 0.5;
            settings.MarginBottom = 0.5;
            settings.MarginLeft = 0.5;
            settings.MarginRight = 0.5;
            settings.ScaleFactor = 1.0;

            bool success = await PreviewBrowser.CoreWebView2.PrintToPdfAsync(dialog.FileName, settings);

            if (success)
            {
                StatusText.Text = $"Exported: {Path.GetFileName(dialog.FileName)}";
            }
            else
            {
                MessageBox.Show(this, "PDF export was cancelled or failed.", "Export failed", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "PDF export failed";
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Export failed", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusText.Text = "PDF export failed";
        }
    }

    /// <summary>
    /// Extracts the first heading from markdown to suggest as a filename.
    /// </summary>
    /// <param name="markdown">The markdown text to parse.</param>
    /// <returns>Sanitized filename with .docx extension, or null if no heading found.</returns>
    private static string? GetSuggestedFileName(string markdown)
    {
        if (string.IsNullOrWhiteSpace(markdown))
        {
            return null;
        }

        string[] lines = markdown.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

        foreach (string line in lines)
        {
            string trimmedLine = line.TrimStart();
            if (trimmedLine.StartsWith("#"))
            {
                string headerText = trimmedLine.TrimStart('#').Trim();
                if (!string.IsNullOrWhiteSpace(headerText))
                {
                    string sanitized = SanitizeFileName(headerText);
                    if (!string.IsNullOrWhiteSpace(sanitized))
                    {
                        return sanitized + WordExtension;
                    }
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Sanitizes a string for use as a filename by removing markdown formatting
    /// and replacing invalid characters with underscores.
    /// </summary>
    private static string SanitizeFileName(string fileName)
    {
        string cleaned = fileName;

        cleaned = Regex.Replace(cleaned, @"\*\*([^\*]+)\*\*", "$1");
        cleaned = Regex.Replace(cleaned, @"\*([^\*]+)\*", "$1");
        cleaned = Regex.Replace(cleaned, @"__([^_]+)__", "$1");
        cleaned = Regex.Replace(cleaned, @"_([^_]+)_", "$1");
        cleaned = Regex.Replace(cleaned, @"~~([^~]+)~~", "$1");
        cleaned = Regex.Replace(cleaned, @"`([^`]+)`", "$1");

        char[] invalidChars = Path.GetInvalidFileNameChars();
        string sanitized = string.Concat(cleaned.Select(c => invalidChars.Contains(c) ? '_' : c));

        sanitized = Regex.Replace(sanitized, @"_+", "_");
        sanitized = sanitized.Trim('_', ' ');

        if (sanitized.Length > MaxFileNameLength)
        {
            sanitized = sanitized.Substring(0, MaxFileNameLength).TrimEnd('_', ' ');
        }

        return sanitized;
    }

    private void HeadingApplyButton_Click(object? sender, RoutedEventArgs e)
    {
        InsertHeading(_selectedHeadingLevel);
    }

    private void HeadingComboBox_SelectionChanged(object? sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (sender is ComboBox comboBox && comboBox.SelectedIndex >= 0)
        {
            _selectedHeadingLevel = comboBox.SelectedIndex + 1;
        }
    }

    private void ExitMenuItem_Click(object? sender, RoutedEventArgs e)
    {
        Close();
    }

    private void InsertHeading(int level)
    {
        string prefix = new string('#', Math.Clamp(level, 1, 6)) + " ";
        PrefixSelectedLines(prefix);
        StatusText.Text = $"Inserted H{Math.Clamp(level, 1, 6)} heading markers";
    }

    private void BoldButton_Click(object? sender, RoutedEventArgs e)
    {
        WrapSelection(BoldMarker, BoldPlaceholder);
        StatusText.Text = StatusBoldApplied;
    }

    private void ItalicButton_Click(object? sender, RoutedEventArgs e)
    {
        WrapSelection(ItalicMarker, ItalicPlaceholder);
        StatusText.Text = StatusItalicApplied;
    }

    private void StrikethroughButton_Click(object? sender, RoutedEventArgs e)
    {
        WrapSelection(StrikethroughMarker, StrikethroughPlaceholder);
        StatusText.Text = StatusStrikethroughApplied;
    }

    private void InlineCodeButton_Click(object? sender, RoutedEventArgs e)
    {
        WrapSelection(InlineCodeMarker, CodePlaceholder);
        StatusText.Text = StatusInlineCodeApplied;
    }

    private void CodeBlockButton_Click(object? sender, RoutedEventArgs e)
    {
        string selectedText = EditorTextBox.SelectedText;
        if (string.IsNullOrEmpty(selectedText))
        {
            selectedText = CodePlaceholder;
        }

        string replacement = $"{CodeBlockMarker}{Environment.NewLine}{selectedText}{Environment.NewLine}{CodeBlockMarker}";
        ReplaceSelection(replacement, replacement.Length);
        StatusText.Text = StatusCodeBlockInserted;
    }

    private void HorizontalRuleButton_Click(object? sender, RoutedEventArgs e)
    {
        string replacement = $"{Environment.NewLine}{HorizontalRuleMarker}{Environment.NewLine}";
        ReplaceSelection(replacement, replacement.Length);
        StatusText.Text = StatusHorizontalRuleInserted;
    }

    private void BulletListButton_Click(object? sender, RoutedEventArgs e)
    {
        PrefixSelectedLines(BulletListPrefix);
        StatusText.Text = StatusBulletListInserted;
    }

    private void NumberedListButton_Click(object? sender, RoutedEventArgs e)
    {
        TransformSelectedLinesWithIndex((line, index) => $"{index + 1}. {line}", ListItemPlaceholder);
        StatusText.Text = StatusNumberedListInserted;
    }

    private void QuoteButton_Click(object? sender, RoutedEventArgs e)
    {
        PrefixSelectedLines(QuotePrefix);
        StatusText.Text = StatusQuoteInserted;
    }

    private void LinkButton_Click(object? sender, RoutedEventArgs e)
    {
        string selectedText = string.IsNullOrWhiteSpace(EditorTextBox.SelectedText) ? LinkTextPlaceholder : EditorTextBox.SelectedText;
        string replacement = $"[{selectedText}]({LinkUrlPlaceholder})";
        ReplaceSelection(replacement, replacement.IndexOf(LinkUrlPlaceholder, StringComparison.Ordinal));
        StatusText.Text = StatusLinkInserted;
    }

    private void ImageButton_Click(object? sender, RoutedEventArgs e)
    {
        string selectedText = string.IsNullOrWhiteSpace(EditorTextBox.SelectedText) ? ImageAltPlaceholder : EditorTextBox.SelectedText;
        string replacement = $"![{selectedText}]({ImageUrlPlaceholder})";
        ReplaceSelection(replacement, replacement.IndexOf(ImageUrlPlaceholder, StringComparison.Ordinal));
        StatusText.Text = StatusImageInserted;
    }

    /// <summary>
    /// Wraps the selected text (or a placeholder) with markdown syntax markers.
    /// </summary>
    /// <param name="marker">The markdown syntax to wrap with (e.g., "**" for bold).</param>
    /// <param name="placeholder">Text to use if no selection exists.</param>
    private void WrapSelection(string marker, string placeholder)
    {
        string selected = EditorTextBox.SelectedText;
        if (string.IsNullOrEmpty(selected))
        {
            selected = placeholder;
        }

        string replacement = $"{marker}{selected}{marker}";
        int caretOffset = marker.Length + selected.Length + marker.Length;
        ReplaceSelection(replacement, caretOffset);
    }

    private void PrefixSelectedLines(string prefix)
    {
        TransformSelectedLines(line => $"{prefix}{line}", "text");
    }

    /// <summary>
    /// Transforms each line of selected text using the provided function.
    /// </summary>
    private void TransformSelectedLines(Func<string, string> transformer, string fallbackText)
    {
        string selected = EditorTextBox.SelectedText;
        if (string.IsNullOrEmpty(selected))
        {
            selected = fallbackText;
        }

        string normalized = selected.Replace("\r\n", "\n");
        string[] lines = normalized.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            lines[i] = transformer(lines[i]);
        }

        string replacement = string.Join(Environment.NewLine, lines);
        ReplaceSelection(replacement, replacement.Length);
    }

    private void TransformSelectedLinesWithIndex(Func<string, int, string> transformer, string fallbackText)
    {
        string selected = EditorTextBox.SelectedText;
        if (string.IsNullOrEmpty(selected))
        {
            selected = fallbackText;
        }

        string normalized = selected.Replace("\r\n", "\n");
        string[] lines = normalized.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            lines[i] = transformer(lines[i], i);
        }

        string replacement = string.Join(Environment.NewLine, lines);
        ReplaceSelection(replacement, replacement.Length);
    }

    private void ReplaceSelection(string replacementText, int caretOffsetAfterInsert)
    {
        int selectionStart = EditorTextBox.SelectionStart;
        EditorTextBox.SelectedText = replacementText;
        EditorTextBox.Focus();
        EditorTextBox.CaretIndex = selectionStart + Math.Clamp(caretOffsetAfterInsert, 0, replacementText.Length);
    }

    private void ViewMarkdownMenuItem_Click(object? sender, RoutedEventArgs e)
    {
        // If unchecking the last visible view, automatically check the other view
        if (!ViewMarkdownMenuItem.IsChecked && !ViewPreviewMenuItem.IsChecked)
        {
            ViewPreviewMenuItem.IsChecked = true;
        }
        UpdateViewVisibility();
    }

    private void ViewPreviewMenuItem_Click(object? sender, RoutedEventArgs e)
    {
        // If unchecking the last visible view, automatically check the other view
        if (!ViewMarkdownMenuItem.IsChecked && !ViewPreviewMenuItem.IsChecked)
        {
            ViewMarkdownMenuItem.IsChecked = true;
        }
        UpdateViewVisibility();
    }

    /// <summary>
    /// Updates UI visibility and layout based on which views are enabled.
    /// Supports split view, markdown-only, and preview-only modes.
    /// </summary>
    private void UpdateViewVisibility()
    {
        bool showEditor = ViewMarkdownMenuItem.IsChecked;
        bool showPreview = ViewPreviewMenuItem.IsChecked;

        if (showEditor && showPreview)
        {
            EditorBorder.Visibility = Visibility.Visible;
            ViewSplitter.Visibility = Visibility.Visible;
            PreviewBorder.Visibility = Visibility.Visible;
            EditorColumn.Width = new GridLength(1, GridUnitType.Star);
            SplitterColumn.Width = new GridLength(10);
            PreviewColumn.Width = new GridLength(1, GridUnitType.Star);
        }
        else if (showEditor)
        {
            EditorBorder.Visibility = Visibility.Visible;
            ViewSplitter.Visibility = Visibility.Collapsed;
            PreviewBorder.Visibility = Visibility.Collapsed;
            EditorColumn.Width = new GridLength(1, GridUnitType.Star);
            SplitterColumn.Width = new GridLength(0);
            PreviewColumn.Width = new GridLength(0);
        }
        else
        {
            // Preview only
            EditorBorder.Visibility = Visibility.Collapsed;
            ViewSplitter.Visibility = Visibility.Collapsed;
            PreviewBorder.Visibility = Visibility.Visible;
            EditorColumn.Width = new GridLength(0);
            SplitterColumn.Width = new GridLength(0);
            PreviewColumn.Width = new GridLength(1, GridUnitType.Star);
        }

        Visibility editorToolsVisibility = showEditor ? Visibility.Visible : Visibility.Collapsed;
        ToolbarSeparator1.Visibility = editorToolsVisibility;
        ToolbarBoldButton.Visibility = editorToolsVisibility;
        ToolbarItalicButton.Visibility = editorToolsVisibility;
        ToolbarStrikeButton.Visibility = editorToolsVisibility;
        ToolbarCodeButton.Visibility = editorToolsVisibility;
        ToolbarSeparator2.Visibility = editorToolsVisibility;
        ToolbarLinkButton.Visibility = editorToolsVisibility;
        ToolbarImageButton.Visibility = editorToolsVisibility;
        ToolbarSeparator3.Visibility = editorToolsVisibility;
        ToolbarBulletButton.Visibility = editorToolsVisibility;
        ToolbarNumberButton.Visibility = editorToolsVisibility;
        ToolbarQuoteButton.Visibility = editorToolsVisibility;
        ToolbarSeparator4.Visibility = editorToolsVisibility;
        HeadingLabel.Visibility = editorToolsVisibility;
        HeadingComboBox.Visibility = editorToolsVisibility;
        ToolbarHeadingButton.Visibility = editorToolsVisibility;
    }

    private void EditorBorder_PreviewDragOver(object? sender, DragEventArgs e)
    {
        HandleDragOver(e);
    }

    private void EditorBorder_PreviewDrop(object? sender, DragEventArgs e)
    {
        HandleFileDrop(e);
    }

    private void PreviewBorder_PreviewDragOver(object? sender, DragEventArgs e)
    {
        HandleDragOver(e);
    }

    private void PreviewBorder_PreviewDrop(object? sender, DragEventArgs e)
    {
        HandleFileDrop(e);
    }

    private void EditorTextBox_PreviewDragOver(object? sender, DragEventArgs e)
    {
        HandleDragOver(e);
    }

    private void EditorTextBox_PreviewDrop(object? sender, DragEventArgs e)
    {
        HandleFileDrop(e);
    }

    private void EditorTextBox_Drop(object? sender, DragEventArgs e)
    {
        HandleFileDrop(e);
    }

    private void PreviewBrowser_NavigationStarting(object? sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationStartingEventArgs e)
    {
        if (e.Uri != null && e.Uri.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
        {
            e.Cancel = true;
            OpenFileFromUri(e.Uri);
        }
    }

    private void CoreWebView2_NewWindowRequested(object? sender, Microsoft.Web.WebView2.Core.CoreWebView2NewWindowRequestedEventArgs e)
    {
        if (e.Uri != null && e.Uri.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
        {
            e.Handled = true;
            OpenFileFromUri(e.Uri);
        }
    }

    private void OpenFileFromUri(string uri)
    {
        try
        {
            string filePath = new Uri(uri).LocalPath;
            if (Path.GetExtension(filePath).Equals(MarkdownExtension, StringComparison.OrdinalIgnoreCase))
            {
                string content = File.ReadAllText(filePath);
                EditorTextBox.Text = content;
                _currentFileName = Path.GetFileName(filePath);
                StatusText.Text = $"Opened: {_currentFileName}";
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Error opening file", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusText.Text = "Failed to open file";
        }
    }

    /// <summary>
    /// Handles drag-over events to show appropriate cursor for .md file drops.
    /// </summary>
    private void HandleDragOver(DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            if (files.Length == 1 && Path.GetExtension(files[0]).Equals(MarkdownExtension, StringComparison.OrdinalIgnoreCase))
            {
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
                return;
            }
        }

        e.Effects = DragDropEffects.None;
        e.Handled = true;
    }

    /// <summary>
    /// Handles file drop events to open .md files in the editor.
    /// </summary>
    private async void HandleFileDrop(DragEventArgs e)
    {
        e.Handled = true;

        try
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                if (files.Length == 1 && Path.GetExtension(files[0]).Equals(MarkdownExtension, StringComparison.OrdinalIgnoreCase))
                {
                    string filePath = files[0];
                    string content = await File.ReadAllTextAsync(filePath);
                    EditorTextBox.Text = content;
                    _currentFileName = Path.GetFileName(filePath);
                    StatusText.Text = $"Opened: {_currentFileName}";
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Error opening file", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusText.Text = "Failed to open file";
        }
    }
}
