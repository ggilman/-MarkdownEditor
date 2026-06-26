using Microsoft.Win32;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;

namespace MarkdownUtilsApp;

public partial class MainWindow : Window
{
    private const string DefaultMarkdown = "# Markdown Utils Example\n\nUse this sample to verify preview and Word export formatting.\n\n## Heading Levels\n\n### H3 Example\n\n#### H4 Example\n\n##### H5 Example\n\n###### H6 Example\n\n## Text Styles\n\nThis line uses **bold**, *italic*, ~~strikethrough~~, and `inline code`.\n\nA markdown link: [Markdown Guide](https://www.markdownguide.org)\n\nAn image placeholder: ![Sample image](https://example.com/image.png)\n\n---\n\n## Lists\n\n- Bullet one\n- Bullet two\n  - Nested bullet\n\n1. First step\n2. Second step\n   1. Nested numbered step\n\n- [x] Task complete\n- [ ] Task pending\n\n## Quote\n\n> Preview pane is read-only and this quote should be styled in Word export.\n\n## Code Block\n\n```csharp\nusing System;\n\nConsole.WriteLine(\"Hello markdown\");\n```\n\n## Table\n\n| Feature | Preview | Word Export |\n| --- | --- | --- |\n| Headings | Yes | Yes |\n| Tables | Yes | Yes |\n| Checkboxes | Yes | Yes |\n| Code Language Label | Yes | Yes |";
    private int _selectedHeadingLevel = 2;
    private bool _isWebView2Initialized = false;

    public MainWindow()
    {
        InitializeComponent();
        EditorTextBox.Text = DefaultMarkdown;
        InitializeWebView2Async();
    }

    private async void InitializeWebView2Async()
    {
        try
        {
            await PreviewBrowser.EnsureCoreWebView2Async();
            _isWebView2Initialized = true;

            PreviewBrowser.CoreWebView2.NewWindowRequested += CoreWebView2_NewWindowRequested;

            // Render initial content
            RenderPreview();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Failed to initialize preview: {ex.Message}", "Initialization Error", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void EditorTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        RenderPreview();
    }

    private void EditorTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
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

    private void RenderPreview()
    {
        string markdown = EditorTextBox.Text;
        string htmlBody = EnhanceCodeBlocksHtml(MarkdownConverter.ToHtml(markdown));

        string htmlDocument = "<!doctype html><html><head><meta charset=\"utf-8\">" +
            "<style>body{font-family:Segoe UI,Arial,sans-serif;padding:12px;line-height:1.5;color:#222;}" +
            ".code-block{margin:1em 0;border-radius:6px;overflow:hidden;border:1px solid #d8d8d8;background:#f3f3f3;}" +
            ".code-header{position:relative;background:#333;color:#f2f2f2;padding:7px 72px 7px 10px;min-height:30px;" +
            "box-sizing:border-box;white-space:normal;}" +
            ".code-lang-badge{display:inline-block;font-size:12px;line-height:1.2;color:#f2f2f2;}" +
            ".code-block pre{margin:0;background:#f3f3f3;overflow:auto;}" +
            ".code-block pre > code{display:block;padding:10px;}" +
            ".code-copy-btn{position:absolute;right:10px;top:6px;border:1px solid #666;background:#4a4a4a;color:#fff;" +
            "border-radius:4px;padding:2px 8px;font-size:12px;line-height:1.2;cursor:pointer;opacity:0;visibility:hidden;" +
            "transition:opacity .15s ease;}" +
            ".code-block:hover .code-copy-btn,.code-block:focus-within .code-copy-btn{opacity:1;visibility:visible;}" +
            ".code-copy-btn:hover{background:#5a5a5a;}" +
            "code{background:#f3f3f3;padding:2px 4px;border-radius:4px;}" +
            "blockquote{border-left:4px solid #999;padding-left:10px;color:#444;margin-left:0;}" +
            "table{border-collapse:collapse;}th,td{border:1px solid #ccc;padding:6px;}</style>" +
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
            "if(copyText(text)){button.innerText='Copied';setTimeout(function(){button.innerText='Copy';},1000);}" +
            "else{button.innerText='Failed';setTimeout(function(){button.innerText='Copy';},1000);}" +
            "return false;" +
            "};" +
            "})();</script></head><body>" +
            htmlBody + "</body></html>";

        if (_isWebView2Initialized)
        {
            PreviewBrowser.NavigateToString(htmlDocument);
        }
    }

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
                string badgeText = string.IsNullOrWhiteSpace(language) ? "code" : language;
                string badge = $"<div class=\"code-header\"><span class=\"code-lang-badge\">{System.Net.WebUtility.HtmlEncode(badgeText)}</span>" +
                               "<button type=\"button\" class=\"code-copy-btn\" onclick=\"return copyCode(this);\">Copy</button></div>";

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

    private void NewButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(EditorTextBox.Text))
        {
            EditorTextBox.Focus();
            StatusText.Text = "Editor already clear";
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
        StatusText.Text = "New document";
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        SaveFileDialog dialog = new()
        {
            Filter = "Markdown files (*.md)|*.md|All files (*.*)|*.*",
            DefaultExt = ".md",
            FileName = "document.md",
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
            StatusText.Text = "Save failed";
        }
    }

    private void ImportButton_Click(object sender, RoutedEventArgs e)
    {
        OpenFileDialog dialog = new()
        {
            Filter = "Markdown files (*.md)|*.md|All files (*.*)|*.*",
            Title = "Import Markdown"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            EditorTextBox.Text = File.ReadAllText(dialog.FileName);
            StatusText.Text = $"Imported: {Path.GetFileName(dialog.FileName)}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Import failed", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusText.Text = "Import failed";
        }
    }

    private void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        string suggestedFileName = GetSuggestedFileName(EditorTextBox.Text) ?? "document.docx";

        SaveFileDialog dialog = new()
        {
            Filter = "Word document (*.docx)|*.docx",
            DefaultExt = ".docx",
            FileName = suggestedFileName,
            Title = "Export to Word"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            bool autoNumberHeadings = AutoNumberHeadingsCheckBox.IsChecked ?? true;
            bool useTitleStyle = UseTitleStyleCheckBox.IsChecked ?? true;
            DocxExporter.ExportFromMarkdown(EditorTextBox.Text, dialog.FileName, autoNumberHeadings, useTitleStyle);
            StatusText.Text = $"Exported: {Path.GetFileName(dialog.FileName)}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Export failed", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusText.Text = "Export failed";
        }
    }

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
                        return sanitized + ".docx";
                    }
                }
            }
        }

        return null;
    }

    private static string SanitizeFileName(string fileName)
    {
        char[] invalidChars = Path.GetInvalidFileNameChars();
        string sanitized = string.Concat(fileName.Select(c => invalidChars.Contains(c) ? '_' : c));
        sanitized = sanitized.Trim();

        if (sanitized.Length > 200)
        {
            sanitized = sanitized.Substring(0, 200).TrimEnd();
        }

        return sanitized;
    }

    private void HeadingApplyButton_Click(object sender, RoutedEventArgs e)
    {
        InsertHeading(_selectedHeadingLevel);
    }

    private void HeadingDropdownButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not ToggleButton toggleButton || toggleButton.ContextMenu is null)
        {
            return;
        }

        toggleButton.ContextMenu.PlacementTarget = toggleButton;
        toggleButton.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
        toggleButton.ContextMenu.IsOpen = true;
    }

    private void HeadingLevelMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem menuItem)
        {
            return;
        }

        if (!int.TryParse(menuItem.Tag?.ToString(), out int level))
        {
            return;
        }

        _selectedHeadingLevel = Math.Clamp(level, 1, 6);
        HeadingApplyButton.Content = $"H{_selectedHeadingLevel}";
        StatusText.Text = $"Selected H{_selectedHeadingLevel}";

        HeadingDropdownButton.IsChecked = false;
    }

    private void InsertHeading(int level)
    {
        string prefix = new string('#', Math.Clamp(level, 1, 6)) + " ";
        PrefixSelectedLines(prefix);
        StatusText.Text = $"Inserted H{Math.Clamp(level, 1, 6)} heading markers";
    }

    private void BoldButton_Click(object sender, RoutedEventArgs e)
    {
        WrapSelection("**", "bold text");
        StatusText.Text = "Applied bold formatting";
    }

    private void ItalicButton_Click(object sender, RoutedEventArgs e)
    {
        WrapSelection("*", "italic text");
        StatusText.Text = "Applied italic formatting";
    }

    private void StrikethroughButton_Click(object sender, RoutedEventArgs e)
    {
        WrapSelection("~~", "strikethrough text");
        StatusText.Text = "Applied strikethrough formatting";
    }

    private void InlineCodeButton_Click(object sender, RoutedEventArgs e)
    {
        WrapSelection("`", "code");
        StatusText.Text = "Applied inline code formatting";
    }

    private void CodeBlockButton_Click(object sender, RoutedEventArgs e)
    {
        string selectedText = EditorTextBox.SelectedText;
        if (string.IsNullOrEmpty(selectedText))
        {
            selectedText = "code";
        }

        string replacement = $"```\n{selectedText}\n```";
        ReplaceSelection(replacement, replacement.Length);
        StatusText.Text = "Inserted code block";
    }

    private void HorizontalRuleButton_Click(object sender, RoutedEventArgs e)
    {
        string replacement = $"{Environment.NewLine}---{Environment.NewLine}";
        ReplaceSelection(replacement, replacement.Length);
        StatusText.Text = "Inserted horizontal rule";
    }

    private void BulletListButton_Click(object sender, RoutedEventArgs e)
    {
        PrefixSelectedLines("- ");
        StatusText.Text = "Inserted bullet list markers";
    }

    private void NumberedListButton_Click(object sender, RoutedEventArgs e)
    {
        TransformSelectedLinesWithIndex((line, index) => $"{index + 1}. {line}", "list item");
        StatusText.Text = "Inserted numbered list markers";
    }

    private void QuoteButton_Click(object sender, RoutedEventArgs e)
    {
        PrefixSelectedLines("> ");
        StatusText.Text = "Inserted quote markers";
    }

    private void LinkButton_Click(object sender, RoutedEventArgs e)
    {
        string selectedText = string.IsNullOrWhiteSpace(EditorTextBox.SelectedText) ? "link text" : EditorTextBox.SelectedText;
        string replacement = $"[{selectedText}](https://example.com)";
        ReplaceSelection(replacement, replacement.IndexOf("https://example.com", StringComparison.Ordinal));
        StatusText.Text = "Inserted markdown link";
    }

    private void ImageButton_Click(object sender, RoutedEventArgs e)
    {
        string selectedText = string.IsNullOrWhiteSpace(EditorTextBox.SelectedText) ? "alt text" : EditorTextBox.SelectedText;
        string replacement = $"![{selectedText}](https://example.com/image.png)";
        ReplaceSelection(replacement, replacement.IndexOf("https://example.com/image.png", StringComparison.Ordinal));
        StatusText.Text = "Inserted markdown image";
    }

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

    private void ViewMarkdownMenuItem_Click(object sender, RoutedEventArgs e)
    {
        // If unchecking the last visible view, automatically check the other view
        if (!ViewMarkdownMenuItem.IsChecked && !ViewPreviewMenuItem.IsChecked)
        {
            ViewPreviewMenuItem.IsChecked = true;
        }
        UpdateViewVisibility();
    }

    private void ViewPreviewMenuItem_Click(object sender, RoutedEventArgs e)
    {
        // If unchecking the last visible view, automatically check the other view
        if (!ViewMarkdownMenuItem.IsChecked && !ViewPreviewMenuItem.IsChecked)
        {
            ViewMarkdownMenuItem.IsChecked = true;
        }
        UpdateViewVisibility();
    }

    private void UpdateViewVisibility()
    {
        bool showEditor = ViewMarkdownMenuItem.IsChecked;
        bool showPreview = ViewPreviewMenuItem.IsChecked;

        // Update visibility and column widths
        if (showEditor && showPreview)
        {
            // Both visible - split view
            EditorBorder.Visibility = Visibility.Visible;
            ViewSplitter.Visibility = Visibility.Visible;
            PreviewBorder.Visibility = Visibility.Visible;
            EditorColumn.Width = new GridLength(1, GridUnitType.Star);
            SplitterColumn.Width = new GridLength(10);
            PreviewColumn.Width = new GridLength(1, GridUnitType.Star);
        }
        else if (showEditor)
        {
            // Markdown only
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
    }

    private void EditorTextBox_PreviewDragOver(object sender, DragEventArgs e)
    {
        HandleDragOver(e);
    }

    private void EditorTextBox_PreviewDrop(object sender, DragEventArgs e)
    {
        HandleFileDrop(e);
    }

    private void EditorTextBox_Drop(object sender, DragEventArgs e)
    {
        HandleFileDrop(e);
    }

    private void PreviewBrowser_NavigationStarting(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationStartingEventArgs e)
    {
        if (e.Uri != null && e.Uri.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
        {
            e.Cancel = true;
            OpenFileFromUri(e.Uri);
        }
    }

    private void CoreWebView2_NewWindowRequested(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NewWindowRequestedEventArgs e)
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
            if (Path.GetExtension(filePath).Equals(".md", StringComparison.OrdinalIgnoreCase))
            {
                string content = File.ReadAllText(filePath);
                EditorTextBox.Text = content;
                StatusText.Text = $"Opened: {Path.GetFileName(filePath)}";
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Error opening file", MessageBoxButton.OK, MessageBoxImage.Error);
            StatusText.Text = "Failed to open file";
        }
    }

    private void HandleDragOver(DragEventArgs e)
    {
        // Check if the drag data contains files
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            // Check if it's a single .md file
            if (files.Length == 1 && Path.GetExtension(files[0]).Equals(".md", StringComparison.OrdinalIgnoreCase))
            {
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
                return;
            }
        }

        e.Effects = DragDropEffects.None;
        e.Handled = true;
    }

    private void HandleFileDrop(DragEventArgs e)
    {
        // Always mark as handled to prevent default behavior
        e.Handled = true;

        try
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                if (files.Length == 1 && Path.GetExtension(files[0]).Equals(".md", StringComparison.OrdinalIgnoreCase))
                {
                    string filePath = files[0];
                    string content = File.ReadAllText(filePath);
                    EditorTextBox.Text = content;
                    StatusText.Text = $"Opened: {Path.GetFileName(filePath)}";
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
