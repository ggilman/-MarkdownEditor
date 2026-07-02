using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using MdTable = Markdig.Extensions.Tables.Table;
using MdTableCell = Markdig.Extensions.Tables.TableCell;
using MdTableRow = Markdig.Extensions.Tables.TableRow;
using WTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using WTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using Level = DocumentFormat.OpenXml.Wordprocessing.Level;
using NumberingFormat = DocumentFormat.OpenXml.Wordprocessing.NumberingFormat;

namespace MarkdownEditor;

/// <summary>
/// Exports markdown documents to Word (.docx) format using OpenXML.
/// Supports headings, lists, tables, code blocks, and inline formatting with high fidelity.
/// </summary>
internal static class DocxExporter
{
    // Page dimensions (in twips: 1/20 of a point, 1/1440 of an inch)
    private const int PageWidthTwips = 12240;  // 8.5 inches
    private const int PageHeightTwips = 15840; // 11 inches
    private const int PageMarginTwips = 1440;   // 1 inch
    private const int PageContentWidthTwips = 9360;

    // Numbering IDs
    private const int BaseNumberingIdForLists = 10;
    private const int AbstractNumBulletList = 1;
    private const int AbstractNumNumberedList = 2;
    private const int AbstractNumHeadingNumbering = 3;

    // Spacing values (twips)
    private const string ParagraphSpacingAfter = "200";    // 10pt
    private const string ParagraphLineSpacing = "360";     // 18pt
    private const string ListItemSpacingAfter = "120";    // 6pt
    private const string HeadingSpacingBefore = "240";    // 12pt
    private const string HeadingSpacingAfter = "120";     // 6pt
    private const string QuoteSpacingAfter = "120";       // 6pt
    private const string CodeHeaderSpacingBefore = "120";  // 6pt
    private const string CodeHeaderSpacingAfter = "40";    // 2pt
    private const string CodeBlockSpacingBefore = "120";   // 6pt
    private const string CodeBlockSpacingAfter = "120";    // 6pt
    private const string CodeBlockLineSpacing = "276";     // 13.8pt

    // Indentation values (twips)
    private const int IndentPerLevel = 360;        // 0.25 inch per level
    private const int BaseListIndent = 720;        // 0.5 inch
    private const int HangingIndent = 360;         // 0.25 inch
    private const int QuoteIndentLeft = 720;       // 0.5 inch
    private const int TaskListIndentHanging = 240; // Checkbox indent
    private const string CodeBlockIndentLeft = "360";   // 0.25 inch
    private const string CodeBlockIndentRight = "120";  // ~0.08 inch

    // Font settings
    private const string CodeFontName = "Consolas";
    private const string CodeFontSize = "20";      // 10pt (half-points)
    private const string CodeHeaderFontName = "Calibri";

    // Color values
    private const string HeadingColorBlue = "2E74B5";      // Title and H1-H2
    private const string HeadingColorDarkBlue = "1F4D78";  // H3-H5
    private const string CodeBackgroundGray = "F3F3F3";
    private const string CodeBlockBorderGray = "D8D8D8";
    private const string CodeHeaderBackgroundDark = "333333";
    private const string CodeHeaderForegroundLight = "F2F2F2";
    private const string QuoteBorderGray = "999999";
    private const string TableBorderGray = "CCCCCC";
    private const string HyperlinkColor = "0563C1";        // Standard blue
    private const string QuoteTextGray = "444444";

    // Font sizes (half-points: size * 2)
    private const string TitleFontSize = "52";     // 26pt
    private const string Heading1FontSize = "32";  // 16pt
    private const string Heading2FontSize = "26";  // 13pt
    private const string Heading3FontSize = "24";  // 12pt
    private const string Heading4FontSize = "22";  // 11pt
    private const string Heading5FontSize = "20";  // 10pt
    private const string CodeHeaderFontSize = "18";  // 9pt

    // Character spacing
    private const string TitleCharacterSpacing = "-10";  // Slightly condensed

    // Border widths (eighths of a point)
    private const uint QuoteBorderWidth = 24;            // 3pt
    private const uint CodeBlockBorderWidth = 4;         // 0.5pt
    private const uint TableBorderWidth = 4;              // 0.5pt
    private const uint TableOuterBorderWidth = 8;        // 1pt
    private const uint TableInnerBorderWidth = 6;        // 0.75pt
    private const uint HorizontalRuleBorderWidth = 6;    // 0.75pt

    // Task list markers
    private const string TaskCheckedMarker = "☑ ";
    private const string TaskUncheckedMarker = "☐ ";

    // Regex patterns
    private const string HeadingNumberPattern = @"^\s*\d+(\.\d+)*\.?\s*";
    private const string LanguageClassPrefix = "language-";

    // Style IDs
    private const string StyleIdNormal = "Normal";
    private const string StyleIdTitle = "Title";
    private const string StyleIdHeading = "Heading";
    private const string StyleIdCodeBlock = "CodeBlock";
    private const string StyleIdQuote = "Quote";
    private const string StyleIdHyperlink = "Hyperlink";

    // Table and cell constants
    private const string TableAutoWidthValue = "0";

    // List numbering constants
    private const int MaxListLevels = 9;
    private const int MaxHeadingLevels = 5;
    private const int StartNumberValue = 1;
    private const string ListHangingIndent = "360";

    // Bullet characters for different nesting levels
    private const string BulletCharPrimary = "\uF0B7";   // Standard bullet
    private const string BulletCharSecondary = "o";      // Circle
    private const string BulletCharTertiary = "\uF0A7";  // Square

    // Font names for bullets
    private const string BulletFontSymbol = "Symbol";
    private const string BulletFontCourier = "Courier New";

    // Style priorities
    private const int HeadingStylePriority = 9;
    private const int TitleStylePriority = 10;

    // Zero spacing value
    private const string ZeroSpacing = "0";

    private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder()
        .UseAdvancedExtensions()
        .DisableHtml()
        .Build();

    /// <summary>
    /// Exports markdown text to a Word document (.docx) file.
    /// </summary>
    /// <param name="markdown">The markdown text to export.</param>
    /// <param name="outputPath">The file path where the .docx will be created.</param>
    /// <param name="autoNumberHeadings">If true, strips manual numbers from headings and applies Word's multilevel numbering.</param>
    /// <param name="useTitleStyle">If true, maps markdown # to Word's Title style; otherwise uses Heading 1.</param>
    public static void ExportFromMarkdown(string markdown, string outputPath, bool autoNumberHeadings = true, bool useTitleStyle = true)
    {
        if (string.IsNullOrWhiteSpace(outputPath))
        {
            throw new ArgumentException("Output path is required.", nameof(outputPath));
        }

        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        if (File.Exists(outputPath))
        {
            File.Delete(outputPath);
        }

        MarkdownDocument document = Markdown.Parse(markdown ?? string.Empty, Pipeline);

        using WordprocessingDocument wordDocument = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        EnsureStylesPart(mainPart, autoNumberHeadings, useTitleStyle);
        EnsureNumberingPart(mainPart, autoNumberHeadings);
        Body body = mainPart.Document.Body!;
        Dictionary<string, string> hyperlinkRelationshipIds = new(StringComparer.OrdinalIgnoreCase);
        int listNumberingCounter = 0;

        foreach (Block block in document)
        {
            AppendBlock(body, block, mainPart, hyperlinkRelationshipIds, autoNumberHeadings, useTitleStyle, ref listNumberingCounter);
        }

        body.Append(new SectionProperties(
            new PageSize { Width = (uint)PageWidthTwips, Height = (uint)PageHeightTwips },
            new PageMargin { Top = PageMarginTwips, Right = PageMarginTwips, Bottom = PageMarginTwips, Left = PageMarginTwips }
        ));

        mainPart.Document.Save();
    }

    private static void AppendBlock(Body body, Block block, MainDocumentPart mainPart, Dictionary<string, string> hyperlinkRelationshipIds, bool autoNumberHeadings, bool useTitleStyle, ref int listNumberingCounter)
    {
        switch (block)
        {
            case HeadingBlock heading:
                body.Append(CreateHeadingParagraph(heading, autoNumberHeadings, useTitleStyle));
                break;
            case ParagraphBlock paragraph:
                body.Append(CreateParagraph(paragraph.Inline, mainPart, hyperlinkRelationshipIds));
                break;
            case QuoteBlock quote:
                foreach (Block child in quote)
                {
                    Paragraph quoteParagraph = CreateQuoteParagraph(ExtractPlainText(child));
                    body.Append(quoteParagraph);
                }
                break;
            case FencedCodeBlock fenced:
                AppendFencedCodeBlock(body, fenced);
                break;
            case CodeBlock code:
                body.Append(CreateCodeParagraph(ExtractLinesText(code)));
                break;
            case ListBlock list:
                AppendList(body, list, 0, mainPart, hyperlinkRelationshipIds, ref listNumberingCounter);
                break;
            case MdTable table:
                body.Append(CreateTable(table, mainPart, hyperlinkRelationshipIds));
                break;
            case ThematicBreakBlock:
                body.Append(CreateHorizontalRule());
                break;
            default:
                if (block is ContainerBlock container)
                {
                    foreach (Block child in container)
                    {
                        AppendBlock(body, child, mainPart, hyperlinkRelationshipIds, autoNumberHeadings, useTitleStyle, ref listNumberingCounter);
                    }
                }
                else
                {
                    string fallback = ExtractPlainText(block);
                    if (!string.IsNullOrWhiteSpace(fallback))
                    {
                        body.Append(CreateParagraphFromText(fallback));
                    }
                }
                break;
        }
    }

    private static void AppendList(Body body, ListBlock list)
    {
        int counter = 0;
        AppendList(body, list, 0, null, null, ref counter);
    }

    private static void AppendList(Body body, ListBlock list, int depth)
    {
        int counter = 0;
        AppendList(body, list, depth, null, null, ref counter);
    }

    private static void AppendList(Body body, ListBlock list, int depth, MainDocumentPart? mainPart, Dictionary<string, string>? hyperlinkRelationshipIds, ref int listNumberingCounter)
    {
        // Create a unique numbering instance for this list to prevent Word from continuing numbering from previous lists
        NumberingDefinitionsPart? numberingPart = mainPart?.NumberingDefinitionsPart;
        int numberingId;

        if (numberingPart?.Numbering != null)
        {
            int abstractNumId = list.IsOrdered ? AbstractNumNumberedList : AbstractNumBulletList;

            numberingId = BaseNumberingIdForLists + listNumberingCounter++;
            NumberingInstance newInstance = new() { NumberID = numberingId };
            newInstance.Append(new AbstractNumId { Val = abstractNumId });
            numberingPart.Numbering.Append(newInstance);
            numberingPart.Numbering.Save();
        }
        else
        {
            numberingId = list.IsOrdered ? AbstractNumNumberedList : AbstractNumBulletList;
        }

        foreach (ListItemBlock item in list.OfType<ListItemBlock>())
        {
            bool? taskState = TryGetTaskState(item);
            bool isFirstParagraphInItem = true;

            foreach (Block child in item)
            {
                if (child is ListBlock nestedList)
                {
                    AppendList(body, nestedList, depth + 1, mainPart, hyperlinkRelationshipIds, ref listNumberingCounter);
                    continue;
                }

                if (child is ParagraphBlock paragraphBlock && paragraphBlock.Inline is not null)
                {
                    Paragraph listParagraph;

                    if (taskState.HasValue)
                    {
                        string marker = taskState.Value ? TaskCheckedMarker : TaskUncheckedMarker;
                        listParagraph = CreateListItemParagraph(
                            marker,
                            paragraphBlock.Inline,
                            mainPart,
                            hyperlinkRelationshipIds,
                            indentLeftTwips: IndentPerLevel * (depth + 1),
                            indentHangingTwips: TaskListIndentHanging
                        );
                    }
                    else if (isFirstParagraphInItem)
                    {
                        listParagraph = CreateNativeListItemParagraph(
                            paragraphBlock.Inline,
                            mainPart,
                            hyperlinkRelationshipIds,
                            numberingId,
                            depth
                        );
                        isFirstParagraphInItem = false;
                    }
                    else
                    {
                        listParagraph = CreateNativeListItemParagraph(
                            paragraphBlock.Inline,
                            mainPart,
                            hyperlinkRelationshipIds,
                            numberingId,
                            depth
                        );
                    }

                    body.Append(listParagraph);
                }
                else
                {
                    string lineText = ExtractPlainText(child);
                    if (!string.IsNullOrWhiteSpace(lineText))
                    {
                        if (taskState.HasValue)
                        {
                            string marker = taskState.Value ? TaskCheckedMarker : TaskUncheckedMarker;
                            body.Append(CreateParagraphFromText(marker + lineText, indentLeftTwips: IndentPerLevel * (depth + 1), indentHangingTwips: TaskListIndentHanging));
                        }
                        else
                        {
                            Paragraph listParagraph = CreateNativeListItemParagraphFromText(
                                lineText,
                                numberingId,
                                depth
                            );
                            body.Append(listParagraph);
                        }
                    }
                }
            }
        }
    }

    private static bool? TryGetTaskState(ListItemBlock item)
    {
        foreach (ParagraphBlock paragraph in item.OfType<ParagraphBlock>())
        {
            if (paragraph.Inline is null)
            {
                continue;
            }

            foreach (Inline inline in paragraph.Inline)
            {
                string? typeName = inline.GetType().FullName;
                if (typeName is null || !typeName.Contains("TaskList", StringComparison.Ordinal))
                {
                    continue;
                }

                if (inline is Markdig.Extensions.TaskLists.TaskList taskList)
                {
                    return taskList.Checked;
                }

                return false;
            }
        }

        return null;
    }

    private static Paragraph CreateParagraph(ContainerInline? inline, MainDocumentPart mainPart, Dictionary<string, string> hyperlinkRelationshipIds, string? paragraphStyleId = null)
    {
        Paragraph paragraph = new();
        ParagraphProperties props = new(
            new SpacingBetweenLines { After = ParagraphSpacingAfter, Line = ParagraphLineSpacing, LineRule = LineSpacingRuleValues.Auto }
        );

        if (!string.IsNullOrWhiteSpace(paragraphStyleId))
        {
            props.Append(new ParagraphStyleId { Val = paragraphStyleId });
        }

        paragraph.Append(props);

        if (inline is null)
        {
            return paragraph;
        }

        AppendInlines(paragraph, inline, mainPart, hyperlinkRelationshipIds);
        return paragraph;
    }

    private static Paragraph CreateListItemParagraph(string marker, ContainerInline inline, MainDocumentPart? mainPart, Dictionary<string, string>? hyperlinkRelationshipIds, int indentLeftTwips, int indentHangingTwips)
    {
        Paragraph paragraph = new();
        ParagraphProperties props = new(
            new SpacingBetweenLines { After = ListItemSpacingAfter, Line = ParagraphLineSpacing, LineRule = LineSpacingRuleValues.Auto },
            new Indentation 
            { 
                Left = indentLeftTwips.ToString(), 
                Hanging = indentHangingTwips.ToString() 
            }
        );

        paragraph.Append(props);

        paragraph.Append(new Run(new Text(marker) { Space = SpaceProcessingModeValues.Preserve }));

        if (inline is not null && mainPart is not null && hyperlinkRelationshipIds is not null)
        {
            AppendInlines(paragraph, inline, mainPart, hyperlinkRelationshipIds);
        }

        return paragraph;
    }

    private static Paragraph CreateNativeListItemParagraph(ContainerInline inline, MainDocumentPart? mainPart, Dictionary<string, string>? hyperlinkRelationshipIds, int numberingId, int levelIndex)
    {
        Paragraph paragraph = new();
        ParagraphProperties props = new(
            new SpacingBetweenLines { After = ListItemSpacingAfter, Line = ParagraphLineSpacing, LineRule = LineSpacingRuleValues.Auto },
            new NumberingProperties(
                new NumberingLevelReference { Val = levelIndex },
                new NumberingId { Val = numberingId }
            )
        );

        paragraph.Append(props);

        if (inline is not null && mainPart is not null && hyperlinkRelationshipIds is not null)
        {
            AppendInlines(paragraph, inline, mainPart, hyperlinkRelationshipIds);
        }

        return paragraph;
    }

    private static Paragraph CreateNativeListItemParagraphFromText(string text, int numberingId, int levelIndex)
    {
        Paragraph paragraph = new();
        ParagraphProperties props = new(
            new SpacingBetweenLines { After = ListItemSpacingAfter, Line = ParagraphLineSpacing, LineRule = LineSpacingRuleValues.Auto },
            new NumberingProperties(
                new NumberingLevelReference { Val = levelIndex },
                new NumberingId { Val = numberingId }
            )
        );

        paragraph.Append(props);
        paragraph.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

        return paragraph;
    }

    private static Paragraph CreateParagraphFromText(string text, int? indentLeftTwips = null, int? indentHangingTwips = null)
    {
        Paragraph paragraph = new();
        if (indentLeftTwips.HasValue || indentHangingTwips.HasValue)
        {
            Indentation indentation = new();
            if (indentLeftTwips.HasValue)
            {
                indentation.Left = indentLeftTwips.Value.ToString();
            }

            if (indentHangingTwips.HasValue)
            {
                indentation.Hanging = indentHangingTwips.Value.ToString();
            }

            paragraph.Append(new ParagraphProperties(indentation));
        }

        paragraph.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }

    private static Paragraph CreateTableCellParagraph(MdTableCell cell, MainDocumentPart? mainPart, Dictionary<string, string>? hyperlinkRelationshipIds, bool isHeader)
    {
        Paragraph paragraph = new();
        ParagraphProperties props = new();

        if (isHeader)
        {
            props.Append(new Justification { Val = JustificationValues.Center });
        }

        paragraph.Append(props);

        // Process each block in the cell (usually just one paragraph block)
        foreach (Block block in cell)
        {
            if (block is ParagraphBlock paragraphBlock && paragraphBlock.Inline is not null)
            {
                // Process inline content with formatting
                if (mainPart is not null && hyperlinkRelationshipIds is not null)
                {
                    AppendInlines(paragraph, paragraphBlock.Inline, mainPart, hyperlinkRelationshipIds, isHeader);
                }
                else
                {
                    // Fallback to plain text if we don't have the necessary parts
                    string plainText = ExtractPlainText(block);
                    if (!string.IsNullOrWhiteSpace(plainText))
                    {
                        Run run = new(new Text(plainText) { Space = SpaceProcessingModeValues.Preserve });
                        if (isHeader)
                        {
                            run.RunProperties = new RunProperties(new Bold());
                        }
                        paragraph.Append(run);
                    }
                }
            }
            else
            {
                // Handle other block types (code blocks, etc.) as plain text
                string plainText = ExtractPlainText(block);
                if (!string.IsNullOrWhiteSpace(plainText))
                {
                    Run run = new(new Text(plainText) { Space = SpaceProcessingModeValues.Preserve });
                    if (isHeader)
                    {
                        run.RunProperties = new RunProperties(new Bold());
                    }
                    paragraph.Append(run);
                }
            }
        }

        return paragraph;
    }

    /// <summary>
    /// Creates a Word paragraph for a markdown heading with appropriate style and optional numbering.
    /// </summary>
    /// <param name="heading">The markdown heading block.</param>
    /// <param name="autoNumberHeadings">If true, strips manual numbers and applies automatic numbering.</param>
    /// <param name="useTitleStyle">If true, H1 uses Title style instead of Heading 1.</param>
    private static Paragraph CreateHeadingParagraph(HeadingBlock heading, bool autoNumberHeadings, bool useTitleStyle)
    {
        int level = Math.Clamp(heading.Level, 1, 6);

        // Determine the Word style based on useTitleStyle setting:
        // If useTitleStyle is true:
        //   # (level 1) -> Title
        //   ## (level 2) -> Heading 1
        //   ### (level 3) -> Heading 2, etc.
        // If useTitleStyle is false:
        //   # (level 1) -> Heading 1
        //   ## (level 2) -> Heading 2, etc.
        string styleId;
        int numberingLevel;

        if (useTitleStyle)
        {
            styleId = level == 1 ? StyleIdTitle : $"{StyleIdHeading}{level - 1}";
            numberingLevel = level - 2; // level 2 -> 0, level 3 -> 1, etc.
        }
        else
        {
            styleId = $"{StyleIdHeading}{level}";
            numberingLevel = level - 1; // level 1 -> 0, level 2 -> 1, etc.
        }

        ParagraphProperties props = new(
            new ParagraphStyleId { Val = styleId },
            new SpacingBetweenLines { Before = HeadingSpacingBefore, After = HeadingSpacingAfter }
        );

        // Add automatic numbering if enabled and appropriate
        // For useTitleStyle=true: only H2-H6 get numbered (not Title)
        // For useTitleStyle=false: all H1-H6 get numbered
        bool shouldNumber = autoNumberHeadings && (!useTitleStyle || level > 1);

        if (shouldNumber)
        {
            props.Append(new NumberingProperties(
                new NumberingLevelReference { Val = numberingLevel },
                new NumberingId { Val = 3 } // Numbering ID 3 for heading numbering
            ));
        }

        Paragraph paragraph = new(props);

        // Process inline content to preserve formatting like bold, italic, code, etc.
        if (heading.Inline is not null)
        {
            bool isFirstInline = true;
            foreach (Inline inline in heading.Inline)
            {
                // Strip leading numbers from first inline element if auto-numbering is enabled
                if (autoNumberHeadings && isFirstInline && inline is LiteralInline literal)
                {
                    string text = literal.Content.ToString();
                    string strippedText = System.Text.RegularExpressions.Regex.Replace(text, HeadingNumberPattern, "");
                    if (strippedText != text)
                    {
                        if (!string.IsNullOrWhiteSpace(strippedText))
                        {
                            paragraph.Append(new Run(new Text(strippedText) { Space = SpaceProcessingModeValues.Preserve }));
                        }
                        isFirstInline = false;
                        continue;
                    }
                }

                AppendHeadingInline(paragraph, inline);
                isFirstInline = false;
            }
        }

        return paragraph;
    }

    private static void AppendHeadingInline(Paragraph paragraph, Inline inline)
    {
        switch (inline)
        {
            case LiteralInline literal:
                paragraph.Append(new Run(new Text(literal.Content.ToString()) { Space = SpaceProcessingModeValues.Preserve }));
                break;
            case LineBreakInline:
                paragraph.Append(new Run(new Break()));
                break;
            case CodeInline codeInline:
                paragraph.Append(new Run(
                    new RunProperties(
                        new RunFonts { Ascii = CodeFontName, HighAnsi = CodeFontName },
                        new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = CodeBackgroundGray }
                    ),
                    new Text(codeInline.Content ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve }
                ));
                break;
            case EmphasisInline emphasis:
                AppendHeadingEmphasis(paragraph, emphasis);
                break;
            case ContainerInline container:
                foreach (Inline child in container)
                {
                    AppendHeadingInline(paragraph, child);
                }
                break;
            default:
                string fallbackText = ExtractInlinePlainText(inline);
                if (!string.IsNullOrWhiteSpace(fallbackText))
                {
                    paragraph.Append(new Run(new Text(fallbackText) { Space = SpaceProcessingModeValues.Preserve }));
                }
                break;
        }
    }

    private static void AppendHeadingEmphasis(Paragraph paragraph, EmphasisInline emphasis)
    {
        bool isStrike = emphasis.DelimiterChar == '~';
        bool isBold = !isStrike && emphasis.DelimiterCount >= 2;
        bool isItalic = !isStrike && emphasis.DelimiterCount == 1;

        foreach (Inline child in emphasis)
        {
            string text = ExtractInlinePlainText(child);
            RunProperties props = new();

            if (isBold)
            {
                props.Append(new Bold());
            }

            if (isItalic)
            {
                props.Append(new Italic());
            }

            if (isStrike)
            {
                props.Append(new Strike());
            }

            paragraph.Append(new Run(props, new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        }
    }

    private static Paragraph CreateQuoteParagraph(string text)
    {
        Paragraph paragraph = new(
            new ParagraphProperties(
                new ParagraphStyleId { Val = StyleIdQuote },
                new Indentation { Left = QuoteIndentLeft.ToString() },
                new SpacingBetweenLines { After = QuoteSpacingAfter },
                new ParagraphBorders(
                    new LeftBorder { Val = BorderValues.Single, Size = QuoteBorderWidth, Color = QuoteBorderGray }
                )
            )
        );
        paragraph.Append(CreateTextRunWithBreaks(
            text ?? string.Empty,
            new RunProperties(new Color { Val = QuoteTextGray })
        ));
        return paragraph;
    }

    private static Paragraph CreateHorizontalRule()
    {
        Paragraph paragraph = new(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = ListItemSpacingAfter, After = ListItemSpacingAfter },
                new ParagraphBorders(
                    new BottomBorder 
                    { 
                        Val = BorderValues.Single, 
                        Size = HorizontalRuleBorderWidth, 
                        Color = QuoteBorderGray 
                    }
                )
            ),
            new Run(new Text(string.Empty))
        );
        return paragraph;
    }

    private static void AppendFencedCodeBlock(Body body, FencedCodeBlock fenced)
    {
        string language = fenced.Info?.ToString().Trim() ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(language))
        {
            body.Append(new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId { Val = StyleIdCodeBlock },
                    new SpacingBetweenLines { Before = CodeHeaderSpacingBefore, After = CodeHeaderSpacingAfter },
                    new Indentation { Left = CodeBlockIndentLeft, Right = CodeBlockIndentRight },
                    new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = CodeHeaderBackgroundDark }
                ),
                new Run(
                    new RunProperties(
                        new Bold(),
                        new Color { Val = CodeHeaderForegroundLight },
                        new RunFonts { Ascii = CodeHeaderFontName, HighAnsi = CodeHeaderFontName },
                        new FontSize { Val = CodeHeaderFontSize }
                    ),
                    new Text(language) { Space = SpaceProcessingModeValues.Preserve }
                )
            ));
        }

        body.Append(CreateCodeParagraph(ExtractLinesText(fenced)));
    }

    private static Paragraph CreateCodeParagraph(string code)
    {
        Paragraph paragraph = new(
            new ParagraphProperties(
                new ParagraphStyleId { Val = StyleIdCodeBlock },
                new SpacingBetweenLines { Before = CodeBlockSpacingBefore, After = CodeBlockSpacingAfter, Line = CodeBlockLineSpacing, LineRule = LineSpacingRuleValues.Auto },
                new Indentation { Left = CodeBlockIndentLeft, Right = CodeBlockIndentRight },
                new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = CodeBackgroundGray },
                new ParagraphBorders(
                    new TopBorder { Val = BorderValues.Single, Size = CodeBlockBorderWidth, Color = CodeBlockBorderGray },
                    new BottomBorder { Val = BorderValues.Single, Size = CodeBlockBorderWidth, Color = CodeBlockBorderGray },
                    new LeftBorder { Val = BorderValues.Single, Size = CodeBlockBorderWidth, Color = CodeBlockBorderGray },
                    new RightBorder { Val = BorderValues.Single, Size = CodeBlockBorderWidth, Color = CodeBlockBorderGray }
                )
            )
        );

        paragraph.Append(CreateTextRunWithBreaks(
            code,
            new RunProperties(new RunFonts { Ascii = CodeFontName, HighAnsi = CodeFontName }, new FontSize { Val = CodeFontSize })
        ));

        return paragraph;
    }

    private static Run CreateTextRunWithBreaks(string text, RunProperties? runProperties = null)
    {
        Run run = runProperties is null ? new Run() : new Run((RunProperties)runProperties.CloneNode(true));
        string normalized = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
        string[] lines = normalized.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            run.Append(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
            if (i < lines.Length - 1)
            {
                run.Append(new Break());
            }
        }

        return run;
    }

    /// <summary>
    /// Recursively processes inline elements (bold, italic, code, links) and appends them to a paragraph.
    /// Preserves all inline formatting when converting from markdown to Word.
    /// </summary>
    private static void AppendInlines(Paragraph paragraph, ContainerInline inline, MainDocumentPart mainPart, Dictionary<string, string> hyperlinkRelationshipIds)
    {
        AppendInlines(paragraph, inline, mainPart, hyperlinkRelationshipIds, false);
    }

    private static void AppendInlines(Paragraph paragraph, ContainerInline inline, MainDocumentPart mainPart, Dictionary<string, string> hyperlinkRelationshipIds, bool forceBold)
    {
        foreach (Inline child in inline)
        {
            switch (child)
            {
                case LiteralInline literal:
                    Run literalRun = new(new Text(literal.Content.ToString()) { Space = SpaceProcessingModeValues.Preserve });
                    if (forceBold)
                    {
                        literalRun.RunProperties = new RunProperties(new Bold());
                    }
                    paragraph.Append(literalRun);
                    break;
                case LineBreakInline:
                    paragraph.Append(new Run(new Break()));
                    break;
                case CodeInline codeInline:
                    RunProperties codeProps = new(
                        new RunFonts { Ascii = CodeFontName, HighAnsi = CodeFontName },
                        new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = CodeBackgroundGray }
                    );
                    if (forceBold)
                    {
                        codeProps.Append(new Bold());
                    }
                    paragraph.Append(new Run(
                        codeProps,
                        new Text(codeInline.Content ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve }
                    ));
                    break;
                case EmphasisInline emphasis:
                    AppendEmphasis(paragraph, emphasis, forceBold);
                    break;
                case LinkInline link:
                    AppendLink(paragraph, link, mainPart, hyperlinkRelationshipIds, forceBold);
                    break;
                case ContainerInline container:
                    AppendInlines(paragraph, container, mainPart, hyperlinkRelationshipIds, forceBold);
                    break;
                default:
                    string fallbackText = ExtractInlinePlainText(child);
                    if (!string.IsNullOrWhiteSpace(fallbackText))
                    {
                        Run fallbackRun = new(new Text(fallbackText) { Space = SpaceProcessingModeValues.Preserve });
                        if (forceBold)
                        {
                            fallbackRun.RunProperties = new RunProperties(new Bold());
                        }
                        paragraph.Append(fallbackRun);
                    }
                    break;
            }
        }
    }

    private static void AppendEmphasis(Paragraph paragraph, EmphasisInline emphasis)
    {
        AppendEmphasis(paragraph, emphasis, false);
    }

    private static void AppendEmphasis(Paragraph paragraph, EmphasisInline emphasis, bool forceBold)
    {
        bool isStrike = emphasis.DelimiterChar == '~';
        bool isBold = !isStrike && emphasis.DelimiterCount >= 2;
        bool isItalic = !isStrike && emphasis.DelimiterCount == 1;

        foreach (Inline child in emphasis)
        {
            string text = ExtractInlinePlainText(child);
            RunProperties props = new();
            if (isBold || forceBold)
            {
                props.Append(new Bold());
            }

            if (isItalic)
            {
                props.Append(new Italic());
            }

            if (isStrike)
            {
                props.Append(new Strike());
            }

            paragraph.Append(new Run(props, new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        }
    }

    private static void AppendLink(Paragraph paragraph, LinkInline link, MainDocumentPart mainPart, Dictionary<string, string> hyperlinkRelationshipIds)
    {
        AppendLink(paragraph, link, mainPart, hyperlinkRelationshipIds, false);
    }

    private static void AppendLink(Paragraph paragraph, LinkInline link, MainDocumentPart mainPart, Dictionary<string, string> hyperlinkRelationshipIds, bool forceBold)
    {
        string text = string.IsNullOrWhiteSpace(link.Title)
            ? string.Concat(link.OfType<Inline>().Select(ExtractInlinePlainText))
            : link.Title!;

        if (string.IsNullOrWhiteSpace(text))
        {
            text = link.Url ?? string.Empty;
        }

        string? url = link.Url;
        if (!string.IsNullOrWhiteSpace(url) && Uri.TryCreate(url, UriKind.Absolute, out Uri? uri))
        {
            if (!hyperlinkRelationshipIds.TryGetValue(url, out string? relationshipId))
            {
                HyperlinkRelationship rel = mainPart.AddHyperlinkRelationship(uri, true);
                relationshipId = rel.Id;
                hyperlinkRelationshipIds[url] = relationshipId;
            }

            RunProperties linkProps = new(
                new RunStyle { Val = StyleIdHyperlink },
                new Underline { Val = UnderlineValues.Single },
                new Color { Val = HyperlinkColor }
            );
            if (forceBold)
            {
                linkProps.Append(new Bold());
            }

            Hyperlink hyperlink = new(
                new Run(
                    linkProps,
                    new Text(text) { Space = SpaceProcessingModeValues.Preserve }
                )
            )
            {
                Id = relationshipId,
                History = OnOffValue.FromBoolean(true)
            };

            paragraph.Append(hyperlink);
            return;
        }

        RunProperties plainLinkProps = new(new Underline { Val = UnderlineValues.Single });
        if (forceBold)
        {
            plainLinkProps.Append(new Bold());
        }
        paragraph.Append(new Run(
            plainLinkProps,
            new Text(text) { Space = SpaceProcessingModeValues.Preserve }
        ));
    }

    private static WTable CreateTable(MdTable table)
    {
        return CreateTable(table, null, null);
    }

    /// <summary>
    /// Creates a Word table from a markdown table with borders and formatting.
    /// Preserves inline formatting in cells and applies bold to headers.
    /// </summary>
    private static WTable CreateTable(MdTable table, MainDocumentPart? mainPart, Dictionary<string, string>? hyperlinkRelationshipIds)
    {
        WTable wordTable = new();

        TableProperties properties = new(
            new TableWidth { Width = TableAutoWidthValue, Type = TableWidthUnitValues.Auto },
            new TableLayout { Type = TableLayoutValues.Autofit },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = TableOuterBorderWidth, Color = TableBorderGray },
                new BottomBorder { Val = BorderValues.Single, Size = TableOuterBorderWidth, Color = TableBorderGray },
                new LeftBorder { Val = BorderValues.Single, Size = TableOuterBorderWidth, Color = TableBorderGray },
                new RightBorder { Val = BorderValues.Single, Size = TableOuterBorderWidth, Color = TableBorderGray },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = TableInnerBorderWidth, Color = TableBorderGray },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = TableInnerBorderWidth, Color = TableBorderGray }
            )
        );
        wordTable.Append(properties);

        foreach (MdTableRow row in table.OfType<MdTableRow>())
        {
            WTableRow wordRow = new();
            if (row.IsHeader)
            {
                wordRow.Append(new TableRowProperties(new TableHeader()));
            }

            foreach (MdTableCell cell in row.OfType<MdTableCell>())
            {
                Paragraph paragraph = CreateTableCellParagraph(cell, mainPart, hyperlinkRelationshipIds, row.IsHeader);

                WTableCell wordCell = new(
                    new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Auto, Width = TableAutoWidthValue },
                        new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                    ),
                    paragraph
                );

                wordRow.Append(wordCell);
            }

            wordTable.Append(wordRow);
        }

        return wordTable;
    }

    private static string ExtractPlainText(Block block)
    {
        switch (block)
        {
            case ParagraphBlock paragraph when paragraph.Inline is not null:
                return string.Concat(paragraph.Inline.Select(ExtractInlinePlainText));
            case HeadingBlock heading when heading.Inline is not null:
                return string.Concat(heading.Inline.Select(ExtractInlinePlainText));
            case FencedCodeBlock fenced:
                return ExtractLinesText(fenced);
            case CodeBlock code:
                return ExtractLinesText(code);
            case ListBlock list:
                return string.Join(Environment.NewLine, list.Select(ExtractPlainText));
            case ListItemBlock item:
                return string.Join(" ", item.Select(ExtractPlainText));
            case ContainerBlock container:
                return string.Join(Environment.NewLine, container.Select(ExtractPlainText));
            default:
                if (block is LeafBlock leaf)
                {
                    return ExtractLinesText(leaf);
                }

                return string.Empty;
        }
    }

    private static string ExtractInlinePlainText(Inline inline)
    {
        return inline switch
        {
            LiteralInline literal => literal.Content.ToString(),
            CodeInline code => code.Content,
            LineBreakInline => Environment.NewLine,
            EmphasisInline emphasis => string.Concat(emphasis.Select(ExtractInlinePlainText)),
            LinkInline link => string.Concat(link.Select(ExtractInlinePlainText)),
            ContainerInline container => string.Concat(container.Select(ExtractInlinePlainText)),
            _ => string.Empty
        };
    }

    private static string ExtractLinesText(LeafBlock block)
    {
        if (block.Lines.Lines is null || block.Lines.Count == 0)
        {
            return string.Empty;
        }

        return string.Join(Environment.NewLine, block.Lines.Lines.Select(l => l.ToString()));
    }

    /// <summary>
    /// Creates and configures Word styles (Title, Headings, CodeBlock) for the document.
    /// Styles match the preview appearance with consistent colors and sizing.
    /// </summary>
    private static void EnsureStylesPart(MainDocumentPart mainPart, bool autoNumberHeadings, bool useTitleStyle)
    {
        StyleDefinitionsPart stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();

        bool stylesExist = stylesPart.Styles.Elements<Style>().Any(s => s.StyleId?.Value == "CodeBlock");

        if (!stylesExist)
        {
            // Add title style if using title mode
            if (useTitleStyle)
            {
                stylesPart.Styles.Append(CreateTitleStyle());
            }

            // Add heading styles (now we only need H1-H5 since H6 markdown -> H5 Word)
            for (int i = 1; i <= 5; i++)
            {
                stylesPart.Styles.Append(CreateHeadingStyle(i));
            }

            // Add code block style
            stylesPart.Styles.Append(CreateCodeBlockStyle());

            stylesPart.Styles.Save();
        }
    }

    private static Style CreateHeadingStyle(int level)
    {
        Style style = new() 
        { 
            Type = StyleValues.Paragraph, 
            StyleId = $"Heading{level}",
            Default = false
        };

        style.Append(new StyleName { Val = $"Heading {level}" });
        style.Append(new BasedOn { Val = StyleIdNormal });
        style.Append(new NextParagraphStyle { Val = StyleIdNormal });
        style.Append(new LinkedStyle { Val = $"{StyleIdHeading}{level}Char" });
        style.Append(new UIPriority { Val = HeadingStylePriority });
        style.Append(new PrimaryStyle());

        // Font sizes for different heading levels
        string fontSize = level switch
        {
            1 => Heading1FontSize,
            2 => Heading2FontSize,
            3 => Heading3FontSize,
            4 => Heading4FontSize,
            5 => Heading5FontSize,
            _ => Heading5FontSize
        };

        // Paragraph properties
        StyleParagraphProperties paragraphProps = new();
        paragraphProps.Append(new KeepNext());
        paragraphProps.Append(new KeepLines());
        paragraphProps.Append(new SpacingBetweenLines 
        { 
            Before = HeadingSpacingBefore, 
            After = "0",
            Line = HeadingSpacingBefore,
            LineRule = LineSpacingRuleValues.Auto
        });
        paragraphProps.Append(new OutlineLevel { Val = level - 1 });
        style.Append(paragraphProps);

        // Run properties (character formatting)
        StyleRunProperties runProps = new();
        runProps.Append(new Bold());
        runProps.Append(new FontSize { Val = fontSize });
        runProps.Append(new FontSizeComplexScript { Val = fontSize });

        // Different colors for different heading levels
        string color = level switch
        {
            1 => HeadingColorBlue,
            2 => HeadingColorBlue,
            _ => HeadingColorDarkBlue
        };
        runProps.Append(new Color { Val = color });

        style.Append(runProps);

        return style;
    }

    private static Style CreateTitleStyle()
    {
        Style style = new() 
        { 
            Type = StyleValues.Paragraph, 
            StyleId = "Title",
            Default = false
        };

        style.Append(new StyleName { Val = StyleIdTitle });
        style.Append(new BasedOn { Val = StyleIdNormal });
        style.Append(new NextParagraphStyle { Val = StyleIdNormal });
        style.Append(new LinkedStyle { Val = $"{StyleIdTitle}Char" });
        style.Append(new UIPriority { Val = TitleStylePriority });
        style.Append(new PrimaryStyle());

        // Paragraph properties
        StyleParagraphProperties paragraphProps = new();
        paragraphProps.Append(new SpacingBetweenLines 
        { 
            Before = HeadingSpacingBefore, 
            After = ZeroSpacing,
            Line = HeadingSpacingBefore,
            LineRule = LineSpacingRuleValues.Auto
        });
        paragraphProps.Append(new ContextualSpacing());
        style.Append(paragraphProps);

        // Run properties - Title is typically larger and different color
        StyleRunProperties runProps = new();
        runProps.Append(new Bold());
        runProps.Append(new FontSize { Val = TitleFontSize });
        runProps.Append(new FontSizeComplexScript { Val = TitleFontSize });
        runProps.Append(new Color { Val = HeadingColorBlue });
        runProps.Append(new Spacing { Val = int.Parse(TitleCharacterSpacing) });
        style.Append(runProps);

        return style;
    }

    private static Style CreateCodeBlockStyle()
    {
        Style style = new() { Type = StyleValues.Paragraph, StyleId = StyleIdCodeBlock, CustomStyle = true };
        style.Append(new StyleName { Val = "Code Block" });
        style.Append(new BasedOn { Val = StyleIdNormal });
        style.Append(new NextParagraphStyle { Val = StyleIdNormal });
        style.Append(new StyleRunProperties(
            new RunFonts { Ascii = CodeFontName, HighAnsi = CodeFontName },
            new FontSize { Val = CodeFontSize }
        ));
        return style;
    }

    /// <summary>
    /// Creates numbering definitions for bullets, numbered lists, and optional heading numbering.
    /// Each list instance gets a unique ID to prevent Word from continuing numbering across separate lists.
    /// </summary>
    private static void EnsureNumberingPart(MainDocumentPart mainPart, bool autoNumberHeadings)
    {
        NumberingDefinitionsPart numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();

        if (numberingPart.Numbering != null)
        {
            return; // Already initialized
        }

        // Create numbering definitions
        Numbering numbering = new();

        // Abstract numbering definition for bullet lists (ID 1)
        AbstractNum abstractBullet = new() { AbstractNumberId = 1 };
        abstractBullet.Append(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });

        // Define levels for bullet lists
        for (int i = 0; i < MaxListLevels; i++)
        {
            Level bulletLevel = new() { LevelIndex = i };
            bulletLevel.Append(new NumberingFormat { Val = NumberFormatValues.Bullet });

            // Use proper bullet character - this is the character code for the standard bullet
            string bulletChar = i % 3 == 0 ? BulletCharPrimary : i % 3 == 1 ? BulletCharSecondary : BulletCharTertiary;
            bulletLevel.Append(new LevelText { Val = bulletChar });
            bulletLevel.Append(new LevelJustification { Val = LevelJustificationValues.Left });

            PreviousParagraphProperties bulletPpr = new();
            bulletPpr.Append(new Indentation 
            { 
                Left = (BaseListIndent + (IndentPerLevel * i)).ToString(), 
                Hanging = ListHangingIndent 
            });
            bulletLevel.Append(bulletPpr);

            NumberingSymbolRunProperties bulletRpr = new();
            // For proper bullet display, use Symbol font for \uF0B7 and \uF0A7, Courier New for 'o'
            string fontName = bulletChar == BulletCharSecondary ? BulletFontCourier : BulletFontSymbol;
            bulletRpr.Append(new RunFonts { Hint = FontTypeHintValues.Default, Ascii = fontName, HighAnsi = fontName });
            bulletLevel.Append(bulletRpr);

            abstractBullet.Append(bulletLevel);
        }

        // Abstract numbering definition for numbered lists (ID 2)
        AbstractNum abstractNumbered = new() { AbstractNumberId = 2 };
        abstractNumbered.Append(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });

        // Define levels for numbered lists
        for (int i = 0; i < MaxListLevels; i++)
        {
            Level numberedLevel = new() { LevelIndex = i };
            numberedLevel.Append(new StartNumberingValue { Val = StartNumberValue });
            numberedLevel.Append(new NumberingFormat { Val = NumberFormatValues.Decimal });
            numberedLevel.Append(new LevelText { Val = $"%{i + 1}." });
            numberedLevel.Append(new LevelJustification { Val = LevelJustificationValues.Left });

            PreviousParagraphProperties numberedPpr = new();
            numberedPpr.Append(new Indentation 
            { 
                Left = (BaseListIndent + (IndentPerLevel * i)).ToString(), 
                Hanging = ListHangingIndent 
            });
            numberedLevel.Append(numberedPpr);

            abstractNumbered.Append(numberedLevel);
        }

        // Add abstract numbering definitions
        numbering.Append(abstractBullet);
        numbering.Append(abstractNumbered);

        // Create numbering instances that reference the abstract definitions
        NumberingInstance bulletInstance = new() { NumberID = 1 };
        bulletInstance.Append(new AbstractNumId { Val = 1 });

        NumberingInstance numberedInstance = new() { NumberID = 2 };
        numberedInstance.Append(new AbstractNumId { Val = 2 });

        numbering.Append(bulletInstance);
        numbering.Append(numberedInstance);

        // Add heading numbering if auto-numbering is enabled
        if (autoNumberHeadings)
        {
            AbstractNum abstractHeading = new() { AbstractNumberId = 3 };
            abstractHeading.Append(new MultiLevelType { Val = MultiLevelValues.Multilevel });

            // Define levels for headings (H1-H5)
            for (int i = 0; i < MaxHeadingLevels; i++)
            {
                Level headingLevel = new() { LevelIndex = i };
                headingLevel.Append(new StartNumberingValue { Val = StartNumberValue });
                headingLevel.Append(new NumberingFormat { Val = NumberFormatValues.Decimal });

                // Build level text: %1. for H1, %1.%2. for H2, %1.%2.%3. for H3, etc.
                string levelText = string.Join(".", Enumerable.Range(1, i + 1).Select(n => $"%{n}")) + ".";
                headingLevel.Append(new LevelText { Val = levelText });
                headingLevel.Append(new LevelJustification { Val = LevelJustificationValues.Left });

                PreviousParagraphProperties headingPpr = new();
                headingPpr.Append(new Indentation 
                { 
                    Left = ZeroSpacing, // No indent for headings
                    Hanging = ZeroSpacing 
                });
                headingLevel.Append(headingPpr);

                abstractHeading.Append(headingLevel);
            }

            numbering.Append(abstractHeading);

            NumberingInstance headingInstance = new() { NumberID = 3 };
            headingInstance.Append(new AbstractNumId { Val = 3 });
            numbering.Append(headingInstance);
        }

        numberingPart.Numbering = numbering;
        numberingPart.Numbering.Save();
    }
}
