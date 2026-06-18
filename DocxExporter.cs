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

namespace MarkdownUtilsApp;

internal static class DocxExporter
{
    private const int PageContentWidthTwips = 9360;
    private const int BaseNumberingIdForLists = 10; // Start list numbering IDs at 10 to avoid conflicts

    private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder()
        .UseAdvancedExtensions()
        .DisableHtml()
        .Build();

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
            new PageSize { Width = 12240, Height = 15840 },
            new PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 }
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
            // Determine which abstract numbering definition to use
            // AbstractNum 1 = bullets, AbstractNum 2 = numbered
            int abstractNumId = list.IsOrdered ? 2 : 1;

            // Create a new numbering instance for this specific list
            numberingId = BaseNumberingIdForLists + listNumberingCounter++;
            NumberingInstance newInstance = new() { NumberID = numberingId };
            newInstance.Append(new AbstractNumId { Val = abstractNumId });
            numberingPart.Numbering.Append(newInstance);
            numberingPart.Numbering.Save();
        }
        else
        {
            // Fallback: use the old static numbering IDs if numbering part isn't available yet
            numberingId = list.IsOrdered ? 2 : 1;
        }

        // Debug: Log what type of list we're processing
        System.Diagnostics.Debug.WriteLine($"Processing list: IsOrdered={list.IsOrdered}, BulletType={list.BulletType}, OrderedStart={list.OrderedStart}, NumberingId={numberingId}");

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
                        // For task lists, use manual checkbox markers since Word doesn't have native task list support
                        string marker = taskState.Value ? "☑ " : "☐ ";
                        listParagraph = CreateListItemParagraph(
                            marker,
                            paragraphBlock.Inline,
                            mainPart,
                            hyperlinkRelationshipIds,
                            indentLeftTwips: 360 * (depth + 1),
                            indentHangingTwips: 240
                        );
                    }
                    else if (isFirstParagraphInItem)
                    {
                        // First paragraph in list item uses Word's native numbering
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
                        // Continuation paragraphs in multi-paragraph list items
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
                            string marker = taskState.Value ? "☑ " : "☐ ";
                            body.Append(CreateParagraphFromText(marker + lineText, indentLeftTwips: 360 * (depth + 1), indentHangingTwips: 240));
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

                PropertyInfo? checkedProperty = inline.GetType().GetProperty("Checked") ?? inline.GetType().GetProperty("IsChecked");
                if (checkedProperty?.PropertyType == typeof(bool))
                {
                    return (bool)checkedProperty.GetValue(inline)!;
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
            new SpacingBetweenLines { After = "200", Line = "360", LineRule = LineSpacingRuleValues.Auto }
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
            new SpacingBetweenLines { After = "120", Line = "360", LineRule = LineSpacingRuleValues.Auto },
            new Indentation 
            { 
                Left = indentLeftTwips.ToString(), 
                Hanging = indentHangingTwips.ToString() 
            }
        );

        paragraph.Append(props);

        // Add the bullet/number marker as plain text
        paragraph.Append(new Run(new Text(marker) { Space = SpaceProcessingModeValues.Preserve }));

        // Add the inline formatted content
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
            new SpacingBetweenLines { After = "120", Line = "360", LineRule = LineSpacingRuleValues.Auto },
            new NumberingProperties(
                new NumberingLevelReference { Val = levelIndex },
                new NumberingId { Val = numberingId }
            )
        );

        paragraph.Append(props);

        // Add the inline formatted content
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
            new SpacingBetweenLines { After = "120", Line = "360", LineRule = LineSpacingRuleValues.Auto },
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
            styleId = level == 1 ? "Title" : $"Heading{level - 1}";
            numberingLevel = level - 2; // level 2 -> 0, level 3 -> 1, etc.
        }
        else
        {
            styleId = $"Heading{level}";
            numberingLevel = level - 1; // level 1 -> 0, level 2 -> 1, etc.
        }

        ParagraphProperties props = new(
            new ParagraphStyleId { Val = styleId },
            new SpacingBetweenLines { Before = "240", After = "120" }
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
                    string strippedText = System.Text.RegularExpressions.Regex.Replace(text, @"^\s*\d+(\.\d+)*\.?\s*", "");
                    if (strippedText != text && !string.IsNullOrWhiteSpace(strippedText))
                    {
                        paragraph.Append(new Run(new Text(strippedText) { Space = SpaceProcessingModeValues.Preserve }));
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
                        new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" },
                        new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" }
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
                new ParagraphStyleId { Val = "Quote" },
                new Indentation { Left = "720" },
                new SpacingBetweenLines { After = "120" },
                new ParagraphBorders(
                    new LeftBorder { Val = BorderValues.Single, Size = 24, Color = "999999" }
                )
            )
        );
        paragraph.Append(CreateTextRunWithBreaks(
            text ?? string.Empty,
            new RunProperties(new Color { Val = "444444" })
        ));
        return paragraph;
    }

    private static Paragraph CreateHorizontalRule()
    {
        Paragraph paragraph = new(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "120", After = "120" },
                new ParagraphBorders(
                    new BottomBorder 
                    { 
                        Val = BorderValues.Single, 
                        Size = 6, 
                        Color = "999999" 
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
                    new ParagraphStyleId { Val = "CodeBlock" },
                    new SpacingBetweenLines { Before = "120", After = "40" },
                    new Indentation { Left = "360", Right = "120" },
                    new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "333333" }
                ),
                new Run(
                    new RunProperties(
                        new Bold(),
                        new Color { Val = "F2F2F2" },
                        new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                        new FontSize { Val = "18" }
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
                new ParagraphStyleId { Val = "CodeBlock" },
                new SpacingBetweenLines { Before = "120", After = "120", Line = "276", LineRule = LineSpacingRuleValues.Auto },
                new Indentation { Left = "360", Right = "120" },
                new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" },
                new ParagraphBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Color = "D8D8D8" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "D8D8D8" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "D8D8D8" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Color = "D8D8D8" }
                )
            )
        );

        paragraph.Append(CreateTextRunWithBreaks(
            code,
            new RunProperties(new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" }, new FontSize { Val = "20" })
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
                        new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" },
                        new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F3F3F3" }
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
                new RunStyle { Val = "Hyperlink" },
                new Underline { Val = UnderlineValues.Single },
                new Color { Val = "0563C1" }
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

    private static WTable CreateTable(MdTable table, MainDocumentPart? mainPart, Dictionary<string, string>? hyperlinkRelationshipIds)
    {
        WTable wordTable = new();

        TableProperties properties = new(
            new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto },
            new TableLayout { Type = TableLayoutValues.Autofit },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 8, Color = "CCCCCC" },
                new BottomBorder { Val = BorderValues.Single, Size = 8, Color = "CCCCCC" },
                new LeftBorder { Val = BorderValues.Single, Size = 8, Color = "CCCCCC" },
                new RightBorder { Val = BorderValues.Single, Size = 8, Color = "CCCCCC" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6, Color = "CCCCCC" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 6, Color = "CCCCCC" }
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
                        new TableCellWidth { Type = TableWidthUnitValues.Auto, Width = "0" },
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
        style.Append(new BasedOn { Val = "Normal" });
        style.Append(new NextParagraphStyle { Val = "Normal" });
        style.Append(new LinkedStyle { Val = $"Heading{level}Char" });
        style.Append(new UIPriority { Val = 9 });
        style.Append(new PrimaryStyle());

        // Font sizes for different heading levels
        string fontSize = level switch
        {
            1 => "32", // 16pt
            2 => "26", // 13pt
            3 => "24", // 12pt
            4 => "22", // 11pt
            5 => "20", // 10pt
            _ => "20"  // 10pt
        };

        // Paragraph properties
        StyleParagraphProperties paragraphProps = new();
        paragraphProps.Append(new KeepNext());
        paragraphProps.Append(new KeepLines());
        paragraphProps.Append(new SpacingBetweenLines 
        { 
            Before = "240", 
            After = "0",
            Line = "240",
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
            1 => "2E74B5", // Dark blue
            2 => "2E74B5", // Dark blue
            _ => "1F4D78"  // Darker blue for H3-H6
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

        style.Append(new StyleName { Val = "Title" });
        style.Append(new BasedOn { Val = "Normal" });
        style.Append(new NextParagraphStyle { Val = "Normal" });
        style.Append(new LinkedStyle { Val = "TitleChar" });
        style.Append(new UIPriority { Val = 10 });
        style.Append(new PrimaryStyle());

        // Paragraph properties
        StyleParagraphProperties paragraphProps = new();
        paragraphProps.Append(new SpacingBetweenLines 
        { 
            Before = "240", 
            After = "0",
            Line = "240",
            LineRule = LineSpacingRuleValues.Auto
        });
        paragraphProps.Append(new ContextualSpacing());
        style.Append(paragraphProps);

        // Run properties - Title is typically larger and different color
        StyleRunProperties runProps = new();
        runProps.Append(new Bold());
        runProps.Append(new FontSize { Val = "52" }); // 26pt - larger than H1
        runProps.Append(new FontSizeComplexScript { Val = "52" });
        runProps.Append(new Color { Val = "2E74B5" }); // Dark blue
        runProps.Append(new Spacing { Val = -10 }); // Slight character spacing
        style.Append(runProps);

        return style;
    }

    private static Style CreateCodeBlockStyle()
    {
        Style style = new() { Type = StyleValues.Paragraph, StyleId = "CodeBlock", CustomStyle = true };
        style.Append(new StyleName { Val = "Code Block" });
        style.Append(new BasedOn { Val = "Normal" });
        style.Append(new NextParagraphStyle { Val = "Normal" });
        style.Append(new StyleRunProperties(
            new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" },
            new FontSize { Val = "20" }
        ));
        return style;
    }

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

        // Define 9 levels for bullet lists
        for (int i = 0; i < 9; i++)
        {
            Level bulletLevel = new() { LevelIndex = i };
            bulletLevel.Append(new NumberingFormat { Val = NumberFormatValues.Bullet });

            // Use proper bullet character - this is the character code for the standard bullet
            string bulletChar = i % 3 == 0 ? "\uF0B7" : i % 3 == 1 ? "o" : "\uF0A7";
            bulletLevel.Append(new LevelText { Val = bulletChar });
            bulletLevel.Append(new LevelJustification { Val = LevelJustificationValues.Left });

            PreviousParagraphProperties bulletPpr = new();
            bulletPpr.Append(new Indentation 
            { 
                Left = (720 + (360 * i)).ToString(), 
                Hanging = "360" 
            });
            bulletLevel.Append(bulletPpr);

            NumberingSymbolRunProperties bulletRpr = new();
            // For proper bullet display, use Symbol font for \uF0B7 and \uF0A7, Courier New for 'o'
            string fontName = bulletChar == "o" ? "Courier New" : "Symbol";
            bulletRpr.Append(new RunFonts { Hint = FontTypeHintValues.Default, Ascii = fontName, HighAnsi = fontName });
            bulletLevel.Append(bulletRpr);

            abstractBullet.Append(bulletLevel);
        }

        // Abstract numbering definition for numbered lists (ID 2)
        AbstractNum abstractNumbered = new() { AbstractNumberId = 2 };
        abstractNumbered.Append(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });

        // Define 9 levels for numbered lists
        for (int i = 0; i < 9; i++)
        {
            Level numberedLevel = new() { LevelIndex = i };
            numberedLevel.Append(new StartNumberingValue { Val = 1 });
            numberedLevel.Append(new NumberingFormat { Val = NumberFormatValues.Decimal });
            numberedLevel.Append(new LevelText { Val = $"%{i + 1}." });
            numberedLevel.Append(new LevelJustification { Val = LevelJustificationValues.Left });

            PreviousParagraphProperties numberedPpr = new();
            numberedPpr.Append(new Indentation 
            { 
                Left = (720 + (360 * i)).ToString(), 
                Hanging = "360" 
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

            // Define 5 levels for headings (H1-H5)
            for (int i = 0; i < 5; i++)
            {
                Level headingLevel = new() { LevelIndex = i };
                headingLevel.Append(new StartNumberingValue { Val = 1 });
                headingLevel.Append(new NumberingFormat { Val = NumberFormatValues.Decimal });

                // Build level text: %1. for H1, %1.%2. for H2, %1.%2.%3. for H3, etc.
                string levelText = string.Join(".", Enumerable.Range(1, i + 1).Select(n => $"%{n}")) + ".";
                headingLevel.Append(new LevelText { Val = levelText });
                headingLevel.Append(new LevelJustification { Val = LevelJustificationValues.Left });

                PreviousParagraphProperties headingPpr = new();
                headingPpr.Append(new Indentation 
                { 
                    Left = (0).ToString(), // No indent for headings
                    Hanging = "0" 
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
