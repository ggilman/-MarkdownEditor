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

namespace MarkdownUtilsApp;

internal static class DocxExporter
{
    private const int PageContentWidthTwips = 9360;

    private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder()
        .UseAdvancedExtensions()
        .DisableHtml()
        .Build();

    public static void ExportFromMarkdown(string markdown, string outputPath)
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
        EnsureStylesPart(mainPart);
        Body body = mainPart.Document.Body!;
        Dictionary<string, string> hyperlinkRelationshipIds = new(StringComparer.OrdinalIgnoreCase);

        foreach (Block block in document)
        {
            AppendBlock(body, block, mainPart, hyperlinkRelationshipIds);
        }

        body.Append(new SectionProperties(
            new PageSize { Width = 12240, Height = 15840 },
            new PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 }
        ));

        mainPart.Document.Save();
    }

    private static void AppendBlock(Body body, Block block, MainDocumentPart mainPart, Dictionary<string, string> hyperlinkRelationshipIds)
    {
        switch (block)
        {
            case HeadingBlock heading:
                body.Append(CreateHeadingParagraph(heading));
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
                AppendList(body, list, 0, mainPart, hyperlinkRelationshipIds);
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
                        AppendBlock(body, child, mainPart, hyperlinkRelationshipIds);
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
        AppendList(body, list, 0);
    }

    private static void AppendList(Body body, ListBlock list, int depth)
    {
        AppendList(body, list, depth, null, null);
    }

    private static void AppendList(Body body, ListBlock list, int depth, MainDocumentPart? mainPart, Dictionary<string, string>? hyperlinkRelationshipIds)
    {
        int index = 1;
        foreach (ListItemBlock item in list.OfType<ListItemBlock>())
        {
            string marker = GetListMarker(list, item, index);
            bool addedAny = false;
            foreach (Block child in item)
            {
                if (child is ListBlock nestedList)
                {
                    AppendList(body, nestedList, depth + 1, mainPart, hyperlinkRelationshipIds);
                    continue;
                }

                if (child is ParagraphBlock paragraphBlock && paragraphBlock.Inline is not null)
                {
                    Paragraph listParagraph = CreateListItemParagraph(
                        marker,
                        paragraphBlock.Inline,
                        mainPart,
                        hyperlinkRelationshipIds,
                        indentLeftTwips: 360 * (depth + 1),
                        indentHangingTwips: 240
                    );
                    body.Append(listParagraph);
                    marker = list.IsOrdered ? "   " : "  ";
                    addedAny = true;
                }
                else
                {
                    string lineText = ExtractPlainText(child);
                    if (string.IsNullOrWhiteSpace(lineText))
                    {
                        continue;
                    }

                    body.Append(CreateParagraphFromText(marker + lineText, indentLeftTwips: 360 * (depth + 1), indentHangingTwips: 240));
                    marker = list.IsOrdered ? "   " : "  ";
                    addedAny = true;
                }
            }

            if (!addedAny)
            {
                body.Append(CreateParagraphFromText(marker.TrimEnd(), indentLeftTwips: 360 * (depth + 1), indentHangingTwips: 240));
            }

            index++;
        }
    }

    private static string GetListMarker(ListBlock list, ListItemBlock item, int index)
    {
        bool? taskState = TryGetTaskState(item);
        if (taskState.HasValue)
        {
            return taskState.Value ? "☑ " : "☐ ";
        }

        return list.IsOrdered ? $"{index}. " : "• ";
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

    private static Paragraph CreateHeadingParagraph(HeadingBlock heading)
    {
        int level = Math.Clamp(heading.Level, 1, 6);
        string text = heading.Inline is null ? string.Empty : string.Concat(heading.Inline.Select(ExtractInlinePlainText));

        string size = level switch
        {
            1 => "36",
            2 => "30",
            3 => "26",
            4 => "24",
            5 => "22",
            _ => "20"
        };

        Paragraph paragraph = new(
            new ParagraphProperties(
                new ParagraphStyleId { Val = $"Heading{level}" },
                new SpacingBetweenLines { Before = "220", After = "120" },
                new OutlineLevel { Val = level - 1 }
            )
        );

        paragraph.Append(CreateTextRunWithBreaks(
            text,
            new RunProperties(new Bold(), new FontSize { Val = size })
        ));

        return paragraph;
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

    private static void EnsureStylesPart(MainDocumentPart mainPart)
    {
        StyleDefinitionsPart stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();

        if (stylesPart.Styles.Elements<Style>().Any(s => s.StyleId?.Value == "CodeBlock"))
        {
            return;
        }

        stylesPart.Styles.Append(CreateCodeBlockStyle());
        stylesPart.Styles.Save();
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
}
