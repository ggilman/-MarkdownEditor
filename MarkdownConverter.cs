using Markdig;

namespace MarkdownUtilsApp;

internal static class MarkdownConverter
{
    private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder()
        .UseAdvancedExtensions()
        .DisableHtml()
        .Build();

    public static string ToHtml(string markdown)
    {
        return Markdown.ToHtml(markdown ?? string.Empty, Pipeline);
    }
}
