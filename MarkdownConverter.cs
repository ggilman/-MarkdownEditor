using Markdig;
using Markdig.Extensions.Mathematics;

namespace MarkdownEditor;

/// <summary>
/// Converts Markdown text to HTML for preview rendering.
/// Uses Markdig pipeline with advanced extensions and LaTeX math support.
/// </summary>
internal static class MarkdownConverter
{
    private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder()
        .UseAdvancedExtensions()
        .UseMathematics()
        .DisableHtml()
        .Build();

    /// <summary>
    /// Converts markdown text to HTML.
    /// </summary>
    /// <param name="markdown">The markdown text to convert.</param>
    /// <returns>HTML string ready for display in WebView2.</returns>
    public static string ToHtml(string markdown)
    {
        return Markdown.ToHtml(markdown ?? string.Empty, Pipeline);
    }
}
