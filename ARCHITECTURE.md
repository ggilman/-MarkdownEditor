# Markdown Editor - Technical Architecture

## Overview

Markdown Editor is a WPF desktop application built on .NET 9 that provides a rich Markdown editing experience with live preview and high-fidelity Word export capabilities.

## Architecture Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                         MainWindow                          │
│  ┌──────────────────┐           ┌──────────────────────┐    │
│  │  Markdown Editor │           │  Formatted Preview   │    │
│  │   (TextBox)      │  ◄────►   │   (WebBrowser)       │    │
│  │                  │           │                      │    │
│  │  - Toolbar       │           │  - HTML Rendering    │    │
│  │  - Shortcuts     │           │  - CSS Styling       │    │
│  └──────────────────┘           └──────────────────────┘    │
│           │                                 │               │
│           │                                 │               │
└───────────┼─────────────────────────────────┼───────────────┘
            │                                 │
            │                                 │
    ┌───────┴──────┐               ┌──────────┴────────┐
    │  File I/O    │               │ MarkdownConverter │
    │  - Save .md  │               │  (Markdig)        │
    │  - Import .md│               │  - Parse MD       │
    │  - Export    │               │  - Generate HTML  │
    └───────┬──────┘               └───────────────────┘
            │
            │
    ┌───────┴────────────────────────────────┐
    │         DocxExporter                   │
    │  ┌──────────────────────────────────┐  │
    │  │  Markdig Parsing                 │  │
    │  │  - Parse markdown to AST         │  │
    │  │  - Block elements (H, P, List)   │  │
    │  │  - Inline elements (Bold, Code)  │  │
    │  └─────────────────┬────────────────┘  │
    │                    │                   │
    │  ┌─────────────────┴────────────────┐  │
    │  │  OpenXML Generation              │  │
    │  │  - Document structure            │  │
    │  │  - Styles (Title, Headings)      │  │
    │  │  - Numbering (Bullets, Numbers)  │  │
    │  │  - Paragraphs & Runs             │  │
    │  └─────────────────┬────────────────┘  │
    │                    │                   │
    └────────────────────┼───────────────────┘
                         │
                 ┌───────┴──────┐
                 │  .docx File  │
                 └──────────────┘
```

## Component Details

### MainWindow (UI Layer)

**File**: `MainWindow.xaml` + `MainWindow.xaml.cs`

**Responsibilities**:
- User interface layout and rendering
- User input handling (keyboard, mouse, toolbar clicks)
- File operations (new, save, import, export)
- View state management (editor/preview visibility)
- Live preview updates via `MarkdownConverter`

**Key Methods**:
- `EditorTextBox_TextChanged`: Triggers preview rendering on every keystroke
- `EditorTextBox_PreviewKeyDown`: Handles keyboard shortcuts (Ctrl+B, Ctrl+I, etc.)
- `ExportButton_Click`: Initiates Word export with selected options
- `UpdateViewVisibility`: Manages which panels are visible
- `WrapSelection`: Helper for wrapping text with markdown syntax
- `PrefixSelectedLines`: Helper for adding prefixes (headings, quotes, etc.)

**UI Elements**:
- Menu bar with View options
- Toolbar with formatting buttons
- Split view with editor (left) and preview (right)
- GridSplitter for resizing panels
- Status bar for feedback messages

### MarkdownConverter (Preview Rendering)

**File**: `MarkdownConverter.cs`

**Responsibilities**:
- Convert Markdown to HTML for preview
- Apply consistent CSS styling
- Handle Markdig advanced extensions

**Pipeline Configuration**:
```csharp
var pipeline = new MarkdownPipelineBuilder()
    .UseAdvancedExtensions()  // Tables, task lists, etc.
    .DisableHtml()            // Sanitize raw HTML
    .Build();
```

**CSS Styling**:
- Matches Word export appearance
- Syntax highlighting for code blocks
- Table borders and spacing
- Consistent fonts and colors

### DocxExporter (Word Export Engine)

**File**: `DocxExporter.cs`

**Responsibilities**:
- Parse Markdown using Markdig
- Generate OpenXML document structure
- Create and apply Word styles
- Create and manage numbering definitions
- Preserve inline formatting in complex structures

#### Export Pipeline

```
Markdown Text
     │
     ▼
[Markdig Parse]
     │
     ▼
Markdown AST (Abstract Syntax Tree)
     │
     ▼
[Traverse Blocks]
     │
     ├─► HeadingBlock      → CreateHeadingParagraph
     ├─► ParagraphBlock    → CreateParagraph
     ├─► ListBlock         → AppendList
     ├─► TableBlock        → CreateTable
     ├─► CodeBlock         → CreateCodeParagraph
     ├─► QuoteBlock        → CreateQuoteParagraph
     └─► ThematicBreak     → CreateHorizontalRule
     │
     ▼
OpenXML Elements (Paragraph, Run, Table)
     │
     ▼
.docx File
```

#### Key Methods

**Document Setup**:
- `ExportFromMarkdown`: Main entry point, orchestrates export
- `EnsureStylesPart`: Creates Word styles (Title, Heading1-5, CodeBlock)
- `EnsureNumberingPart`: Creates numbering definitions (bullets, numbers, heading numbering)

**Block Processing**:
- `AppendBlock`: Routes block types to appropriate handlers
- `CreateHeadingParagraph`: Generates heading with proper style and optional numbering
- `CreateParagraph`: Creates paragraph with inline formatting
- `AppendList`: Processes lists with unique numbering instances
- `CreateTable`: Generates tables with borders and formatting
- `CreateCodeParagraph`: Creates code blocks with monospace font
- `CreateQuoteParagraph`: Creates blockquotes with borders
- `CreateHorizontalRule`: Inserts horizontal line separator

**Inline Processing**:
- `AppendInlines`: Recursively processes inline elements
- `AppendHeadingInline`: Special inline handling for headings
- Handles: Bold, Italic, Code, Links, Line breaks, Plain text

**Style Creation**:
- `CreateTitleStyle`: Large, bold, colored title
- `CreateHeadingStyle(level)`: Hierarchical heading styles with colors and sizing
- `CreateCodeBlockStyle`: Monospace font for code

**Numbering System**:
- **AbstractNum 1**: Bullet list definition (Symbol font, bullet characters)
- **AbstractNum 2**: Numbered list definition (decimal numbering)
- **AbstractNum 3**: Heading numbering (hierarchical: 1, 1.1, 1.1.1)
- **NumberingInstances**: Unique instance per list (prevents continuation)

#### List Numbering Architecture

**Problem**: Multiple lists sharing the same NumberingId causes Word to continue numbering.

**Solution**: 
```csharp
// Each list gets unique NumberingId (10, 11, 12, ...)
int newNumberingId = BaseNumberingIdForLists + listNumberingCounter++;

// Instance references the appropriate AbstractNum (1 or 2)
NumberingInstance instance = new() { NumberID = newNumberingId };
instance.Append(new AbstractNumId { Val = list.IsOrdered ? 2 : 1 });
```

This ensures each list is independent.

#### Inline Formatting Preservation

**Challenge**: Lists and tables contain rich inline content (bold, code, links).

**Solution**: Instead of extracting plain text, traverse the inline AST:

```csharp
// OLD (loses formatting):
string text = ExtractPlainText(block);

// NEW (preserves formatting):
AppendInlines(paragraph, inline, mainPart, hyperlinkRelationshipIds);
```

The `AppendInlines` method recursively walks the inline tree and creates appropriate Run elements with formatting properties.

## Data Flow

### Editing Flow
```
User types → TextBox.TextChanged
    │
RenderPreview()
    │
MarkdownConverter.ConvertToHtml()
    │
WebBrowser.NavigateToString(html)
```

### Export Flow
```
User clicks Export → ExportButton_Click
    │
Read checkbox options (autoNumber, useTitleStyle)
    │
ShowSaveFileDialog (suggests filename from first heading)
    │
DocxExporter.ExportFromMarkdown(markdown, path, options)
    │
Markdig.Parse(markdown) → AST
    │
EnsureStylesPart + EnsureNumberingPart
    │
For each block in AST:
    AppendBlock → creates OpenXML elements
    │
Document.Save()
```

## Configuration & Options

### Auto-number headings

**When enabled**:
1. Heading text is scanned for leading numbers: `## 1.2.3 Title`
2. Numbers are stripped via regex: `^\s*\d+(\.\d+)*\.?\s*`
3. NumberingProperties added to paragraph with NumberingId=3
4. Word applies hierarchical numbering automatically

**Implementation**:
```csharp
if (autoNumberHeadings && isFirstInline && inline is LiteralInline literal)
{
    string text = literal.Content.ToString();
    string stripped = Regex.Replace(text, @"^\s*\d+(\.\d+)*\.?\s*", "");
    // Use stripped text
}
```

### # as Title

**When enabled**:
- Markdown level 1 → Word "Title" style (no numbering)
- Markdown level 2+ → Word "Heading 1-5" styles
- Numbering starts at level 2 (numberingLevel = level - 2)

**When disabled**:
- Markdown level 1 → Word "Heading 1" style
- Markdown level 2+ → Word "Heading 2-6" styles  
- Numbering starts at level 1 (numberingLevel = level - 1)

**Implementation**:
```csharp
if (useTitleStyle)
{
    styleId = level == 1 ? "Title" : $"Heading{level - 1}";
    numberingLevel = level - 2;
}
else
{
    styleId = $"Heading{level}";
    numberingLevel = level - 1;
}
```

## Styling System

### Word Style Hierarchy

```
Normal (base)
├─ Title (if useTitleStyle)
├─ Heading 1
├─ Heading 2
├─ Heading 3
├─ Heading 4
├─ Heading 5
└─ CodeBlock (custom)
```

### Style Properties

**Title**:
- Font Size: 26pt (52 half-points)
- Bold: Yes
- Color: #2E74B5 (dark blue)
- Character Spacing: -10 (slightly condensed)

**Heading 1**:
- Font Size: 16pt (32 half-points)
- Bold: Yes
- Color: #2E74B5
- Outline Level: 0

**Heading 2**:
- Font Size: 13pt (26 half-points)
- Bold: Yes
- Color: #2E74B5
- Outline Level: 1

**Heading 3-5**:
- Progressively smaller sizes (12pt, 11pt, 10pt)
- Bold: Yes
- Color: #1F4D78 (darker blue)
- Outline Levels: 2, 3, 4

**CodeBlock**:
- Font: Consolas (monospace)
- Font Size: 10pt (20 half-points)
- Based on: Normal

### Spacing & Indentation

**Headings**:
- Before: 240 twips (12pt)
- After: 120 twips (6pt)
- Line spacing: Auto

**Paragraphs**:
- After: 200 twips (10pt)
- Line spacing: 360 twips (18pt)

**Lists**:
- After: 120 twips (6pt)
- Left indent: 720 + (360 * depth) twips
- Hanging: 360 twips

## Performance Considerations

### Preview Rendering
- Updates on every keystroke (TextChanged event)
- Markdig parsing is fast (< 10ms for typical documents)
- HTML rendering in WebBrowser is efficient
- No throttling currently implemented

### Export Performance
- Single-pass document traversal
- Styles and numbering created once at start
- Efficient inline processing without backtracking
- Typical export: < 100ms for medium documents

### Memory Management
- WordprocessingDocument disposed via `using` statement
- No long-lived document references
- Markdown AST is GC-eligible after export completes

## Error Handling

### File Operations
- Directory creation for save paths
- File deletion before export (overwrite)
- Exception handling with MessageBox feedback
- Status bar updates for success/failure

### Edge Cases Handled
- Empty markdown documents
- Missing inline content
- Nested lists (up to 9 levels)
- Very long headings (> 200 chars in filename)
- Invalid filename characters (sanitized)
- Task lists (checkbox rendering)
- Tables without headers
- Code blocks without language tags

## Testing Strategies

### Manual Testing
- Default sample document covers major features
- Visual comparison between preview and Word export
- Testing all toolbar buttons and shortcuts
- Verifying export options (both checkbox states)
- View menu toggle combinations

### Test Cases
1. Empty document export
2. Heading-only document (numbering test)
3. Complex nested lists
4. Tables with inline formatting
5. Code blocks (fenced and indented)
6. Mixed ordered/unordered lists
7. Task lists with checkboxes
8. Links and images
9. Blockquotes with multiple paragraphs
10. All inline formatting combinations

## Future Enhancements

### Potential Features
- Export to PDF
- Import from Word
- Custom CSS for preview
- Markdown templates
- Find/Replace
- Spell checking
- Document statistics (word count, reading time)
- Export style customization
- Saved export preferences
- Recent files list
- Drag-and-drop file import

### Technical Improvements
- Preview rendering throttling/debouncing
- Async file operations
- Undo/Redo stack beyond TextBox default
- Syntax highlighting in editor
- Line numbers
- Markdown syntax validation
- Performance profiling for large documents
- Unit tests for DocxExporter
- Integration tests for full workflow

## Dependencies

### NuGet Packages

**DocumentFormat.OpenXml**
- Version: Latest stable
- Purpose: Word document generation
- License: MIT
- Alternatives: None (standard for .docx manipulation)

**Markdig**
- Version: Latest stable  
- Purpose: Markdown parsing
- License: BSD-2-Clause
- Features used:
  - Advanced extensions (tables, task lists, strikethrough)
  - AST traversal
  - Inline element parsing

## Build & Deployment

### Build Configuration
- Target Framework: net9.0-windows
- Output Type: WinExe (Windows application)
- Platform: x64
- Configuration: Debug / Release

### Deployment
- Self-contained: No (requires .NET 9 runtime)
- Single-file: Optional
- Trimming: Not enabled (uses reflection in OpenXML)

### Required Runtime
- .NET 9 Desktop Runtime
- Windows 10 or later
- WebBrowser control requires IE11 rendering engine

## Code Quality

### Standards
- C# 12 language features
- Nullable reference types enabled
- Modern C# patterns (pattern matching, target-typed new)
- Minimal null checks via `?` operator

### Naming Conventions
- PascalCase for public members
- camelCase for private fields
- _camelCase for private instance fields
- SCREAMING_SNAKE_CASE for constants
- Descriptive method names (verbs)
- Descriptive variable names (nouns)

### Documentation
- XML comments for public APIs
- Inline comments for complex logic
- README for user documentation
- This document for technical architecture

## Conclusion

Markdown Editor demonstrates a clean separation of concerns:
- **UI Layer**: WPF/XAML for user interaction
- **Conversion Layer**: Markdig for parsing, custom HTML generation
- **Export Layer**: OpenXML for high-fidelity Word export

The architecture is extensible for future enhancements while maintaining simplicity and performance for the core editing and export workflow.
