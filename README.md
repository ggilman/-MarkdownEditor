# Markdown Utils

A desktop application for editing, previewing, and exporting Markdown documents to Microsoft Word format with advanced formatting options.

## Features

### Markdown Editing
- Full-featured Markdown editor with syntax support
- Live formatted preview pane
- Keyboard shortcuts for common formatting operations
- Toolbar buttons for quick access to Markdown syntax

### Export to Word (.docx)
- Export Markdown documents to Microsoft Word format
- Preserves formatting including:
  - Headings with configurable styles
  - Bold, italic, strikethrough, and inline code
  - Bulleted and numbered lists with proper Word native formatting
  - Tables with borders and formatting
  - Code blocks with syntax highlighting
  - Blockquotes
  - Horizontal rules
  - Links and images
  - Task lists (checkboxes)

### Configurable Export Options
- **Auto-number headings**: Automatically apply Word's multilevel numbering to headings
- **# as Title**: Choose whether Markdown `#` maps to Word's Title style or Heading 1

### View Options
- **Split View**: Edit and preview side-by-side (default)
- **Markdown Only**: Focus on editing without the preview
- **Formatted Preview Only**: View the rendered document without the editor

## Requirements

- **Platform**: Windows
- **.NET**: .NET 9 or later
- **Dependencies**:
  - DocumentFormat.OpenXml (for Word export)
  - Markdig (for Markdown parsing)

## Installation

1. Clone or download this repository
2. Open the solution in Visual Studio 2022 or later
3. Restore NuGet packages
4. Build and run the application

## Usage

### Basic Editing

1. **New Document**: Click "New" or press Ctrl+N to start a new document
2. **Open Markdown**: Click "Import .md" to open an existing Markdown file
3. **Save Markdown**: Click "Save .md" to save your work as a `.md` file

### Formatting Shortcuts

| Shortcut | Action |
| --- | --- |
| Ctrl+B | Bold |
| Ctrl+I | Italic |
| Ctrl+K | Insert link |
| Ctrl+2 | Insert H2 heading |
| Ctrl+3 | Insert H3 heading |

### Toolbar Buttons

- **H2 (with dropdown)**: Insert headings (H1-H6)
- **Bold**: Wrap selection in `**bold**`
- **Italic**: Wrap selection in `*italic*`
- **Strike**: Wrap selection in `~~strikethrough~~`
- **Inline Code**: Wrap selection in `` `code` ``
- **Code Block**: Create fenced code block with ` ``` `
- **Rule**: Insert horizontal rule `---`
- **Bullet List**: Create bulleted list with `*`
- **Numbered List**: Create numbered list with `1.`
- **Quote**: Create blockquote with `>`
- **Link**: Insert Markdown link `[text](url)`
- **Image**: Insert image reference `![alt](url)`

### Export to Word

1. Click "Export .docx" button
2. Choose export options:
   - **Auto-number headings**: Strips manual numbers from headings and applies Word's multilevel numbering
   - **# as Title**: Maps Markdown `#` to Word Title style (checked) or Heading 1 (unchecked)
3. Select destination and filename (automatically suggests filename from first heading)
4. The document will be created with full formatting

### View Menu

Use the **View** menu to control which panels are visible:

- **Markdown Editor**: Show/hide the Markdown editor pane
- **Formatted Preview**: Show/hide the live preview pane

Both options can be checked for split view, or only one for focused editing/preview. At least one must always be visible.

## Export Options Explained

### Auto-number headings

When **checked** (default):
- The exporter scans headings for leading numbers (e.g., "## 1.2.3 Overview")
- These numbers are stripped from the heading text
- Word's native multilevel numbering is applied automatically
- Numbering is hierarchical (1, 1.1, 1.1.1, etc.)

When **unchecked**:
- Headings are exported as-is without stripping numbers
- No automatic numbering is applied
- Manual numbers in your Markdown are preserved

### # as Title

When **checked** (default):
- Markdown `#` (H1) ? Word **Title** style
- Markdown `##` (H2) ? Word **Heading 1** style
- Markdown `###` (H3) ? Word **Heading 2** style
- And so on...

When **unchecked**:
- Markdown `#` (H1) ? Word **Heading 1** style
- Markdown `##` (H2) ? Word **Heading 2** style
- Markdown `###` (H3) ? Word **Heading 3** style
- Standard markdown-to-heading mapping

## Supported Markdown Syntax

### Headings
```markdown
# Heading 1
## Heading 2
### Heading 3
#### Heading 4
##### Heading 5
###### Heading 6
```

### Text Formatting
```markdown
**bold text**
*italic text*
~~strikethrough~~
`inline code`
```

### Lists
```markdown
* Bullet item
* Another item
  * Nested bullet

1. Numbered item
2. Another item
   1. Nested numbered item

- [ ] Task item (unchecked)
- [x] Completed task
```

### Links and Images
```markdown
[Link text](https://example.com)
![Alt text](image-url.png)
```

### Code Blocks
````markdown
```csharp
using System;
Console.WriteLine("Hello World");
```
````

### Blockquotes
```markdown
> This is a quoted paragraph.
> It can span multiple lines.
```

### Tables
```markdown
| Header 1 | Header 2 | Header 3 |
| --- | --- | --- |
| Cell 1 | Cell 2 | Cell 3 |
| Row 2 | Data | More data |
```

### Horizontal Rules
```markdown
---
```

## Word Export Details

### Formatting Preservation

The Word export maintains high fidelity to the preview:

- **Headings**: Styled with Word's built-in heading styles including color and sizing
- **Lists**: Native Word bullet and numbered lists (not plain text)
- **Tables**: Bordered tables with bold headers and centered header text
- **Code Blocks**: Monospace font (Consolas) with gray background shading
- **Inline Code**: Monospace font with light gray background
- **Blockquotes**: Left border with gray shading
- **Horizontal Rules**: Actual horizontal line separators (not dashes)

### List Numbering

Each list in the document receives its own unique numbering instance. This means:
- Multiple bulleted lists won't continue numbering from each other
- Numbered lists restart at 1 for each new list
- Nested lists are properly indented with correct numbering hierarchy

### Inline Formatting in Lists and Tables

Inline formatting (bold, italic, code, links) is fully preserved within:
- List items
- Table cells
- Blockquotes

## Architecture

### Technologies
- **Framework**: WPF (.NET 9)
- **Markdown Parser**: Markdig with advanced extensions
- **Word Export**: DocumentFormat.OpenXml
- **Preview Rendering**: WebBrowser control with HTML conversion

### Project Structure
```
MarkdownUtilsApp/
??? MainWindow.xaml          # UI layout
??? MainWindow.xaml.cs       # UI event handlers and editing logic
??? DocxExporter.cs          # Word document generation
??? MarkdownConverter.cs     # HTML preview rendering
??? MarkdownUtilsApp.csproj  # Project file
```

### Key Components

**MainWindow**: 
- Markdown editor with toolbar
- Live preview pane
- File operations (new, save, import, export)
- View menu for panel visibility
- Keyboard shortcuts

**DocxExporter**:
- Markdown to OpenXML conversion
- Style definitions (Title, Headings, Code Block)
- Numbering definitions (bullets, numbers, heading numbering)
- Block and inline element processing
- Unique list instance creation

**MarkdownConverter**:
- Markdown to HTML conversion for preview
- Uses Markdig pipeline with advanced extensions

## Development

### Building
```bash
dotnet build
```

### Running
```bash
dotnet run
```

### Adding New Markdown Features

1. **Parser Support**: Ensure Markdig supports the feature (enable extensions if needed)
2. **Preview Rendering**: Update `MarkdownConverter.cs` if custom HTML rendering is required
3. **Word Export**: Add handling in `DocxExporter.cs`:
   - Add new block type to `AppendBlock` switch statement
   - Create helper methods for OpenXML generation
   - Define any required styles in `EnsureStylesPart`

## Troubleshooting

### Preview not updating
- The preview updates automatically on text changes
- If stuck, try toggling the view menu options or restarting the app

### Export errors
- Ensure the destination path is writable
- Close any open Word documents with the same filename
- Check that required NuGet packages are installed

### Lists showing numbers instead of bullets
- This was fixed in commit 7384322
- Ensure you have the latest version
- Each list now gets a unique NumberingInstance

### Formatting lost in export
- Verify the markdown syntax is correct in the editor
- Check the preview - what you see should match the Word export
- Some advanced HTML/CSS features may not translate to Word

## Contributing

Feel free to submit issues or pull requests for:
- Bug fixes
- New Markdown features
- Export improvements
- UI enhancements
- Documentation updates

## License

This project is provided as-is for personal and educational use.

## Version History

### Latest (af72f26)
- Added View menu with Markdown Editor and Formatted Preview toggles
- Smart toggle ensures at least one view is always visible
- Dynamic grid layout based on selected view mode

### Previous (7384322)
- Added "# as Title" checkbox for configurable heading mapping
- Fixed list numbering issue where separate lists continued numbering
- Each list now gets unique NumberingInstance
- Removed StartNumberingValue from bullet list definitions

### Earlier (76d92fb)
- Initial feature set
- Word export with formatting preservation
- Auto-number headings option
- Native Word lists and tables
- Inline formatting preservation

## Contact

For questions or support, please open an issue in the repository.
