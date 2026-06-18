# Markdown Utils - User Guide

## Table of Contents
1. [Getting Started](#getting-started)
2. [Interface Overview](#interface-overview)
3. [Editing Markdown](#editing-markdown)
4. [Formatting Toolbar](#formatting-toolbar)
5. [Keyboard Shortcuts](#keyboard-shortcuts)
6. [View Options](#view-options)
7. [Exporting to Word](#exporting-to-word)
8. [Export Options Explained](#export-options-explained)
9. [Tips & Tricks](#tips--tricks)
10. [Troubleshooting](#troubleshooting)

---

## Getting Started

### First Launch

When you first open Markdown Utils, you'll see a split-screen interface:
- **Left side**: Markdown editor with a toolbar
- **Right side**: Live formatted preview of your document

The editor starts with a sample document that demonstrates all supported Markdown features. Feel free to edit or delete this content.

### Creating Your First Document

1. Click the **"New"** button to start with a blank document
2. Type your Markdown in the left editor pane
3. Watch the formatted preview update automatically on the right
4. Click **"Save .md"** when you're ready to save your work

---

## Interface Overview

### Main Window Layout

```
???????????????????????????????????????????????????????????????
? File  View                                                  ? Menu Bar
???????????????????????????????????????????????????????????????
? [New] [Save .md] [Import .md] [Export .docx]               ?
? ? Auto-number headings  ? # as Title           Status Text  ? Toolbar
???????????????????????????????????????????????????????????????
?                          ?                                  ?
?  Raw Markdown            ?  Formatted Preview (Read Only)   ?
?  [Formatting Buttons]    ?                                  ?
?  ??????????????????????? ?  ??????????????????????????????? ?
?  ?                     ? ?  ?                             ? ?
?  ?   Editor            ? ?  ?    Preview                  ? ?
?  ?   TextBox           ? ?  ?    WebBrowser               ? ?
?  ?                     ? ?  ?                             ? ?
?  ?                     ? ?  ?                             ? ?
?  ??????????????????????? ?  ??????????????????????????????? ?
???????????????????????????????????????????????????????????????
```

### Menu Bar

**View Menu**:
- ? **Markdown Editor**: Toggle the left editor pane
- ? **Formatted Preview**: Toggle the right preview pane

Both options can be checked (split view) or only one (focused view). The app ensures at least one is always visible.

### Toolbar Buttons (Top Row)

- **New**: Clear the editor and start a new document
- **Save .md**: Save the current markdown to a file
- **Import .md**: Open an existing markdown file
- **Export .docx**: Export to Microsoft Word format

### Checkboxes (Top Row)

- **Auto-number headings**: Apply Word numbering to headings during export
- **# as Title**: Map Markdown # to Word Title style (vs. Heading 1)

### Status Bar

Displays feedback messages at the bottom right:
- "Ready" - Application idle
- "Saved: filename.md" - File saved successfully
- "Exported: filename.docx" - Word export completed
- "Applied bold formatting" - Formatting applied
- Error messages when operations fail

---

## Editing Markdown

### Basic Text

Simply type text normally. Paragraphs are separated by blank lines:

```markdown
This is the first paragraph.

This is the second paragraph.
```

### Headings

Use `#` symbols at the start of a line:

```markdown
# Heading 1 (largest)
## Heading 2
### Heading 3
#### Heading 4
##### Heading 5
###### Heading 6 (smallest)
```

**Tip**: Use the **H2** button (with dropdown) in the toolbar to quickly insert heading markers.

### Text Formatting

- **Bold**: Wrap text with `**double asterisks**`
- **Italic**: Wrap text with `*single asterisks*`
- **Bold + Italic**: Use `***triple asterisks***`
- **Strikethrough**: Wrap with `~~tildes~~`
- **Inline Code**: Wrap with `` `backticks` ``

**Example**:
```markdown
This is **bold**, this is *italic*, and this is `code`.
```

### Lists

**Bullet Lists**:
```markdown
* First item
* Second item
  * Nested item (2 spaces indent)
* Third item
```

**Numbered Lists**:
```markdown
1. First step
2. Second step
   1. Nested step
3. Third step
```

**Task Lists** (checkboxes):
```markdown
- [ ] Incomplete task
- [x] Completed task
- [ ] Another task
```

### Links

```markdown
[Link text here](https://example.com)
[Link with title](https://example.com "Hover tooltip")
```

### Images

```markdown
![Alternative text](https://example.com/image.png)
![Local image](./images/photo.jpg)
```

### Code Blocks

For multi-line code, use triple backticks:

````markdown
```csharp
using System;

Console.WriteLine("Hello World");
```
````

**Tip**: The language name after the backticks (e.g., `csharp`) enables syntax highlighting in the preview.

### Blockquotes

Use `>` at the start of lines:

```markdown
> This is a quoted paragraph.
> It can span multiple lines.
>
> You can include multiple paragraphs.
```

### Tables

Use pipes `|` and dashes `-`:

```markdown
| Column 1 | Column 2 | Column 3 |
| -------- | -------- | -------- |
| Cell A   | Cell B   | Cell C   |
| Cell D   | Cell E   | Cell F   |
```

**Alignment**:
- Left: `| --- |`
- Center: `| :-: |`
- Right: `| ---: |`

### Horizontal Rules

Use three or more dashes on their own line:

```markdown
---
```

---

## Formatting Toolbar

The toolbar provides quick access to Markdown syntax. Click a button to insert syntax at the cursor or wrap the selected text.

### Heading Button (H2 ?)

- **Main button (H2)**: Inserts the currently selected heading level
- **Dropdown arrow (?)**: Opens a menu to choose H1-H6
- After choosing, the button updates to show the new level (e.g., "H3")

**Usage**:
1. Select text or place cursor
2. Click **H2** (or your chosen level)
3. Heading markers are added: `## Selected text`

### Text Formatting Buttons

| Button | Action | Syntax Added |
| --- | --- | --- |
| **Bold** | Make text bold | `**text**` |
| **Italic** | Make text italic | `*text*` |
| **Strike** | Add strikethrough | `~~text~~` |
| **Inline Code** | Format as code | `` `text` `` |

**How they work**:
- If text is selected: Wraps selection with syntax
- If nothing selected: Inserts placeholder like `**bold text**`

### Block Formatting Buttons

| Button | Action | Syntax Added |
| --- | --- | --- |
| **Code Block** | Insert fenced code block | ` ```\ntext\n``` ` |
| **Rule** | Insert horizontal line | `---` |
| **Quote** | Make text a blockquote | `> text` |

### List Buttons

| Button | Action | Syntax Added |
| --- | --- | --- |
| **Bullet List** | Create bulleted list | `* text` |
| **Numbered List** | Create numbered list | `1. text` |

**Multi-line behavior**:
- If multiple lines are selected, each line becomes a list item
- Existing list markers are not duplicated

### Link and Image Buttons

| Button | Action | Syntax Added |
| --- | --- | --- |
| **Link** | Insert hyperlink | `[text](url)` |
| **Image** | Insert image reference | `![alt](url)` |

**After inserting**:
- Replace `text`/`alt` with description
- Replace `url` with actual web address or file path

---

## Keyboard Shortcuts

Speed up your editing with these shortcuts:

| Shortcut | Action |
| --- | --- |
| **Ctrl+B** | Bold (wrap with `**`) |
| **Ctrl+I** | Italic (wrap with `*`) |
| **Ctrl+K** | Insert link |
| **Ctrl+2** | Insert H2 heading |
| **Ctrl+3** | Insert H3 heading |

**Tip**: You can select text first, then use a shortcut to wrap it with formatting.

---

## View Options

Use the **View** menu to customize your workspace:

### Split View (Default)
- Both checkboxes enabled
- Editor on left, preview on right
- Drag the center splitter to resize

**Best for**: Writing while seeing live preview

### Markdown Only
- Only "Markdown Editor" checked
- Editor takes full window width
- Preview hidden

**Best for**: Focused writing, working with large documents, or when preview isn't needed

### Formatted Preview Only
- Only "Formatted Preview" checked
- Preview takes full window width
- Editor hidden

**Best for**: Reading final document, presenting to others, or reviewing formatting

### Switching Views

**Method 1**: Menu
- Click **View** menu
- Click **Markdown Editor** or **Formatted Preview** to toggle

**Method 2**: Keyboard
- Press **Alt+V** to open View menu
- Use arrow keys to navigate
- Press Enter to toggle an option

**Smart Toggle**: If you try to disable both views, the app automatically enables the other view to ensure something is always visible.

---

## Exporting to Word

### Basic Export

1. Click the **"Export .docx"** button
2. Choose a location and filename in the save dialog
   - The filename is automatically suggested from your first heading
   - You can change it if desired
3. Click **Save**
4. Wait for "Exported: filename.docx" status message
5. Open the file in Microsoft Word

### What Gets Exported

? **Fully Supported**:
- All heading levels (H1-H6)
- Bold, italic, strikethrough, inline code
- Bullet and numbered lists (with proper Word formatting)
- Nested lists (up to 9 levels)
- Tables with borders and formatting
- Code blocks with monospace font
- Blockquotes with left border
- Horizontal rules as actual lines
- Task lists (rendered as ? and ? checkboxes)
- Links (clickable in Word)

?? **Partial Support**:
- Images: Image references are exported but may not display if paths are invalid

### Export Quality

The Word export aims for **preview parity** - what you see in the formatted preview should match what you get in Word:

- **Headings**: Styled with Word's built-in heading styles (Heading 1, Heading 2, etc.)
- **Lists**: Native Word bullets and numbering (not just text characters)
- **Tables**: Properly formatted with borders and bold headers
- **Code**: Monospace Consolas font with gray background
- **Colors**: Headings use blue color scheme for professional appearance

---

## Export Options Explained

Two checkboxes control export behavior:

### Auto-number headings

**When CHECKED (default)**:
- The exporter looks for manual numbers in headings
- Example: `## 1.2.3 System Overview` ? "1.2.3" is detected
- These numbers are removed from the text
- Word's automatic numbering is applied (hierarchical: 1, 1.1, 1.1.1, etc.)
- All headings get consistent numbering

**When UNCHECKED**:
- Headings are exported exactly as written
- Manual numbers (if any) are kept in the text
- No automatic numbering is added
- You have full control over heading text

**When to use**:
- ? Check if: Your document uses numbered sections and you want consistent Word numbering
- ? Uncheck if: Your headings don't need numbers, or you want to keep manual numbers as-is

**Example**:

Markdown input:
```markdown
## 1.2 Overview
### 1.2.1 Hardware
### 1.2.2 Software
## 1.3 Configuration
```

With checkbox **checked**: 
- Word shows: "1. Overview", "1.1. Hardware", "1.2. Software", "2. Configuration"

With checkbox **unchecked**:
- Word shows: "1.2 Overview", "1.2.1 Hardware", "1.2.2 Software", "1.3 Configuration" (no auto-numbering)

### # as Title

**When CHECKED (default)**:
- Markdown `#` maps to Word's **Title** style (special large format)
- Markdown `##` maps to **Heading 1**
- Markdown `###` maps to **Heading 2**
- And so on...

**When UNCHECKED**:
- Markdown `#` maps to Word's **Heading 1**
- Markdown `##` maps to **Heading 2**
- Markdown `###` maps to **Heading 3**
- Standard markdown-to-heading mapping

**When to use**:
- ? Check if: Your document has a main title as `#` and section headings below
- ? Uncheck if: You want traditional heading hierarchy starting at `#` = Heading 1

**Example**:

Markdown input:
```markdown
# Project Documentation
## Overview
### Details
```

With checkbox **checked**:
- "Project Documentation" = Title style (26pt, very large)
- "Overview" = Heading 1 (16pt)
- "Details" = Heading 2 (13pt)

With checkbox **unchecked**:
- "Project Documentation" = Heading 1 (16pt)
- "Overview" = Heading 2 (13pt)
- "Details" = Heading 3 (12pt)

### Combining Both Options

The two options work independently. Common combinations:

| Auto-number | # as Title | Result |
| --- | --- | --- |
| ? | ? | Document with title + numbered headings (most common) |
| ? | ? | Document with title, no numbering |
| ? | ? | Traditional headings with numbering |
| ? | ? | Plain export, no modifications |

---

## Tips & Tricks

### Quickly Change Heading Level

1. Place cursor on a heading line
2. Click the **H2 ?** dropdown
3. Select a new level (e.g., H4)
4. Click the **H4** button
5. Heading is updated: `##` ? `####`

### Convert Plain Lines to List

1. Select multiple lines of text
2. Click **Bullet List** or **Numbered List**
3. Each line becomes a list item

### Preview Not Updating?

The preview updates automatically on every keystroke. If it seems stuck:
- Try typing a space and deleting it
- Toggle the view menu options
- Restart the application

### Suggested Filename

When exporting, the filename is automatically suggested from your document's first heading:
- `# Project Overview` ? suggests "Project Overview.docx"
- Special characters are replaced with underscores
- Long names are truncated to 200 characters

### Working with Long Documents

For documents > 1000 lines:
- Use **Markdown Only** view to speed up editing
- Toggle preview only when you need to check formatting
- The preview render is fast, but disabling it removes the overhead entirely

### Keyboard Navigation in Editor

Standard TextBox shortcuts work:
- **Ctrl+A**: Select all
- **Ctrl+C/X/V**: Copy, cut, paste
- **Ctrl+Z/Y**: Undo, redo
- **Home/End**: Start/end of line
- **Ctrl+Home/End**: Start/end of document

---

## Troubleshooting

### "Export failed" Error

**Possible causes**:
1. Destination path is not writable
2. File is open in Word (locked)
3. Disk is full
4. Invalid characters in filename

**Solutions**:
- Close any open Word documents with the same name
- Choose a different destination folder
- Check disk space
- Verify you have write permissions to the folder

### Preview Shows HTML/Markdown Instead of Formatting

**Cause**: WebBrowser control not initializing properly

**Solution**:
- Restart the application
- Check that .NET 9 Desktop Runtime is installed
- Verify Windows is up to date

### Bullets Show as Numbers in Word

**Cause**: This was a bug fixed in commit 7384322

**Solution**:
- Make sure you have the latest version
- Each list now gets a unique numbering instance
- Restart the app after updating

### Formatting Lost in Lists or Tables

**Cause**: Older versions used plain text extraction

**Solution**:
- Update to latest version
- Inline formatting (bold, code, links) is now preserved everywhere
- Verify in preview first - preview should match export

### Can't Uncheck Both View Options

**Behavior**: This is intentional, not a bug

**Reason**: At least one view must be visible for the app to be useful

**How it works**: When you uncheck the last checked option, the other automatically checks itself

### Word File Won't Open

**Possible causes**:
1. Incomplete export (application crashed mid-export)
2. Corrupted file
3. Word version incompatibility

**Solutions**:
- Try exporting again
- Check that export completed (look for success message)
- Open in newer version of Word (2013+)
- Try Word's repair feature: Open ? Browse ? select file ? Open ? ? Open and Repair

### Preview Looks Different than Word Export

**Minor differences are expected**:
- Preview uses HTML/CSS rendering (WebBrowser)
- Word uses OpenXML rendering
- Colors, spacing, and fonts should be very close
- If major differences, please report as a bug

**Check these**:
- Is syntax correct? (test in preview first)
- Are export options set as expected?
- Try latest version of the app

---

## Getting Help

If you encounter issues not covered here:

1. Check the README.md for general information
2. Check ARCHITECTURE.md for technical details
3. Verify you're using the latest version
4. Review the sample document for syntax examples
5. Try exporting the sample document to verify export works

For bugs or feature requests, please open an issue in the repository with:
- Description of the problem
- Steps to reproduce
- Expected vs. actual behavior
- Sample markdown that demonstrates the issue

---

## Quick Reference Card

### Markdown Syntax

| Element | Syntax | Example |
| --- | --- | --- |
| Heading 1 | `#` | `# Title` |
| Heading 2 | `##` | `## Section` |
| Bold | `**text**` | `**bold**` |
| Italic | `*text*` | `*italic*` |
| Code | `` `text` `` | `` `code` `` |
| Link | `[text](url)` | `[Click](url)` |
| Image | `![alt](url)` | `![Photo](pic.jpg)` |
| Bullet List | `* item` | `* First` |
| Numbered List | `1. item` | `1. First` |
| Task | `- [ ] task` | `- [ ] Todo` |
| Code Block | ` ``` ` | ` ```code``` ` |
| Quote | `> text` | `> Quoted` |
| Rule | `---` | `---` |
| Table | `\| A \| B \|` | See above |

### Keyboard Shortcuts

| Key | Action |
| --- | --- |
| Ctrl+B | Bold |
| Ctrl+I | Italic |
| Ctrl+K | Link |
| Ctrl+2 | H2 |
| Ctrl+3 | H3 |

### Export Options

| Option | Effect |
| --- | --- |
| Auto-number headings | Strips manual numbers, applies Word numbering |
| # as Title | Maps # to Title (vs. Heading 1) |

---

**Happy Markdown editing!** ??
