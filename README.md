# Browser DOCX Merger

A comprehensive TypeScript library for merging DOCX files directly in the browser with full support for styles, numbering, footnotes, endnotes, and media files.

![npm version](https://img.shields.io/npm/v/@jonathanhotono/browser-docx-merger)
![License](https://img.shields.io/badge/license-MIT-blue.svg)
![TypeScript](https://img.shields.io/badge/TypeScript-007ACC?logo=typescript&logoColor=white)

## Features

- ✅ **Browser-only** - No server required, works entirely in the browser
- ✅ **TypeScript support** - Full type definitions included
- ✅ **Multiple formats** - ESM and IIFE global builds
- ✅ **Comprehensive merging** - Styles, numbering, footnotes, themes, and more
- ✅ **Flexible insertion** - Pattern-based, start, or end insertion
- ✅ **Media handling** - Automatic copying and deduplication of images and other media
- ✅ **Relationship mapping** - Proper handling of document relationships
- ✅ **Page breaks** - Optional page break insertion between documents

## Installation

### NPM

```bash
npm install @jonathanhotono/browser-docx-merger
```

### CDN (Browser Global)

```html
<script src="https://unpkg.com/@jonathanhotono/browser-docx-merger/dist/index.global.js"></script>
```

## Usage

### ES Modules

```typescript
import { mergeDocxFromFiles, triggerDownload } from '@jonathanhotono/browser-docx-merger';

const files = Array.from(fileInput.files); // File objects from input
const options = {
  pattern: 'MERGE_HERE', // Insert at paragraphs containing this text
  mergeStyles: true,
  mergeNumbering: true,
  mergeFootnotes: true,
  pageBreaks: true,
  onLog: (message, level) => console.log(`[${level}] ${message}`)
};

try {
  const mergedBlob = await mergeDocxFromFiles(files, options);
  const url = URL.createObjectURL(mergedBlob);
  triggerDownload(url, 'merged-document.docx');
} catch (error) {
  console.error('Merge failed:', error);
}
```

### Browser Global

```html
<script src="https://unpkg.com/@jonathanhotono/browser-docx-merger/dist/index.global.js"></script>
<script>
  // Available as window.DocxMerger
  async function mergeFiles(files) {
    const merged = await DocxMerger.mergeDocxFromFiles(files, {
      insertEnd: true,
      mergeStyles: true
    });
    DocxMerger.triggerDownload(URL.createObjectURL(merged), 'merged.docx');
  }
</script>
```

### Direct Buffer Merging

```typescript
import { mergeDocx } from '@jonathanhotono/browser-docx-merger';

const buffers = [
  new Uint8Array(docx1ArrayBuffer),
  new Uint8Array(docx2ArrayBuffer)
];

const mergedBlob = await mergeDocx(buffers, {
  insertEnd: true,
  mergeStyles: true
});
```

## API Reference

### `mergeDocxFromFiles(files, options?)`

Merges multiple DOCX files from File objects.

**Parameters:**
- `files: File[]` - Array of File objects (from file input)
- `options?: MergeOptions` - Merge configuration

**Returns:** `Promise<Blob>` - The merged DOCX as a Blob

### `mergeDocx(buffers, options?)`

Merges multiple DOCX files from ArrayBuffer/Uint8Array/Blob inputs.

**Parameters:**
- `buffers: (ArrayBuffer|Uint8Array|Blob)[]` - Array of document buffers
- `options?: MergeOptions` - Merge configuration

**Returns:** `Promise<Blob>` - The merged DOCX as a Blob

### `triggerDownload(url, filename)`

Triggers a browser download of the merged document.

**Parameters:**
- `url: string` - Object URL of the Blob
- `filename: string` - Desired filename for download

## Options

```typescript
interface MergeOptions {
  // Insertion mode (exactly one must be specified)
  pattern?: string;          // Insert before paragraphs containing this text
  insertStart?: boolean;     // Insert at document start
  insertEnd?: boolean;       // Insert at document end
  
  // Content merging options
  mergeStyles?: boolean;     // Merge styles.xml, fontTable.xml, theme.xml, etc. (default: true)
  mergeNumbering?: boolean;  // Merge numbering definitions (default: true)
  mergeFootnotes?: boolean;  // Merge footnotes and endnotes (default: true)
  pageBreaks?: boolean;      // Add page breaks between documents (default: true)
  
  // Logging
  onLog?: (message: string, level: 'info'|'ok'|'warn'|'err') => void;
}
```

## Style Merging Features

The library provides comprehensive style merging including:

- **Basic Styles** (`styles.xml`) - Character, paragraph, table, and numbering styles
- **Font Tables** (`fontTable.xml`) - Font definitions and substitutions
- **Document Themes** (`theme1.xml`) - Color schemes, fonts, and effects
- **Web Settings** (`webSettings.xml`) - Web-specific styling options
- **Document Settings** (`settings.xml`) - Style-related document settings
- **Automatic Relationships** - Proper relationship mapping for all style parts

## Examples

### Pattern-Based Merging

```typescript
// Insert documents before paragraphs containing "INSERT_DOCS_HERE"
const merged = await mergeDocxFromFiles(files, {
  pattern: 'INSERT_DOCS_HERE',
  mergeStyles: true,
  pageBreaks: true
});
```

### Sequential Merging

```typescript
// Append all documents to the end with page breaks
const merged = await mergeDocxFromFiles(files, {
  insertEnd: true,
  pageBreaks: true,
  mergeStyles: true,
  mergeNumbering: true
});
```

### Custom Logging

```typescript
const merged = await mergeDocxFromFiles(files, {
  insertEnd: true,
  onLog: (message, level) => {
    const emoji = { info: 'ℹ️', ok: '✅', warn: '⚠️', err: '❌' }[level];
    console.log(`${emoji} ${message}`);
  }
});
```

## Browser Compatibility

- Chrome 61+
- Firefox 60+
- Safari 11+
- Edge 16+

Requires support for:
- ES2020 features
- Blob constructor
- DOMParser/XMLSerializer
- Promise
- Uint8Array

## Development

```bash
# Install dependencies
npm install

# Build library
npm run build

# Run tests
npm test

# Start development server with demo
npm run dev
```

## Demo

See the `demo/` directory for a complete example implementation with drag-and-drop file selection and real-time logging.

Run the demo locally:

```bash
npm run preview
```

Then open http://localhost:5173

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## Support

- 📝 [Issues](https://github.com/jonathanhotono/browser-docx-merger/issues)
- 📖 [Documentation](https://github.com/jonathanhotono/browser-docx-merger#readme)
- 💡 [Discussions](https://github.com/jonathanhotono/browser-docx-merger/discussions)

## Changelog

### v1.0.0
- Initial release
- Full DOCX merging support
- Comprehensive style handling
- TypeScript definitions
- Browser and Node.js compatibility

**Demo Features:**
- Document list showing filename and size
- Drag-and-drop reordering of documents 
- Remove individual documents
- Visual feedback during drag operations

## Basic usage

ESM:

```ts
import { mergeDocx } from 'browser-doc-merger';

const blob = await mergeDocx([file1, file2, file3], {
  insertEnd: true,
  pageBreaks: true,
});
```

Browser (IIFE bundle):

```html
<script src="./dist/index.global.js"></script>
<script>
  const blob = await window.DocxMerger.mergeDocxFromFiles(files, { insertEnd: true });
</script>
```

## Notes
- The first file acts as the base; others are inserted.
- Headers/footers from appended docs aren’t imported; section settings from base remain.
- For large files, merging is CPU/IO bound; avoid blocking the UI thread if integrating into bigger apps.
