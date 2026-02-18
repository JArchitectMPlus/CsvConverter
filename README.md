# SharePoint Client-Side App: CSV to Markdown Converter

A browser-based HTML/JavaScript application hosted on SharePoint for converting CSV data to Markdown format without size limitations.

## Quick Start

1. **Push this repo to GitHub** → Enable GitHub Pages on `/dist` folder
2. **Get your URL**: `https://yourusername.github.io/repo-name/app.html`
3. **Embed in SharePoint** using iframe (see [Deployment](#deployment) for details)

## Overview

This client-side application converts CSV survey data to well-formatted Markdown documents directly in the browser. Unlike Office Scripts (limited to 32,767 characters per cell), this solution has no output size limits and processes all data locally without server transmission.

**Key Benefits:**
- No output size limits
- All processing in browser—data never leaves the client
- Accessible to anyone with SharePoint site access
- No backend infrastructure required
- Single HTML file deployment
- Free hosting via GitHub Pages
- Easy to update and maintain

**Why GitHub Pages + iframe?**
SharePoint's Embed web part blocks inline `<script>` tags for security. By hosting on GitHub Pages and embedding via iframe, we maintain the self-contained single-file architecture while meeting SharePoint's security requirements.

## Features

- **CSV Input Options**: Paste CSV text or upload CSV files
- **Smart Parsing**: Automatically detects survey questions, headers, and response data
- **Markdown Output**: Generates clean Markdown tables with proper formatting
- **Export Options**: Copy to clipboard or download as .md file
- **Client-Side Processing**: All data processing happens locally in your browser
- **Special Character Handling**: Escapes pipe characters and handles edge cases

## Deployment

### Prerequisites

**Important**: SharePoint's Embed web part does not support inline script tags. The app must be hosted externally and embedded via iframe.

### Step 1: Host on GitHub Pages

1. **Push this repository to GitHub**
   ```bash
   git init
   git add .
   git commit -m "Initial commit: SharePoint CSV to Markdown converter"
   git remote add origin https://github.com/yourusername/your-repo-name.git
   git push -u origin main
   ```

2. **Enable GitHub Pages**
   - Go to repository Settings → Pages
   - Under "Source", select: `main` branch
   - Under "Folder", select: `/dist`
   - Click "Save"
   - GitHub will provide your URL: `https://yourusername.github.io/your-repo-name/app.html`

3. **Wait for deployment** (usually 1-2 minutes)
   - Visit your GitHub Pages URL to verify the app loads

### Step 2: Embed in SharePoint

1. Navigate to your SharePoint page (or create a new one)
2. Click "Edit" to enter edit mode
3. Click the "+" icon to add a new web part
4. Search for and select "Embed" web part
5. In the "Website address or embed code" field, paste:
   ```html
   <iframe src="https://yourusername.github.io/your-repo-name/app.html"
           width="100%"
           height="800"
           style="border: none;">
   </iframe>
   ```
6. Click outside the web part to preview
7. Click "Save" or "Publish" to make the page live

**Note**: Replace `yourusername` and `your-repo-name` with your actual GitHub username and repository name.

### Step 3: Updating the App

When you make changes to `dist/app.html`:

1. **Commit and push changes**
   ```bash
   git add dist/app.html
   git commit -m "Update app"
   git push
   ```

2. **Wait for GitHub Pages to rebuild** (1-2 minutes)

3. **Refresh SharePoint page** (or clear cache with Ctrl+Shift+R)
   - SharePoint will automatically show the updated version
   - No need to re-embed the iframe

### Alternative: Local Testing

Simply open `dist/app.html` directly in any modern web browser. No server required for local development and testing.

## Usage

1. Navigate to the SharePoint page hosting the app
2. Either:
   - Paste CSV content into the input area, OR
   - Click "Upload CSV" and select a file
3. Click "Convert"
4. Copy the Markdown output or download as a `.md` file

## Technical Details

**Architecture**: Single HTML file with embedded CSS/JavaScript
</br>
**Processing**: Client-side only—no data transmitted to servers
</br>
**Compatibility**: Chrome, Firefox, Safari, Edge (latest versions)
</br>
**Detection**: Intelligently identifies questions, headers, and data rows using pattern matching

## Configuration

Configurable patterns in JavaScript:
- **Questions**: S1:, [Age], CTP:, ending with ?, hSample patterns
- **Responses**: Total, Mean, Median, Yes/No, demographic categories

## Output Format

Generated Markdown includes:

```markdown
# Survey Data

**Total Records:** 500
**Total Columns:** 15
**Generated:** February 18, 2026

## Survey Questions and Responses

### Question 1

**S1: What is your primary use case?**

#### Response Data

| Response | Total (A) | Pro-to-Pro Switchers (B) | Software-to-Pro Switchers (C) |
|---|---|---|---|
| N=400 | N=100 | N=150 | N=150 |
| Personal use | 45% | 30% | 50% |
| Business use | 55% | 70% | 50% |

### Question 2
...
```

## Repository Structure

```
typescript/officeScript/
├── dist/
│   └── app.html              # Production app (served by GitHub Pages)
├── README.md                 # This file
├── package.json              # TypeScript development dependencies
└── tsconfig.json             # TypeScript configuration
```

**GitHub Pages Configuration:**
- Serves only the `/dist` folder
- All other files remain in the repo but aren't publicly served
- This keeps the deployment clean while maintaining full project context

## Development

### Prerequisites

- Node.js and npm (for TypeScript development/type checking)
- TypeScript 5.x

### Local Development

```bash
# Install dependencies
npm install

# Type check
npx tsc --noEmit
```

### Testing

Test with `sample_survey.csv` (simple) and `TurboTax_sample.csv` (complex). Verify edge cases: special characters, empty cells, large files.

## Customization

Edit `dist/app.html` to customize:
1. Question detection patterns (regex)
2. Header label expectations
3. Response category recognition

## Troubleshooting

- **No questions found**: Verify CSV structure matches expected patterns, check browser console
- **Headers not detected**: Ensure rows contain "Total" or "N=" values
- **SharePoint "Script tags not supported"**: Use iframe embed code (see deployment)
- **Updates not showing**: Wait 1-2 minutes for GitHub Pages rebuild, hard refresh (Ctrl+Shift+R)
- **App errors**: Check browser console, verify modern browser with JavaScript enabled

## References

- [SharePoint Embed Web Part](https://support.microsoft.com/en-us/office/add-content-to-your-page-using-the-embed-web-part-721f3b2f-437f-45ef-ac4e-df29dba74de8)
- Original Python reference implementation (converted to JavaScript)

## Version History

- **v2.0** (February 2026) - SharePoint client-side application
  - Migrated from Office Scripts to HTML/JavaScript
  - Removed 32K character output limit
  - Added file upload capability
  - Added download as .md functionality
  - Single-file deployment model

- **v1.0** (February 2026) - Office Scripts implementation (deprecated)

## License

This application is provided as-is for converting CSV survey data to Markdown format.
