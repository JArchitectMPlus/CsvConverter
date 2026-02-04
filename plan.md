# Office Scripts: CSV to Markdown Converter - Implementation Plan

## Status: READY FOR IMPLEMENTATION

---

## Source File

**Python Reference:** `survey_csv_to_markdown.py` (445 lines)

This script will be converted to Office Scripts TypeScript format.

---

## Decisions (Confirmed)

| Decision | Choice |
|----------|--------|
| **Scope** | Survey-specific (TurboTax format) |
| **Output** | New worksheet "Markdown_Output" |
| **Test Data** | Actual TurboTax CSV available |
| **Approach** | Reimplementation using Python as reference |

---

## Technical Architecture

### Office Scripts Constraints
- No file system access (no `Path`, no file I/O)
- No external libraries (no pandas)
- Data accessed via `ExcelScript.Workbook` API
- Input: 2D array from `Range.getValues()`
- Output: Write to cells/range

### Data Structure Design

```typescript
// Core interfaces
interface QuestionSection {
  questionRowIndex: number;
  questionText: string;
  headerRowIndices: number[];
  dataRowIndices: number[];
  endRowIndex: number;
}

interface MarkdownOutput {
  header: string;
  sections: string[];
  summary: string;
}
```

### Python → TypeScript Translation Map

| Python Construct | TypeScript Equivalent |
|-----------------|----------------------|
| `pandas.read_csv()` | `worksheet.getUsedRange().getValues()` |
| `df.iloc[row]` | `data[row]` |
| `str(cell).strip()` | `String(cell).trim()` |
| `cell_str.lower()` | `cellStr.toLowerCase()` |
| `re.search(pattern, string)` | `/pattern/.test(string)` |
| `len(string)` | `string.length` |

---

## Implementation Phases

### Phase 1: Setup
- [ ] Create test workbook in Excel Online
- [ ] Import sample CSV data
- [ ] Create new Office Script via Automate tab
- [ ] Set up "Markdown_Output" sheet

### Phase 2: Core Functions
- [ ] `firstNonEmptyCell(row: any[]): string`
- [ ] `isQuestionRow(row: any[]): boolean`
- [ ] `isHeaderRow(row: any[]): boolean`
- [ ] `isDataRow(row: any[]): boolean`
- [ ] `findQuestionBoundaries(data: any[][]): number[]`

### Phase 3: Markdown Generation
- [ ] `formatMarkdownTable(headers: string[][], data: string[][]): string`
- [ ] `formatQuestionSection(questionNum: number, text: string, headers: string[][], data: string[][]): string`
- [ ] `escapeMarkdownChars(text: string): string`

### Phase 4: Main Script
- [ ] Read data from active worksheet
- [ ] Process all question sections
- [ ] Generate complete markdown document
- [ ] Write to output sheet
- [ ] Add error handling

### Phase 5: Testing
- [ ] Test with minimal dataset (3 questions)
- [ ] Test edge cases (empty cells, pipe characters, special chars)
- [ ] Test with actual survey data
- [ ] Compare output to Python script results

### Phase 6: Documentation
- [ ] Add inline comments
- [ ] Create Instructions sheet
- [ ] Share via Automate tab

---

## Key Functions to Port from `survey_csv_to_markdown.py`

### 1. `first_non_empty_cell(row)` — Lines 29-38
Returns first non-empty, non-'nan' cell from a row as trimmed string.

### 2. `is_question_row(row)` — Lines 41-68
Regex patterns to identify question rows:
- `^[A-Z]\d+:` (e.g., S1:, A1:)
- `^\[.*\]` (e.g., [Age])
- `\?$` (ends with ?)
- `^CTP:`
- `^h[A-Z]`
- Long text (>50 chars) with question keywords but not response indicators

### 3. `is_header_row(row)` — Lines 71-106
Identifies column headers by:
- "Total" with N= values
- "Total (A)" pattern
- Specific profile labels: `Pro-to-Pro Switchers (B)`, `Software-to-Pro Switchers (C)`, `Pro-to-Software Switchers (D)`

### 4. `is_data_row(row)` — Lines 109-154
Identifies response data by:
- Numeric patterns (digits, percentages, decimals)
- Known response categories (Total, Mean, Yes/No, agree/disagree, etc.)
- Reasonable length (1-120 chars)

### 5. `format_survey_csv_to_markdown(df, filename)` — Lines 157-303
Main conversion logic:
- Find question boundaries
- Process each question section (headers + data rows)
- Combine multiple header rows (descriptive + N= sample sizes)
- Search for N= rows within section or above question start

### 6. `format_question_section(question_num, question, header_rows, data_rows)` — Lines 306-360
Markdown table generation:
- Table headers with `|` separators
- Separator row with `---|`
- Secondary header row (N= values) as first data row
- Data rows with escaped pipe characters

---

## Configuration Constants (from Python)

These will be defined at the top of the TypeScript file for easy customization:

```typescript
// Question detection patterns
const QUESTION_PATTERNS = [
  /^[A-Z]\d+:/,   // S1:, A1:, B1:
  /^\[.*\]/,      // [Age], [Gender]
  /\?$/,          // Ends with ?
  /^CTP:/,        // CTP questions
  /^h[A-Z]/       // hSample, hB1_Flag
];

// Question keywords for long text detection
const QUESTION_KEYWORDS = ['what', 'how', 'which', 'please', 'following'];

// Response indicators (exclude from question detection)
const RESPONSE_INDICATORS = [
  'marketing or advertising agency', 'business entity',
  'educational institution', 'organization', 'government agency'
];

// Known response categories
const RESPONSE_CATEGORIES = [
  'Total', 'Mean', 'Median', 'Standard Deviation',
  'Yes', 'No', 'Male', 'Female', 'Other',
  'Very important', 'Somewhat important', 'Not important',
  'Strongly agree', 'Agree', 'Disagree',
  'Currently use', 'Used', 'Never used',
  'A marketing or advertising agency', 'A business entity',
  'An educational institution', 'An organization or non-profit',
  'A government agency'
];

// Survey-specific header labels (TurboTax format)
const SWITCHER_LABELS = [
  "Pro-to-Pro Switchers (B)",
  "Software-to-Pro Switchers (C)",
  "Pro-to-Software Switchers (D)"
];
```

---

## Verification Plan

1. **Unit verification**: Test each detection function with known inputs
2. **Integration verification**: Run full script on test workbook
3. **Comparison verification**: Compare markdown output to Python script output on same data
4. **Edge case verification**: Empty sheets, single rows, special characters

---

## Files to Create

1. `src/csvToMarkdown.ts` - Main Office Script (will be copied to Excel Online)
2. Test workbook with sample data

---

## Risks

| Risk | Mitigation |
|------|------------|
| Complex header-combining logic (lines 218-287) | Simplify for V1, iterate later |
| Dataset-specific patterns | Document as configuration section |
| Large dataset performance | Batch processing, progress indicator |
| Regex differences Python vs JS | Test all patterns explicitly |