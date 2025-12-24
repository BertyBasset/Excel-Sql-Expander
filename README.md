# SQL Template Expander - Usage Guide

## Overview

The Enhanced SQL Template Expander is a VBA utility for Excel that generates SQL INSERT statements by replacing column references in template strings with actual cell values from your worksheet.

## Quick Start

### Basic Usage

In any cell, use the formula:
```excel
=ExpandTemplate("$A, B, C")
```

If row contains: `A=John`, `B=25`, `C=5.5`  
Output: `'John', 25, 5.5`

### Generate Complete INSERT Statement

```excel
=ExpandTemplate("INSERT INTO users (name, age, score) VALUES ($A, B, C);")
```

Output: `INSERT INTO users (name, age, score) VALUES ('John', 25, 5.5);`

## Type Prefixes

Control how values are formatted using prefix characters before column letters:

| Prefix | Type | Example | Input: `A=123` | Output |
|--------|------|---------|----------------|--------|
| (none) | Auto-detect | `A` | `123` | `123` |
| `$` | String (force quotes) | `$A` | `123` | `'123'` |
| `#` | Numeric (force) | `#A` | `'123'` | `123` |
| `@` | Date/DateTime | `@A` | `2024-01-15` | `'2024-01-15 00:00:00'` |
| `!` | Raw/Literal | `!A` | `NOW()` | `NOW()` |
| `?` | Boolean | `?A` | `TRUE` | `1` |
| `~` | Empty string | `~A` | (empty) | `''` |

### Auto-Detection (No Prefix)

When no prefix is used, the function respects Excel's cell formatting first:
- **Text-formatted cells** (format `@`): Always quoted → `'123'`
- **Date-formatted cells**: Formatted as SQL datetime → `'2024-01-15 00:00:00'`
- **Number-formatted cells**: Unquoted → `42`, `3.14`, `1000.50`
- **Currency/Percentage**: Treated as numbers → `1234.56`
- **General format**: Numeric values unquoted, text values quoted
- **Empty cells**: Become `NULL` (or `''` if `nullForEmpty=FALSE`)

The prefix **always overrides** Excel's cell type.

### Examples by Prefix

#### `$` - Force String
```excel
=ExpandTemplate("$A, $B, $C")
```
Input: `A=123`, `B=Hello`, `C=45.5`  
Output: `'123', 'Hello', '45.5'`

#### `#` - Force Numeric
```excel
=ExpandTemplate("#A, #B")
```
Input: `A=123`, `B=456`  
Output: `123, 456`

#### `@` - Date/DateTime
```excel
=ExpandTemplate("@A, @B")
```
Input: `A=2024-01-15`, `B=2024-12-25 14:30:00`  
Output: `'2024-01-15 00:00:00', '2024-12-25 14:30:00'`

#### `!` - Raw/Literal (SQL Functions)
```excel
=ExpandTemplate("$A, !B, !C")
```
Input: `A=John`, `B=NOW()`, `C=DEFAULT`  
Output: `'John', NOW(), DEFAULT`

Use for SQL functions, keywords, or any value that shouldn't be escaped or quoted.

#### `?` - Boolean
```excel
=ExpandTemplate("$A, ?B, ?C")
```
Input: `A=John`, `B=TRUE`, `C=0`  
Output: `'John', 1, 0`

Converts: `TRUE`→`1`, `FALSE`→`0`, non-zero numbers→`1`, zero→`0`

#### `~` - Empty String Instead of NULL
```excel
=ExpandTemplate("$A, ~B")
```
Input: `A=John`, `B=(empty)`  
Output: `'John', ''`

By default, empty cells produce `NULL`. Use `~` to force empty string `''`.

## Multi-Column Support

Supports column references from **A to ZZZ** (18,278 columns):

```excel
=ExpandTemplate("$A, $Z, $AA, $AB, $ZZ, $AAA")
```

## Named Ranges

Reference named ranges using `{RangeName}` syntax:

```excel
=ExpandTemplate("{CustomerName}, {CustomerAge}, {Score}")
```

If you've named cells:
- `A1` as `CustomerName`
- `B1` as `CustomerAge`
- `C1` as `Score`

The function will read from those named ranges.

## NULL Handling

Empty or blank cells automatically become `NULL`:

```excel
=ExpandTemplate("$A, B, C")
```
Input: `A=John`, `B=(empty)`, `C=25`  
Output: `'John', NULL, 25`

To force empty string instead, use `~` prefix or set optional parameter.

## Special Character Escaping

The function automatically escapes special characters:

| Character | Input | Output |
|-----------|-------|--------|
| Single quote | `O'Brien` | `O''Brien` |
| Line feed (LF) | `Line1\nLine2` | `Line1\\nLine2` |
| Carriage return (CR) | `Text\r` | `Text\\r` |
| Tab | `Col1\tCol2` | `Col1\\t` |

## Advanced Features

### Optional Parameters

Function signature:
```vba
=ExpandTemplate(template, [nullForEmpty], [escapeStyle])
```

#### NULL for Empty (1st optional parameter)
```vba
=ExpandTemplate("$A, B", FALSE)
```

- `TRUE` (default) - Empty cells become `NULL`
- `FALSE` - Empty cells become `''` (empty string)

#### Escape Style (2nd optional parameter)
```vba
=ExpandTemplate("$A, B", TRUE, "SQL")
```

Supported styles:
- `"MySQL"` (default) - MySQL backslash escaping: `\'`, `\"`, `\\`, `\n`, `\r`, `\t`, `\b`, `\0`
- `"SQL"` - SQL Server / ANSI SQL: `''` for quotes, `\n`, `\r`, `\t`
- `"PostgreSQL"` - PostgreSQL escaping: `''` for quotes, `\n`, `\r`, `\t`

### Batch Processing Multiple Rows

Use `ExpandTemplateRange` to process multiple rows at once:

```excel
=ExpandTemplateRange("INSERT INTO users VALUES ($A, B, C);", A2:A10)
```

Or with optional parameters:
```excel
=ExpandTemplateRange("INSERT INTO users VALUES ($A, B, C);", A2:A10, TRUE, "SQL")
```

This returns an array formula with 9 rows of INSERT statements.

**To use:**
1. Select 9 cells vertically (e.g., `E2:E10`)
2. Type the formula
3. Press `Ctrl+Shift+Enter` (array formula)

## Complete Examples

### Example 1: Simple User Insert

**Data:**
| A (Name) | B (Age) | C (Email) |
|----------|---------|-----------|
| John | 25 | john@example.com |

**Formula:**
```excel
=ExpandTemplate("INSERT INTO users (name, age, email) VALUES ($A, B, $C);")
```

**Output:**
```sql
INSERT INTO users (name, age, email) VALUES ('John', 25, 'john@example.com');
```

### Example 2: Mixed Types with SQL Functions

**Data:**
| A (Name) | B (Created) | C (Status) |
|----------|-------------|------------|
| Alice | NOW() | 1 |

**Formula:**
```excel
=ExpandTemplate("INSERT INTO users (name, created_at, active) VALUES ($A, !B, ?C);")
```

**Output:**
```sql
INSERT INTO users (name, created_at, active) VALUES ('Alice', NOW(), 1);
```

### Example 3: Handling NULLs and Quotes

**Data:**
| A (Name) | B (Middle) | C (Note) |
|----------|------------|----------|
| O'Brien | (empty) | Said "Hi" |

**Formula:**
```excel
=ExpandTemplate("INSERT INTO contacts VALUES ($A, $B, $C);")
```

**Output:**
```sql
INSERT INTO contacts VALUES ('O''Brien', NULL, 'Said "Hi"');
```

### Example 4: Date Formatting

**Data:**
| A (Event) | B (Date) | C (Time) |
|-----------|----------|----------|
| Meeting | 2024-03-15 | 14:30:00 |

**Formula:**
```excel
=ExpandTemplate("INSERT INTO events (name, event_date) VALUES ($A, @B);")
```

**Output:**
```sql
INSERT INTO events (name, event_date) VALUES ('Meeting', '2024-03-15 00:00:00');
```

### Example 5: Using Beyond Column Z

**Data:**
| A | B | ... | AA | AB | AC |
|---|---|-----|----|----|-----|
| 1 | 2 | ... | 27 | 28 | 29 |

**Formula:**
```excel
=ExpandTemplate("VALUES (A, B, AA, AB, AC)")
```

**Output:**
```sql
VALUES (1, 2, 27, 28, 29)
```

## Tips & Best Practices

### 1. **Test with Small Datasets First**
Build your template on a few rows before applying to thousands.

### 2. **Use Named Ranges for Clarity**
Instead of `$A, $B, $C`, use `{FirstName}, {LastName}, {Email}`

### 3. **Combine with Excel Formulas**
```excel
=ExpandTemplate("INSERT INTO users VALUES ($A, B, @C);") & CHAR(10)
```
Adds a newline after each statement.

### 4. **Watch for SQL Injection**
This tool doesn't prevent SQL injection. Ensure source data is trusted or use parameterized queries in production.

### 5. **Performance with Large Datasets**
For thousands of rows, consider:
- Using `ExpandTemplateRange` for batch processing
- Copying results to a new sheet
- Running as a macro rather than live formulas

### 6. **Database-Specific Syntax**
Different databases have different requirements:
- **MySQL** (default): Uses backslash escaping `\'`, `\"`
- **SQL Server**: Use `"SQL"` escape style (doubles quotes: `''`)
- **PostgreSQL**: Use `"PostgreSQL"` escape style

Example:
```excel
=ExpandTemplate("$A, B", TRUE, "SQL")  ' For SQL Server
```

## Troubleshooting

| Error | Cause | Solution |
|-------|-------|----------|
| `#REF!` | Invalid column reference | Check column letter is valid (A-ZZZ) |
| `#ERROR:` | Formula syntax error | Verify template string format |
| Missing quotes | Wrong prefix used | Use `$` prefix for strings |
| `NULL` instead of value | Cell is empty | Check source cell has value |
| Wrong date format | Cell not formatted as date | Format cell as date or use `@` prefix |

## Installation

1. Open Excel and press `Alt+F11` to open VBA Editor
2. Insert → Module
3. Paste the code from the Enhanced SQL Template Expander
4. Save as `.xlsm` (macro-enabled workbook)
5. Use `=ExpandTemplate(...)` in any cell

## License & Support

This is a utility function for Excel VBA. Modify and extend as needed for your use case.

For questions or enhancements, refer to the inline code comments or extend the functions for your specific SQL dialect requirements.
