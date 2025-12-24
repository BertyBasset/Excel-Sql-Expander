# SQL Template Expander - Usage Guide

## Overview

The **SQL Template Expander** is a VBA utility for Excel that parses and evaluates a small **domain-specific language** embedded in template strings, expanding column references, literals, and symbolic tokens into fully formed SQL INSERT statements using data from worksheet rows.

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
=ExpandTemplate("!INSERT !INTO users (name, age, score) !VALUES ($A, B, C);")
```

Output: `INSERT INTO users (name, age, score) VALUES ('John', 25, 5.5);`

## Column Reference Case Sensitivity
Column references in templates are **case-sensitive** and must always be **uppercase**.

### Rules
- **Uppercase letters** (A–ZZ) are treated as **column references**
- **Lowercase letters** are treated as **literal text**
- Mixed-case sequences are parsed left-to-right

### Examples
| Template | Meaning                | Result (A=1)      |
| -------- | ---------------------- | ----------------- |
| `A`      | Column A               | `1`               |
| `a`      | Literal text           | `a`               |
| `AA`     | Column AA              | *(value from AA)* |
| `Aa`     | Column A + literal `a` | `1a`              |
| `aA`     | Literal `a` + column A | `a1`              |

### Practical Implications
- Always use uppercase column references (`A`, `B`, `AA`, `ZZ`)
- Lowercase letters are safe to use inside SQL text, identifiers, and literals
- If using upper case literals, prefix with literal marker `!`, or use lower case literals (which do not need `!` marker)
- This design avoids ambiguity between column references and literal text in SQL templates

### Example in Context
```excel
=ExpandTemplate("!INSERT !INTO log !VALUES (A, Aa, aA)")
=ExpandTemplate("insert into log values (A, Aa, aA)")
```
Input: `A=5`

Output:
```excel
INSERT INTO log VALUES (5, 5a, a5)
insert into log values (5, 5a, a5)
```
This behavior is intentional and forms part of the template language’s parsing rules.

## Referencing Template Text in a separate cell

It can often be useful to have the Template source in one cell, with the ExpandTemplate function being invoked on it from a different cell. Using this pattern, you can see both Template and expanded text at the same time. 
Taking the original example where cells are `A=John`, `B=25`, `C=5.5`  

We can then put the template expression into a cell in column D - say D1 

`$A, B, C`

In E1, we can refer to this using

`=ExpandTemplate(D1)`

E1 will then display the expanded output

`'John', 25, 5.5`

## Type Prefixes

Control how values are formatted using prefix characters before column letters:

| Prefix | Type | Example | Input: `A=123` | Output |
|--------|------|---------|----------------|--------|
| (none) | Auto-detect | `A` | `123` | `123` |
| `$` | String (force quotes) | `$A` | `123` | `'123'` |
| `#` | Numeric (force) | `#A` | `'123'` | `123` |
| `@` | Date/DateTime | `@A` | `2024-01-15` | `'2024-01-15 00:00:00'` |
| `!` | Literal text | `!NOW` | (any value) | `NOW` |
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

#### When A=1 (All Prefix Behaviors)

Here's what happens with cell A containing the value `1`:

| Formula | Output | Description |
|---------|--------|-------------|
| `=ExpandTemplate("A")` | `1` | Auto-detect (number) |
| `=ExpandTemplate("$A")` | `'1'` | Force string with quotes |
| `=ExpandTemplate("#A")` | `1` | Force numeric (unquoted) |
| `=ExpandTemplate("@A")` | `NULL` | Date format (invalid date) |
| `=ExpandTemplate("!A")` | `A` | Literal text "A" |
| `=ExpandTemplate("?A")` | `1` | Boolean (non-zero = true) |
| `=ExpandTemplate("~A")` | `1` | Number (never NULL) |

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

#### `!` - Literal Text (Column Names & SQL Keywords)
```excel
=ExpandTemplate("!NOW")
```
Output: `NOW`

```excel
=ExpandTemplate("!NOW()")
```
Output: `NOW()`

```excel
=ExpandTemplate("!DEFAULT")
```
Output: `DEFAULT`

```excel
=ExpandTemplate("!CURRENT_TIMESTAMP")
```
Output: `CURRENT_TIMESTAMP`

The `!` prefix outputs the text **literally** without reading any cell values. This is useful for:
- SQL function names: `!NOW()`, `!UUID()`, `!RAND()`
- SQL keywords: `!DEFAULT`, `!NULL`, `!CURRENT_TIMESTAMP`
- Column names in templates: `!user_id`, `!created_at`

**Important:** `!` does NOT read cell values. Use it only for literal SQL text.

**Special case - Mixed literal and values:**
```excel
=ExpandTemplate("!COALESCE( A, 0)")
```
Input: `A=5`  
Output: `COALESCE( 5, 0)`

When the literal text contains spaces followed by single capital letters (like ` A`), those letters are replaced with their cell values. Everything else remains literal.

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

Supports column references from **A to ZZ** (702 columns):

```excel
=ExpandTemplate("$A, $Z, $AA, $AB, $ZZ)
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

Assigning a name to a range of cells in a column allows you to reference that name across all rows in the range, making formulas and functions easier to read and maintain.


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

| Character | Input | Output (MySQL) |
|-----------|-------|----------------|
| Single quote | `O'Brien` | `O\'Brien` |
| Double quote | `Say "Hi"` | `Say \"Hi\"` |
| Backslash | `C:\path` | `C:\\path` |
| Line feed (LF) | `Line1\nLine2` | `Line1\\nLine2` |
| Carriage return (CR) | `Text\r` | `Text\\r` |
| Tab | `Col1\tCol2` | `Col1\\tCol2` |

*Note: Escaping style depends on the `escapeStyle` parameter (MySQL, SQL, or PostgreSQL)*

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
- `"SQL"` or `"SQLServer"` or `"ANSI"` - SQL Server / ANSI SQL: `''` for quotes, `\n`, `\r`, `\t`
- `"PostgreSQL"` or `"Postgres"` - PostgreSQL escaping: `''` for quotes, `\n`, `\r`, `\t`

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

### Example 1: Simple User Insert using lower case commands

**Data:**
| A (Name) | B (Age) | C (Email) |
|----------|---------|-----------|
| John | 25 | john@example.com |

**Formula:**
```excel
=ExpandTemplate("insert into users (name, age, email) VALUES ($A, B, $C);")
```

**Output:**
```sql
insert into users (name, age, email) VALUES ('John', 25, 'john@example.com');
```

### Example 2: Simple User Insert using upper case commands with literal prefix

**Data:**
| A (Name) | B (Age) | C (Email) |
|----------|---------|-----------|
| John | 25 | john@example.com |

**Formula:**
```excel
=ExpandTemplate("!INSERT !INTO users (name, age, email) VALUES ($A, B, $C);")
```

**Output:**
```sql
INSERT INTO users (name, age, email) VALUES ('John', 25, 'john@example.com');
```



### Example 3: Using SQL Functions with Literal Prefix

**Data:**
| A (Name) | B (Age) | C (Active) |
|----------|---------|------------|
| Alice | 25 | 1 |

**Formula:**
```excel
=ExpandTemplate("INSERT INTO users (name, age, created_at, active) VALUES ($A, B, !NOW(), ?C);")
```

**Output:**
```sql
INSERT INTO users (name, age, created_at, active) VALUES ('Alice', 25, NOW(), 1);
```

### Example 4: Handling NULLs and Quotes

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
INSERT INTO contacts VALUES ('O\'Brien', NULL, 'Said \"Hi\"');
```

### Example 5: Date Formatting

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

### Example 6: Using Beyond Column Z

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

### Example 7: Literal Text with Dynamic Values

**Formula:**
```excel
=ExpandTemplate("!COALESCE( A, !DEFAULT)")
```

**Input:** `A=100`  
**Output:** `COALESCE( 100, DEFAULT)`

The space before `A` causes it to be replaced with the cell value, while `DEFAULT` (preceded by `!`) remains literal.

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

### 7. **Understanding the `!` Literal Prefix**
- Use `!` for SQL keywords and function names: `!NOW()`, `!DEFAULT`, `!CURRENT_TIMESTAMP`
- Do NOT use `!A` expecting it to read cell A - it will output literal "A"
- For mixing literals with values, use space separation: `!COALESCE( A, 0)` where A gets replaced

## Troubleshooting

| Error | Cause | Solution |
|-------|-------|----------|
| `#REF!` | Invalid column reference | Check column letter is valid (A-ZZ) |
| `#ERROR:` | Formula syntax error | Verify template string format |
| Missing quotes | Wrong prefix used | Use `$` prefix for strings |
| `NULL` instead of value | Cell is empty | Check source cell has value |
| Wrong date format | Cell not formatted as date | Format cell as date or use `@` prefix |
| Literal `A` instead of value | Used `!A` prefix | Remove `!` to read cell value: use `A` instead |
| Function name quoted | Missing `!` prefix | Use `!NOW()` not `NOW()` |

## Installation

1. Open Excel and press `Alt+F11` to open VBA Editor
2. Insert → Module
3. Paste the code from the Enhanced SQL Template Expander
4. Save as `.xlsm` (macro-enabled workbook)
5. Use `=ExpandTemplate(...)` in any cell

## License & Support

This is a utility function for Excel VBA. Modify and extend as needed for your use case.

For questions or enhancements, refer to the inline code comments or extend the functions for your specific SQL dialect requirements.
