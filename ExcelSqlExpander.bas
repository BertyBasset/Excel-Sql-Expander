' ========================================================================
' Enhanced SQL Template Expander for Excel VBA
' ========================================================================
' Builds SQL INSERT queries by expanding template strings with cell values
' Supports multiple column formats (A-ZZZ), type prefixes, and advanced features
' ========================================================================

Option Explicit

' Main function: Expands template with values from current row
Function ExpandTemplate(template As String, Optional nullForEmpty As Boolean = True, Optional escapeStyle As String = "MySQL") As String
    On Error GoTo ErrorHandler
    
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    re.Global = True
    ' Pattern: (prefix)(column) or (prefix){named_range}
    re.Pattern = "([#$@!?~]?)([A-Z]{1,2}|{[A-Za-z_][A-Za-z0-9_]*})"
    
    Dim ws As Worksheet
    Dim targetRow As Long
    
    Set ws = Application.Caller.Worksheet
    targetRow = Application.Caller.Row
    
    Dim result As String
    result = template
    
    Dim matches As Object
    Set matches = re.Execute(template)
    
    ' Cache for column numbers to improve performance
    Dim colCache As Object
    Set colCache = CreateObject("Scripting.Dictionary")
    
    ' Process matches in reverse to maintain string positions
    Dim i As Long
    For i = matches.Count - 1 To 0 Step -1
        Dim prefix As String
        Dim colRef As String
        Dim colNum As Long
        Dim cellVal As Variant
        Dim replacement As String
        
        prefix = matches(i).SubMatches(0)
        colRef = matches(i).SubMatches(1)
        
        ' Handle named ranges vs column letters
        If Left(colRef, 1) = "{" Then
            colNum = GetNamedRangeColumn(Mid(colRef, 2, Len(colRef) - 2), ws)
        Else
            If Not colCache.Exists(colRef) Then
                colCache.Add colRef, ColumnLetterToNumber(colRef)
            End If
            colNum = colCache(colRef)
        End If
        
        If colNum = 0 Then
            replacement = "#REF!"
        Else
            cellVal = ws.Cells(targetRow, colNum).Value
            replacement = FormatValue(cellVal, prefix, escapeStyle, nullForEmpty, ws.Cells(targetRow, colNum))
        End If
        
        ' Replace in result string
        result = Left(result, matches(i).FirstIndex) & _
                 replacement & _
                 Mid(result, matches(i).FirstIndex + matches(i).Length + 1)
    Next i
    
    ExpandTemplate = result
    Exit Function
    
ErrorHandler:
    ExpandTemplate = "#ERROR: " & Err.Description
End Function

' Batch function: Expands template for multiple rows
Function ExpandTemplateRange(template As String, rowRange As Range, Optional nullForEmpty As Boolean = True, Optional escapeStyle As String = "MySQL") As Variant
    On Error GoTo ErrorHandler
    
    Dim results() As String
    Dim rowCount As Long
    Dim i As Long
    
    rowCount = rowRange.Rows.Count
    ReDim results(1 To rowCount, 1 To 1)
    
    Dim ws As Worksheet
    Set ws = rowRange.Worksheet
    
    For i = 1 To rowCount
        results(i, 1) = ExpandTemplateForRow(template, ws, rowRange.Rows(i).Row, escapeStyle, nullForEmpty)
    Next i
    
    ExpandTemplateRange = results
    Exit Function
    
ErrorHandler:
    ExpandTemplateRange = "#ERROR: " & Err.Description
End Function

' Helper: Expand template for specific row
Private Function ExpandTemplateForRow(template As String, ws As Worksheet, targetRow As Long, escapeStyle As String, nullForEmpty As Boolean) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    re.Global = True
    re.Pattern = "([#$@!?~]?)([A-Z]{1,2}|{[A-Za-z_][A-Za-z0-9_]*})"
    
    Dim result As String
    result = template
    
    Dim matches As Object
    Set matches = re.Execute(template)
    
    Dim i As Long
    For i = matches.Count - 1 To 0 Step -1
        Dim prefix As String
        Dim colRef As String
        Dim colNum As Long
        Dim cellVal As Variant
        Dim replacement As String
        
        prefix = matches(i).SubMatches(0)
        colRef = matches(i).SubMatches(1)
        
        If Left(colRef, 1) = "{" Then
            colNum = GetNamedRangeColumn(Mid(colRef, 2, Len(colRef) - 2), ws)
        Else
            colNum = ColumnLetterToNumber(colRef)
        End If
        
        If colNum = 0 Then
            replacement = "#REF!"
        Else
            cellVal = ws.Cells(targetRow, colNum).Value
            replacement = FormatValue(cellVal, prefix, escapeStyle, nullForEmpty, ws.Cells(targetRow, colNum))
        End If
        
        result = Left(result, matches(i).FirstIndex) & _
                 replacement & _
                 Mid(result, matches(i).FirstIndex + matches(i).Length + 1)
    Next i
    
    ExpandTemplateForRow = result
End Function

' Helper: Format value based on prefix and type
Private Function FormatValue(cellVal As Variant, prefix As String, escapeStyle As String, nullForEmpty As Boolean, cell As Range) As String
    Dim result As String
    
    ' Handle empty/null values
    If IsEmpty(cellVal) Or Trim(cellVal & "") = "" Then
        If nullForEmpty Or prefix = "" Or prefix = "#" Or prefix = "$" Or prefix = "@" Or prefix = "?" Then
            FormatValue = "NULL"
        Else ' prefix = "~"
            FormatValue = "''"
        End If
        Exit Function
    End If
    
    ' Escape special characters based on style
    Dim escapedVal As String
    escapedVal = EscapeString(CStr(cellVal), escapeStyle)
    
    ' Format based on prefix (prefix overrides Excel type)
    Select Case prefix
        Case "#"  ' Numeric (force)
            result = CStr(cellVal)
            
        Case "$"  ' String (force)
            result = "'" & escapedVal & "'"
            
        Case "@"  ' Date/DateTime
            If IsDate(cellVal) Then
                result = "'" & Format(cellVal, "yyyy-mm-dd hh:nn:ss") & "'"
            Else
                result = "NULL"
            End If
            
        Case "!"  ' Raw/Literal (no escaping, for SQL functions)
            result = CStr(cellVal)
            
        Case "?"  ' Boolean
            If IsNumeric(cellVal) Then
                result = IIf(CDbl(cellVal) <> 0, "1", "0")
            ElseIf VarType(cellVal) = vbBoolean Then
                result = IIf(CBool(cellVal), "1", "0")
            ElseIf UCase(Trim(CStr(cellVal))) = "TRUE" Then
                result = "1"
            ElseIf UCase(Trim(CStr(cellVal))) = "FALSE" Then
                result = "0"
            Else
                result = "NULL"
            End If
            
        Case "~"  ' Force empty string instead of NULL
            result = "'" & escapedVal & "'"
            
        Case Else  ' Auto-detect based on Excel cell type
            result = AutoDetectFormat(cellVal, escapedVal, cell)
    End Select
    
    FormatValue = result
End Function

' Helper: Auto-detect format based on Excel cell type
Private Function AutoDetectFormat(cellVal As Variant, escapedVal As String, cell As Range) As String
    Dim result As String
    Dim numFormat As String
    
    ' Get the cell's number format
    numFormat = cell.NumberFormat
    
    ' Check if cell is explicitly formatted as text
    If numFormat = "@" Then
        result = "'" & escapedVal & "'"
        AutoDetectFormat = result
        Exit Function
    End If
    
    ' Check for date/time formats
    If IsDate(cellVal) Then
        ' Common date format indicators in NumberFormat
        If InStr(1, numFormat, "d", vbTextCompare) > 0 Or _
           InStr(1, numFormat, "m", vbTextCompare) > 0 Or _
           InStr(1, numFormat, "y", vbTextCompare) > 0 Or _
           InStr(1, numFormat, "h", vbTextCompare) > 0 Or _
           numFormat Like "*[$-*]*" Then ' Date formats often contain locale info
            result = "'" & Format(cellVal, "yyyy-mm-dd hh:nn:ss") & "'"
            AutoDetectFormat = result
            Exit Function
        End If
    End If
    
    ' Check for numeric formats (including currency, percentage, scientific)
    If IsNumeric(cellVal) And Not IsDate(cellVal) Then
        result = CStr(cellVal)
    Else
        result = "'" & escapedVal & "'"
    End If
    
    AutoDetectFormat = result
End Function

' Helper: Escape special characters in strings
Private Function EscapeString(str As String, escapeStyle As String) As String
    Dim result As String
    result = str
    
    Select Case UCase(escapeStyle)
        Case "SQL", "SQLSERVER", "ANSI"
            ' SQL standard: single quotes doubled
            result = Replace(result, "'", "''")
            result = Replace(result, Chr(10), "\n")  ' Line feed
            result = Replace(result, Chr(13), "\r")  ' Carriage return
            result = Replace(result, Chr(9), "\t")   ' Tab
            
        Case "MYSQL"
            ' MySQL: backslash escaping
            result = Replace(result, "\", "\\")
            result = Replace(result, "'", "\'")
            result = Replace(result, """", "\""")
            result = Replace(result, Chr(0), "\0")
            result = Replace(result, Chr(10), "\n")
            result = Replace(result, Chr(13), "\r")
            result = Replace(result, Chr(9), "\t")
            result = Replace(result, Chr(8), "\b")   ' Backspace
            
        Case "POSTGRESQL", "POSTGRES"
            ' PostgreSQL: single quotes doubled (like SQL standard)
            result = Replace(result, "'", "''")
            result = Replace(result, Chr(10), "\n")
            result = Replace(result, Chr(13), "\r")
            result = Replace(result, Chr(9), "\t")
            
        Case Else
            ' Default: MySQL
            result = Replace(result, "\", "\\")
            result = Replace(result, "'", "\'")
            result = Replace(result, """", "\""")
            result = Replace(result, Chr(0), "\0")
            result = Replace(result, Chr(10), "\n")
            result = Replace(result, Chr(13), "\r")
            result = Replace(result, Chr(9), "\t")
            result = Replace(result, Chr(8), "\b")
    End Select
    
    EscapeString = result
End Function

' Helper: Convert column letter(s) to number (supports A-ZZZ)
Private Function ColumnLetterToNumber(colLetter As String) As Long
    On Error GoTo ErrorHandler
    
    Dim result As Long
    Dim i As Integer
    Dim char As String
    
    result = 0
    For i = 1 To Len(colLetter)
        char = Mid(colLetter, i, 1)
        result = result * 26 + (Asc(char) - Asc("A") + 1)
    Next i
    
    ColumnLetterToNumber = result
    Exit Function
    
ErrorHandler:
    ColumnLetterToNumber = 0
End Function

' Helper: Get column number from named range
Private Function GetNamedRangeColumn(rangeName As String, ws As Worksheet) As Long
    On Error Resume Next
    
    Dim rng As Range
    Set rng = ws.Parent.Names(rangeName).RefersToRange
    
    If Not rng Is Nothing Then
        GetNamedRangeColumn = rng.Column
    Else
        GetNamedRangeColumn = 0
    End If
End Function

