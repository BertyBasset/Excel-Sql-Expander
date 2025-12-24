' ========================================================================
' Enhanced SQL Template Expander for Excel VBA
' ========================================================================
' Builds SQL INSERT queries by expanding template strings with cell values
' Supports multiple column formats (A-ZZ), type prefixes, and advanced features
' ========================================================================

Option Explicit

' Main function: Expands template with values from current row
Function ExpandTemplate(template As String, Optional nullForEmpty As Boolean = True, Optional escapeStyle As String = "MySQL") As String
    Application.Volatile
    On Error GoTo ErrorHandler
    
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    re.Global = True
    re.Pattern = "(![^, ]+|[#$@?~]?([A-Z]{1,2}|{[A-Za-z_][A-Za-z0-9_]*}))"



    
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
        
        Dim token As String
        token = matches(i).Value
        
        Dim firstChar As String
        firstChar = Left(token, 1)
        
        Select Case firstChar
            Case "!", "#", "$", "@", "?", "~"
                prefix = firstChar
                colRef = Mid(token, 2)
            Case Else
                prefix = ""          ' NO prefix
                colRef = token       ' FULL token is the column
        End Select


        ' Handle literal prefix "!" - output literal text with optional cell value substitution
        If prefix = "!" Then
            replacement = ProcessLiteralTemplate(colRef, ws, targetRow)
        ElseIf Left(colRef, 1) = "{" Then
            colNum = GetNamedRangeColumn(Mid(colRef, 2, Len(colRef) - 2), ws)
        Else
            If Not colCache.Exists(colRef) Then
                colCache.Add colRef, ColumnLetterToNumber(colRef)
            End If
            colNum = colCache(colRef)
        End If
        
        ' Only process non-literal prefixes
        If prefix <> "!" Then
            If colNum = 0 Then
                replacement = "#REF!"
            Else
                cellVal = ws.Cells(targetRow, colNum).Value
                replacement = FormatValue(cellVal, prefix, escapeStyle, nullForEmpty, ws.Cells(targetRow, colNum))
            End If
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

' Helper: Process literal template with optional cell value substitution
Private Function ProcessLiteralTemplate(template As String, ws As Worksheet, targetRow As Long) As String
    ' If template is just a column reference (like "A", "AB", "ABC"), return it as literal
    ' Otherwise, replace standalone single capital letters after spaces with their cell values
    
    Dim result As String
    result = template
    
    ' Check if the entire template is just uppercase letters (column reference)
    Dim allAlpha As Boolean
    allAlpha = True
    Dim i As Long
    
    If Len(template) > 0 And Len(template) <= 3 Then
        For i = 1 To Len(template)
            If Not Mid(template, i, 1) Like "[A-Z]" Then
                allAlpha = False
                Exit For
            End If
        Next i
        If allAlpha Then
            ProcessLiteralTemplate = template
            Exit Function
        End If
    End If
    
    ' For templates with spaces, manually find and replace single letters after spaces
    If InStr(template, " ") > 0 Then
        For i = 2 To Len(template) ' Start at 2 since we need a preceding character
            ' Check if current char is uppercase letter, preceded by space, followed by space or end
            If Mid(template, i, 1) Like "[A-Z]" And Mid(template, i - 1, 1) = " " Then
                Dim nextChar As String
                If i < Len(template) Then
                    nextChar = Mid(template, i + 1, 1)
                Else
                    nextChar = ""
                End If
                
                ' If followed by space or end of string, replace it
                If nextChar = " " Or nextChar = "" Then
                    Dim colLetter As String
                    Dim colNum As Long
                    Dim cellVal As Variant
                    
                    colLetter = Mid(template, i, 1)
                    colNum = ColumnLetterToNumber(colLetter)
                    
                    If colNum > 0 Then
                        cellVal = ws.Cells(targetRow, colNum).Value
                        If IsEmpty(cellVal) Then
                            result = Left(result, i - 1) & Mid(result, i + 1)
                            i = i - 1 ' Adjust position after deletion
                        Else
                            Dim replacement As String
                            replacement = CStr(cellVal)
                            result = Left(result, i - 1) & replacement & Mid(result, i + 1)
                            i = i + Len(replacement) - 1 ' Adjust position after replacement
                        End If
                    End If
                End If
            End If
        Next i
    End If
    
    ProcessLiteralTemplate = result
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
    re.Pattern = "(![^, ]+|[#$@?~]?([A-Z]{1,2}|{[A-Za-z_][A-Za-z0-9_]*}))"


    
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
        
        Dim token As String
        token = matches(i).Value
        
        Dim firstChar As String
        firstChar = Left(token, 1)
        
        Select Case firstChar
            Case "!", "#", "$", "@", "?", "~"
                prefix = firstChar
                colRef = Mid(token, 2)
            Case Else
                prefix = ""          ' NO prefix
                colRef = token       ' FULL token is the column
        End Select

        
        ' Handle literal prefix "!" - output literal text with optional cell value substitution
        If prefix = "!" Then
            replacement = ProcessLiteralTemplate(colRef, ws, targetRow)
        ElseIf Left(colRef, 1) = "{" Then
            colNum = GetNamedRangeColumn(Mid(colRef, 2, Len(colRef) - 2), ws)
        Else
            colNum = ColumnLetterToNumber(colRef)
        End If
        
        ' Only process non-literal prefixes
        If prefix <> "!" Then
            If colNum = 0 Then
                replacement = "#REF!"
            Else
                cellVal = ws.Cells(targetRow, colNum).Value
                replacement = FormatValue(cellVal, prefix, escapeStyle, nullForEmpty, ws.Cells(targetRow, colNum))
            End If
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
    If prefix = "~" And cellVal = "" Then
        FormatValue = "''"
        Exit Function
    End If
    
    
    
    If IsEmpty(cellVal) Or Trim(cellVal & "") = "" Then
        If nullForEmpty Or prefix = "#" Or prefix = "$" Or prefix = "@" Or prefix = "?" Then
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

' Helper: Convert column letter(s) to number (supports A-ZZ)
Private Function ColumnLetterToNumber(colLetter As String) As Long
    Dim result As Long
    Dim i As Long

    ' Enforce A–ZZ only
    If Len(colLetter) < 1 Or Len(colLetter) > 2 Then
        ColumnLetterToNumber = 0
        Exit Function
    End If

    result = 0
    For i = 1 To Len(colLetter)
        result = result * 26 + (Asc(Mid(colLetter, i, 1)) - Asc("A") + 1)
    Next i

    ' Final numeric guard (ZZ = 702)
    If result < 1 Or result > 702 Then
        ColumnLetterToNumber = 0
    Else
        ColumnLetterToNumber = result
    End If
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

