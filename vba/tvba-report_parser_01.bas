Attribute VB_Name = "Module1"
Option Explicit

Public Sub test01()

' loop thru column
' define output target
' define input keyword
' if matched, debug output matched range info
'  --> next add link


End Sub

Public Function find_first_match_row(ByVal searchText As String, ByVal src_lastrow As Long, ByVal src As Worksheet, Optional os As Integer = 0) As Integer
Dim i, j As Integer
j = 2
For i = 2 To src_lastrow
    If src.Range("A" & i).Cells.Interior.color = RGB(68, 114, 196) And src.Range("A" & i).Cells.Font.color = RGB(255, 255, 255) Then
        If InStr(src.Range("A" & i).Cells.Offset(0, os).Text, searchText) Then
            find_first_match_row = j
            Exit For
        End If
        j = j + 1
    End If
Next i
End Function

Sub RealTest()
Attribute RealTest.VB_ProcData.VB_Invoke_Func = "e\n14"

Dim src_lastrow, dest_lastrow As Long
Dim i, j As Long
Dim src, dest As Worksheet

    Set src = Sheets("source")
    src_lastrow = src.Cells(Rows.Count, "A").End(xlUp).Row
    
    Set dest = Sheets("output")
    
    dest.Range("A2:D200").Clear
    'src.Range("B2:B200").Clear
    
' find the first curly bracket



Dim os As Integer

'os_curly = find_first_match_row("}", src_lastrow, src, 1)
'os_curvetrace = find_first_match_row("Curve Trace (2)", src_lastrow, src)
'os_ipos = find_first_match_row("{Latchup", src_lastrow, src, 1)

'dest.Range("C" & os_ineg & ":C" & os_ineg + 10) = "INEG"
'dest.Range("C" & os_ineg & ":C" & os_ineg + 10).Interior.color = RGB(237, 125, 49)

os = 3

j = 2
'MsgBox (src_lastrow)
For i = 2 To src_lastrow
    If src.Range("A" & i).Cells.Interior.color = RGB(68, 114, 196) And src.Range("A" & i).Cells.Font.color = RGB(255, 255, 255) Then

        'src.Hyperlinks.Add src.Range("A" & i).Cells.Offset(0, 1), Address:="", SubAddress:="output!A" & j, TextToDisplay:=dest.Range("A" & j).Cells.Text
        
        'dest link formatting
        If InStr(src.Range("A" & i).Text, "Vs") Then
            'output to col B
            dest.Hyperlinks.Add dest.Range("A" & j).Cells.Offset(0, 1), Address:="", SubAddress:="source!A" & i, TextToDisplay:=src.Range("A" & i).Text & src.Range("A" & i).Offset(0, 1).Text
            src.Hyperlinks.Add src.Range("A" & i).Cells.Offset(0, os), Address:="", SubAddress:="output!B" & j, TextToDisplay:="Summary"
        Else
            dest.Hyperlinks.Add dest.Range("A" & j).Cells, Address:="", SubAddress:="source!A" & i, TextToDisplay:=src.Range("A" & i).Text & src.Range("A" & i).Offset(0, 1).Text
            src.Hyperlinks.Add src.Range("A" & i).Cells.Offset(0, os), Address:="", SubAddress:="output!A" & j, TextToDisplay:="Summary"
        End If
        
        If InStr(src.Range("A" & i).Text, "Setup Stress Levels") Or InStr(src.Range("A" & i).Text, "[spot 1]") Then
            dest.Range("A" & j).Cells.Offset(0, 3) = "main test result"
        End If
        
        
        
        'decorate src cell
        src.Range("A" & i).Offset(0, os).Cells.Interior.color = RGB(68, 114, 196)
        src.Range("A" & i).Offset(0, os).Cells.Font.color = RGB(255, 255, 255)
        
        j = j + 1
    End If
Next i

dest_lastrow = dest.Cells(Rows.Count, "A").End(xlUp).Row
Dim os_temp As Integer

Dim k As Integer
k = 1

For j = 2 To dest_lastrow
    
    If InStr(dest.Range("A" & j).Text, "Metadata") Then
    '0, -9, BG Info, RGB(217, 217, 217)
    dest.Range("C" & j & ":C" & j + 9).Cells.Interior.color = RGB(217, 217, 217)
    dest.Range("C" & j & ":C" & j + 3).Cells = "BG Info:"
    dest.Range("C" & j + 4 & ":C" & j + 9).Cells = "BG Info: Test Flow"
    dest.Range("C" & j + 4 & ":C" & j + 9).Cells.Offset(0, 1) = "main test flow"
    
    '+1, +5, Continuity, 146, 208, 80
    dest.Range("C" & j + 10 & ":C" & j + 14).Cells.Interior.color = RGB(146, 208, 80)
    dest.Range("C" & j + 10 & ":C" & j + 14).Cells = "Continuity"
    
    ElseIf InStr(dest.Range("A" & j).Text, "Pre-curvetrace{") Then
    '0, +3, Pre-curvetrace, (255, 192, 0)
    dest.Range("C" & j & ":C" & j + 3).Cells.Interior.color = RGB(255, 192, 0)
    dest.Range("C" & j & ":C" & j + 3).Cells = "Pre-curvetrace"
    dest.Range("C" & j).Cells.Offset(3, 1) = "main test result"
    
    ElseIf InStr(dest.Range("A" & j).Text, "Setup Stress Levels{") Then
    '0, +11, IPOS/INEG, (91, 155, 213), (237, 125, 49)
    
    If k Mod 2 = 1 Then
        dest.Range("C" & j - 11 & ":C" & j).Cells.Interior.color = RGB(91, 155, 213)
        dest.Range("C" & j - 11 & ":C" & j - 1).Cells = "IPOS"
        dest.Range("C" & j).Cells = "IPOS: RESULT"
    Else
        dest.Range("C" & j - 11 & ":C" & j).Cells.Interior.color = RGB(237, 125, 49)
        dest.Range("C" & j - 11 & ":C" & j - 1).Cells = "INEG"
        dest.Range("C" & j).Cells = "INEG: RESULT"
    End If
        dest.Range("C" & j).Cells.Offset(0, 1) = "main test result"
    k = k + 1
    
    ElseIf InStr(dest.Range("A" & j).Text, "Post-curvetrace{") Then
    '0, +3, Post-curvetrace, (191, 143, 0)
    dest.Range("C" & j & ":C" & j + 3).Cells.Interior.color = RGB(191, 143, 0)
    dest.Range("C" & j & ":C" & j + 3).Cells = "Post-curvetrace"
    
    ElseIf InStr(dest.Range("A" & j).Text, "Continuity END{") Then
    '0, +4, Continuity END, (237, 125, 49)
    dest.Range("C" & j & ":C" & j + 4).Cells.Interior.color = RGB(237, 125, 49)
    dest.Range("C" & j & ":C" & j + 4).Cells = "Continuity END"
    
    Else
    
    End If

Next j
'Call FilterAndCopy
Set dest = Sheets("output")


dest.AutoFilter.ShowAllData
dest.Range("A1:D200").AutoFilter Field:=4, Criteria1:=Array("main test result", "main test flow"), Operator:=xlFilterValues


End Sub



Sub FilterAndCopy()
Dim LastRow As Long

Sheets("output").UsedRange.Offset(0).ClearContents
    With Worksheets("ouput")
        .Range("$A:$D").AutoFilter
        '.Range("$A:$D").AutoFilter field:=4, Criteria1:="<-- main test result"
        '.Range("$A:$E").AutoFilter field:=2, Criteria1:="=String1", Operator:=xlOr, Criteria2:="=string2"
        '.Range("$A:$E").AutoFilter field:=3, Criteria1:=">0"
        '.Range("$A:$E").AutoFilter field:=5, Criteria1:="Number"
        'LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
        '.Range("A1:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Copy _
        '        Destination:=Sheets("Sheet2").Range("A1")
    End With
End Sub

Sub Test()

Dim src_lastrow, dest_lastrow As Long
Dim i, j As Long
Dim src, dest As Worksheet

    Set src = Sheets("source")
    src_lastrow = src.Cells(Rows.Count, "A").End(xlUp).Row
    
    Set dest = Sheets("output")
    dest_lastrow = dest.Cells(Rows.Count, "A").End(xlUp).Row
    'MsgBox (lastrow)
    
    'For j = 2 To dest_lastrow
    '    For i = 2 To src_lastrow
    '        dest.Range("A" & j).Cells.Offset(0, 1).Clear
    '        src.Range("A" & i).Cells.Offset(0, 1).Clear
    '    Next i
    'Next j
    
    'MsgBox (dest_lastrow)
    'MsgBox (src_lastrow)
    
    'MsgBox ("B2:B" & dest_lastrow)
    'MsgBox ("B2:B" & src_lastrow)
    
    dest.Range("A2:A200").Clear
    src.Range("B2:B200").Clear
    
    
    '
    '
    '
    '
For i = 2 To src_lastrow
    For j = 2 To dest_lastrow
            'MsgBox (src.Range("A" & i))
            'MsgBox ("A" & i)
        If src.Range("A" & i) = dest.Range("A" & j) Then
            dest.Range("A" & j).Cells.Offset(0, 1) = dest.Range("A" & j)
            dest.Hyperlinks.Add dest.Range("A" & j).Cells.Offset(0, 1), Address:="", SubAddress:="source!A" & i, TextToDisplay:=Range("A" & j).Text
            src.Hyperlinks.Add src.Range("A" & i).Cells.Offset(0, 1), Address:="", SubAddress:="output!A" & j, TextToDisplay:="Summary"
            'src.Hyperlinks.Add src.Range("A" & i).Cells.Offset(0, 1), Address:="", SubAddress:="output!A" & j, TextToDisplay:=CStr(Range("A" & i).Cells.Text)
                'MsgBox ("A" & i)
                
                'activeSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        '"source!A2"
        End If
    Next j
Next i
    
End Sub


'fill RGB(68, 114, 196)
'font RGB(255, 255, 255)
