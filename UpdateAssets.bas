Attribute VB_Name = "Module2"
Option Explicit

Sub UpdateAssetsFromSelectedMonth()
    Dim wsA As Worksheet, wsM As Worksheet
    Dim y As String, m As String, sheetName As String
    Dim rStart As Long, rEnd As Long, rLogEnd As Long
    Dim i As Long, j As Long, amt As Double
    Dim fromAcc As String, toAcc As String
    Dim openBal As Double

    Set wsA = ThisWorkbook.Sheets("Assets")
    y = Trim(wsA.Range("B1").Value)
    m = Trim(wsA.Range("B2").Value)
    sheetName = y & "-" & m

    If y = "" Or m = "" Then
        MsgBox "?? Please select Year (B1) and Month (B2).", vbExclamation
        Exit Sub
    End If
    If Not SheetExists(sheetName) Then
        MsgBox "? Sheet '" & sheetName & "' not found.", vbExclamation
        Exit Sub
    End If

    Set wsM = ThisWorkbook.Sheets(sheetName)
    rStart = wsA.Columns("A").Find("Category", , , xlWhole).Row + 1
    rEnd = wsA.Cells(wsA.Rows.Count, "B").End(xlUp).Row
    rLogEnd = wsM.Cells(wsM.Rows.Count, "L").End(xlUp).Row

    ' Reset current balances
    wsA.Range("D" & rStart & ":D" & rEnd).Value = wsA.Range("C" & rStart & ":C" & rEnd).Value

    ' Apply income / expenses
    For i = 2 To rLogEnd
        If IsNumeric(wsM.Cells(i, 9)) Then amt = wsM.Cells(i, 9).Value Else amt = 0
        fromAcc = Trim(wsM.Cells(i, 12).Value)
        toAcc = Trim(wsM.Cells(i, 13).Value)

        If fromAcc <> "" And LCase(fromAcc) <> "nil" Then
            For j = rStart To rEnd
                If wsA.Cells(j, 2).Value = fromAcc Then
                    wsA.Cells(j, 4).Value = wsA.Cells(j, 4).Value - amt
                    Exit For
                End If
            Next j
        End If
        If toAcc <> "" And LCase(toAcc) <> "nil" Then
            For j = rStart To rEnd
                If wsA.Cells(j, 2).Value = toAcc Then
                    wsA.Cells(j, 4).Value = wsA.Cells(j, 4).Value + amt
                    Exit For
                End If
            Next j
        End If
    Next i

    ' Calculate Change (£) and (%)
    For j = rStart To rEnd
        If IsNumeric(wsA.Cells(j, 3)) And IsNumeric(wsA.Cells(j, 4)) Then
            openBal = wsA.Cells(j, 3).Value
            wsA.Cells(j, 5).Value = wsA.Cells(j, 4).Value - openBal
            If openBal <> 0 Then
                wsA.Cells(j, 6).Value = wsA.Cells(j, 5).Value / openBal
            Else
                wsA.Cells(j, 6).Value = vbNullString
            End If
        End If
    Next j

    ' Format & highlight
    wsA.Range("C" & rStart & ":E" & rEnd).NumberFormat = "£#,##0.00"
    wsA.Range("F" & rStart & ":F" & rEnd).NumberFormat = "0.0%"
    
    For j = rStart To rEnd
        Dim changeVal As Double, changePct As Double
        changeVal = wsA.Cells(j, 5).Value   ' Change (£)
        changePct = wsA.Cells(j, 6).Value   ' Change (%)
        
        wsA.Range("E" & j & ":F" & j).Interior.ColorIndex = xlNone
        wsA.Range("E" & j & ":F" & j).Font.Color = vbBlack
        
        If changeVal < -1500 Then
            wsA.Cells(j, 5).Interior.Color = vbRed
            wsA.Cells(j, 5).Font.Color = vbWhite
        End If

        If changePct < -0.9 Then
            wsA.Cells(j, 6).Interior.Color = vbRed
            wsA.Cells(j, 6).Font.Color = vbWhite
        End If
    Next j
    
    ' Totals
    wsA.Cells(rEnd + 1, 3).Value = WorksheetFunction.Sum(wsA.Range("C" & rStart & ":C" & rEnd))
    wsA.Cells(rEnd + 1, 4).Value = WorksheetFunction.Sum(wsA.Range("D" & rStart & ":D" & rEnd))
    wsA.Cells(rEnd + 1, 3).NumberFormat = "£#,##0.00"
    wsA.Cells(rEnd + 1, 4).NumberFormat = "£#,##0.00"

    ' Auto-save workbook
    ThisWorkbook.Save

    MsgBox "Assets updated and saved (" & sheetName & ")", vbInformation
End Sub
Private Function SheetExists(s As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Sheets(s) Is Nothing
    On Error GoTo 0
End Function


