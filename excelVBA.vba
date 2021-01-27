
Public Sub CopyData()
    'define source range
    Dim SourceRange As Range
    Set SourceRange = ThisWorkbook.Worksheets("Parser").Range("A2:AD2")
    Set ClearRange = ThisWorkbook.Worksheets("Parser").Range("C3:F1000")
    
    If Len(Range("B11").Value) = 1 Then
    MsgBox "Please fill in the BRAND field."
    Exit Sub
    End If

    'find next free cell in destination sheet
    Dim NextFreeCell As Range
    Set NextFreeCell = ThisWorkbook.Worksheets("WorkLog").Cells(Rows.Count, "A").End(xlUp).Offset(RowOffset:=1)

    'copy & paste
    SourceRange.Copy
    NextFreeCell.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ThisWorkbook.Save
    ClearRange.ClearContents
    Sheet3.Range("B19").Value = ""
    Sheet3.Range("B20").Value = "0.00"
    Sheet3.Range("B21").Value = "Other Autoreleased"
    Sheet3.Range("B22").Value = "No"
    Sheet3.Range("B23").Value = "No"
    Sheet3.Range("B24").Value = "=IF($B$23=""Yes""&$B$17=""PayPal"",""Responding to Request for Info"","""")"
    Sheet3.Range("B32").Value = ""
    Sheet3.Range("B26:B27").Value = ""
    Sheet3.Range("B3").Formula = "=IF($B$17=""Chase"",Formulas!$B2,IF($B$17=""PayPal"",Formulas!$C2,IF($B$17=""Amex"",Formulas!$D2,IF($B$17=""Adyen"",Formulas!$E2,IF($B$17=""JCP"",Formulas!$F2)))))"
    Sheet3.Range("B4").Formula = "=IF($B$17=""Chase"",Formulas!$B3,IF($B$17=""PayPal"",Formulas!$C3,IF($B$17=""Amex"",Formulas!$D3,IF($B$17=""Adyen"",Formulas!$E3,IF($B$17=""JCP"",Formulas!$F3)))))"
    Sheet3.Range("B11").Formula = "=IF($B$17=""Chase"",Formulas!$B10,IF($B$17=""PayPal"",Formulas!$C10,IF($B$17=""Amex"",Formulas!$D10,IF($B$17=""Adyen"",Formulas!$E10,IF($B$17=""JCP"",Formulas!$F10)))))"
    Sheet3.Range("B26").Formula = "=IF($B$24=""Prior Credit"",""[proof of credit - CyberSource screenshot]"","""")"
    
    
    End Sub
 

