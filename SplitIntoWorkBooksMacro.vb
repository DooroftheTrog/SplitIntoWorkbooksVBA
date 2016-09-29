Sub SplitIntoSheetsMacro()
'
' SplitIntoSheetsMacro
'
 Dim i As Integer
    
    For i = 321 To 515
        Dim r As String
        r = "A" & i
        Dim r1 As String
        r1 = "B" & i
        Sheets("LoanOfficers").Select
        Range(r).Select
        Selection.Copy
        Sheets("Active").Select
        ActiveSheet.Range("$A$1:$AD$29542").AutoFilter Field:=28, Criteria1:=Range("LoanOfficers!" & r).Value, Operator:=xlAnd
        Range("A1:AD29542").Select
        Range("P1").Activate
        Application.CutCopyMode = False
        Selection.Copy
        Sheets.Add After:=ActiveSheet
        Dim newSheet As String
        newSheet = ActiveSheet.Name
        Range("A3").Select
        ActiveSheet.Paste
        Range("A1").Select
        Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "Loan Officer NMLS"
        Range("B1").Select
        ActiveCell.FormulaR1C1 = "=R[3]C[27]"
        Range("C1").Select
        ActiveCell.FormulaR1C1 = "=LoanOfficers!" & r1
        Range("C1").Select
        Sheets("LoanOfficers").Select
        Range("B2").Select
        Sheets(newSheet).Select
        Range("C1").Select
        ActiveCell.FormulaR1C1 = Range("LoanOfficers!" & r1).Value
        Cells.Select
        Cells.EntireColumn.AutoFit
        Range("D2").Select
        Range("T1").Select
        ActiveCell.FormulaR1C1 = "Plug In Your Interest Rate"
        Range("T1").Select
        Selection.Style = "Good"
        Range("U1").Select
        Selection.NumberFormat = "0.00%"
        ActiveCell.FormulaR1C1 = "0.04"
        Range("U4").Select
        ActiveCell.FormulaR1C1 = "=SUM(PMT(R1C21/12,RC[-12],RC[-5]+3000)*-1)-RC[-4]"
        Range("U4").Select
        If Range("U" & Rows.Count).End(xlUp).Row = 4 Then
            Selection.AutoFill Destination:=Range("U4:U5")
        Else
            Selection.AutoFill Destination:=Range("U4:U" & Range("U" & Rows.Count).End(xlUp).Row)
        End If
        
        Range("U4:U19").Select
        Cells.Select
        Range("G1").Activate
        Cells.EntireColumn.AutoFit
        Cells.Select
        Range("G1").Activate
        Sheets(newSheet).Select
        Range("C1").Select
        Selection.Copy
        Sheets(newSheet).Select
        Sheets(newSheet).Name = Range("C1").Value
        Range("B24").Select
    Next i
    
End Sub
Sub SplitIntoWorkBooks()
'
' SplitIntoWorkBooks Macro
'
    Dim xPath As String
    xPath = Application.ActiveWorkbook.Path
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each xWs In ThisWorkbook.Sheets
        xWs.Copy
        Application.ActiveWorkbook.SaveAs Filename:=xPath & "\" & xWs.Name & ".xls"
        Application.ActiveWorkbook.Close False
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
'
End Sub
