Attribute VB_Name = "UtilSubs"
Sub HidePayStubCells()
Attribute HidePayStubCells.VB_Description = "hide unused paystub rows"
Attribute HidePayStubCells.VB_ProcData.VB_Invoke_Func = "u\n14"
 Dim i As Long
 Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
For i = 40 To 900
   If Range("P" & i).Value = "FALSE" Then
       Rows(i).EntireRow.Hidden = True
   ElseIf Range("P" & i).Value = "" Then
       Rows(i).EntireRow.Hidden = False
   End If
 Next i
 
 

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub Unhide_All_Sheets()
    Dim wks As Worksheet
 
    For Each wks In ActiveWorkbook.Worksheets
        wks.Visible = xlSheetVisible
    Next wks
End Sub

Sub StartNewWorkbook()
'
' StartNewWorkbook Macro
'

'
    ThisWorkbook.Sheets("Pay Period Dates").Range("R2").Select
    Selection.Copy
    ThisWorkbook.Sheets("Kings").Select
    Range("A10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ThisWorkbook.Sheets("Pay Period Dates").Select
    Range("Q2").Select
    Selection.Copy
    Range("S2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub
