Attribute VB_Name = "HideInvoiceCells"

Function HideInvoiceCellsBroadway()
    Dim broadway As Worksheet
    Dim install As String
    Dim expense As String
    Dim sales As String
    Dim a As Long
    Dim B As Long
    Dim c As Long
    
    Set broadway = ThisWorkbook.Sheets("Tin Roof Broadway")
    
install = "P58"
expense = "P177"
sales = "P208"
 Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

If broadway.Range(install).Value = False Then
    broadway.Rows("58:175").EntireRow.Hidden = True
    Else
    For a = 67 To 164
        If broadway.Range("P" & a).Value = False Then
            broadway.Rows(a).EntireRow.Hidden = True
        ElseIf Range("P" & a).Value = "" Then
            broadway.Rows(a).EntireRow.Hidden = False
        End If
    Next a
    If broadway.Range("P168").Value = False Then
        broadway.Rows("168").EntireRow.Hidden = True
    Else: broadway.Range("P168").EntireRow.Hidden = False
    End If
    If broadway.Range("P166").Value = False Then
        broadway.Rows("166").EntireRow.Hidden = True
    Else: broadway.Range("P166").EntireRow.Hidden = False
    End If
    If broadway.Range("P169").Value = False Then
        broadway.Rows("169").EntireRow.Hidden = True
    Else: broadway.Range("P169").EntireRow.Hidden = False
    End If
End If
 
If broadway.Range(expense).Value = "FALSE" Then
    broadway.Rows("177:1206").EntireRow.Hidden = True
    Else
    For B = 185 To 199
        If broadway.Range("P" & B).Value = "FALSE" Then
            broadway.Rows(B).EntireRow.Hidden = True
        ElseIf Range("P" & B).Value = "" Then
            broadway.Rows(B).EntireRow.Hidden = False
        End If
    Next B
End If
 
If broadway.Range(sales).Value = "FALSE" Then
    broadway.Rows("208:237").EntireRow.Hidden = True
    Else
    For c = 216 To 230
        If broadway.Range("P" & c).Value = "FALSE" Then
            broadway.Rows(c).EntireRow.Hidden = True
        ElseIf Range("P" & c).Value = "" Then
            broadway.Rows(c).EntireRow.Hidden = False
        End If
    Next c
End If
 
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Function


Function HideInvoiceCells(ws As Worksheet)
 Dim install As String
 Dim expense As String
 Dim sales As String
 Dim a As Long
 Dim B As Long
 Dim c As Long
 
install = "M60"
expense = "M177"
sales = "M207"
 Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

If ws.Range(install).Value = False Then
    ws.Rows("60:175").EntireRow.Hidden = True
Else
    For a = 67 To 164
        If ws.Range("M" & a).Value = False Then
            ws.Rows(a).EntireRow.Hidden = True
        ElseIf Range("M" & a).Value = "" Then
            ws.Rows(a).EntireRow.Hidden = False
        End If
    Next a
    If ws.Range("M166").Value = False Then
        ws.Rows("166").EntireRow.Hidden = True
    Else: ws.Rows("166").EntireRow.Hidden = False
    End If
    If ws.Range("M168").Value = False Then
        ws.Rows("168").EntireRow.Hidden = True
    Else: ws.Rows("168").EntireRow.Hidden = False
    End If
    If ws.Range("M169").Value = False Then
        ws.Rows("169").EntireRow.Hidden = True
    Else: ws.Rows("169").EntireRow.Hidden = False
    End If
End If

If ws.Range(expense).Value = "FALSE" Then
    ws.Rows("177:205").EntireRow.Hidden = True
Else
    For B = 184 To 198
        If ws.Range("M" & B).Value = "FALSE" Then
            ws.Rows(B).EntireRow.Hidden = True
        ElseIf ws.Range("M" & B).Value = "" Then
            ws.Rows(B).EntireRow.Hidden = False
        End If
    Next B
End If

If ws.Range(sales).Value = "FALSE" Then
    ws.Rows("207:235").EntireRow.Hidden = True
Else
    For c = 214 To 228
        If ws.Range("M" & c).Value = "FALSE" Then
            ws.Rows(c).EntireRow.Hidden = True
        ElseIf Range("M" & c).Value = "" Then
            ws.Rows(c).EntireRow.Hidden = False
        End If
    Next c
End If

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Function


Sub HideAllInvoiceCells()
    Dim Kings As Worksheet
    Dim Misc As Worksheet
    Dim TRDem As Worksheet
    Dim TRMem As Worksheet
    Dim TRBham As Worksheet
    
    Set Kings = ThisWorkbook.Sheets("Kings")
    Set Misc = ThisWorkbook.Sheets("Misc")
    Set TRDem = ThisWorkbook.Sheets("Tin Roof Demonbreun")
    Set TRMem = ThisWorkbook.Sheets("TR Memphis")
    Set TRBham = ThisWorkbook.Sheets("TR Birmingham")
    
    HideInvoiceCellsBroadway
    HideInvoiceCells ws:=Kings
    HideInvoiceCells ws:=Misc
    HideInvoiceCells ws:=TRDem
    HideInvoiceCells ws:=TRMem
    HideInvoiceCells ws:=TRBham
End Sub


