Attribute VB_Name = "PasteToYTDFunctions"
Function NashvillePasteTotalsToYTD()
    'for detecting range length
    Dim RangeYTD As Range
    Dim LastCellYTD As String
    Dim Range As Range
    Dim LastCell As String
    Dim StartCell As String
    Dim RangeLength As Long
    Dim RangeLengthYTD As Long
    'For easy reference to books and sheets
    Dim TotalsSheet As Worksheet
    Dim PaidOutSheet As Worksheet
    Dim BilledSheet As Worksheet
    'For copy from current book
    Dim TRBwayBilled As Range
    Dim TRBwayPaidOut As Range
    Dim TRDBIlled As Range
    Dim TRDPaidOut As Range
    Dim MiscBilled As Range
    Dim MiscPaidOut As Range
    Dim KingsBilled As Range
    Dim KingsPaidOut As Range
    'For Paste to YTD
    Dim TRBwayPaidOutStart As Range
    Dim TRBwayBilledStart As Range
    Dim TRDPaidOutStart As Range
    Dim TRDBilledStart As Range
    Dim MiscPaidOutStart As Range
    Dim MiscBilledStart As Range
    Dim KingsPaidOutStart As Range
    Dim KingsBilledStart As Range

        
    'Find Length of range current sheet
    Set TotalsSheet = ThisWorkbook.Sheets("Totals")
    With TotalsSheet.Range("B1:BZ1")
        Set Range = .Find(What:="Company Expenses")
        If Not Range Is Nothing Then
            LastCell = Range.Address
            StartCell = "$B$1"
        End If
    End With
    Set Range = TotalsSheet.Range(StartCell & ":" & LastCell)
    RangeLength = Range.Count
    
    'Find Length of range in YTD Sheet
    Set PaidOutSheet = YTDBook.Worksheets("Yearly Paid Out Nash (1099s)")
    With PaidOutSheet.Range("B1:BZ1")
        Set RangeYTD = .Find(What:="Company Expenses")
        If Not Range Is Nothing Then
            LastCellYTD = RangeYTD.Address
        End If
    End With
    Set RangeYTD = PaidOutSheet.Range(StartCell & ":" & LastCellYTD)
    RangeLengthYTD = RangeYTD.Count
    
    'Compare lengths
    If RangeLength <> RangeLengthYTD Then
        MsgBox ("Name Missing in Nashville YTD Workbook")
        Exit Function
    End If
    
    'Copy and paste values only
    Set BilledSheet = YTDBook.Sheets("Yearly Billed Nash (My 1099)")
    
    'Set Ranges to be copied
    Set TRBwayBilled = Range.Offset(1)
    Set TRBwayPaidOut = Range.Offset(2)
    Set TRDBIlled = Range.Offset(4)
    Set TRDPaidOut = Range.Offset(5)
    Set MiscBilled = Range.Offset(7)
    Set MiscPaidOut = Range.Offset(8)
    Set KingsBilled = Range.Offset(10)
    Set KingsPaidOut = Range.Offset(11)
    
    'Set ranges for pasting
    Set TRBwayPaidOutStart = PaidOutSheet.Range("B3")
    Set TRBwayBilledStart = BilledSheet.Range("B3")
    Set TRDPaidOutStart = PaidOutSheet.Range("B33")
    Set TRDBilledStart = BilledSheet.Range("B32")
    Set MiscPaidOutStart = PaidOutSheet.Range("B94")
    Set MiscBilledStart = BilledSheet.Range("B90")
    Set KingsPaidOutStart = PaidOutSheet.Range("B63")
    Set KingsBilledStart = BilledSheet.Range("B61")
    
    PasteInvoiceNumber StartRow:=TRBwayPaidOutStart
    
    PasteValues Billed:=TRBwayBilled, _
        PaidOut:=TRBwayPaidOut, _
        BilledStart:=TRBwayBilledStart, _
        PaidOutStart:=TRBwayPaidOutStart
        
    PasteValues Billed:=TRDBIlled, _
        PaidOut:=TRDPaidOut, _
        BilledStart:=TRDBilledStart, _
        PaidOutStart:=TRDPaidOutStart
        
    PasteValues Billed:=MiscBilled, _
        PaidOut:=MiscPaidOut, _
        BilledStart:=MiscBilledStart, _
        PaidOutStart:=MiscPaidOutStart
        
    PasteValues Billed:=KingsBilled, _
        PaidOut:=KingsPaidOut, _
        BilledStart:=KingsBilledStart, _
        PaidOutStart:=KingsPaidOutStart
    
End Function

Function PasteInvoiceNumber(StartRow As Range)
    Dim InvoiceNumber As String
    InvoiceNumber = ThisWorkbook.Sheets("Kings").Range("K7").Value
    YTDBook.Activate
    Select Case PayPeriodNumber
        Case 1
            StartRow.Offset(0, -1) = InvoiceNumber
        Case 2
            StartRow.Offset(1, -1) = InvoiceNumber
        Case 3
            StartRow.Offset(2, -1) = InvoiceNumber
        Case 4
            StartRow.Offset(3, -1) = InvoiceNumber
        Case 5
            StartRow.Offset(4, -1) = InvoiceNumber
        Case 6
            StartRow.Offset(5, -1) = InvoiceNumber
        Case 7
            StartRow.Offset(6, -1) = InvoiceNumber
        Case 8
            StartRow.Offset(7, -1) = InvoiceNumber
        Case 9
            StartRow.Offset(8, -1) = InvoiceNumber
        Case 10
            StartRow.Offset(9, -1) = InvoiceNumber
        Case 11
            StartRow.Offset(10, -1) = InvoiceNumber
        Case 12
            StartRow.Offset(11, -1) = InvoiceNumber
        Case 13
            StartRow.Offset(12, -1) = InvoiceNumber
        Case 14
            StartRow.Offset(13, -1) = InvoiceNumber
        Case 15
            StartRow.Offset(14, -1) = InvoiceNumber
        Case 16
            StartRow.Offset(15, -1) = InvoiceNumber
        Case 17
            StartRow.Offset(16, -1) = InvoiceNumber
        Case 18
            StartRow.Offset(17, -1) = InvoiceNumber
        Case 19
            StartRow.Offset(18, -1) = InvoiceNumber
        Case 20
            StartRow.Offset(19, -1) = InvoiceNumber
        Case 21
            StartRow.Offset(20, -1) = InvoiceNumber
        Case 22
            StartRow.Offset(21, -1) = InvoiceNumber
        Case 23
            StartRow.Offset(22, -1) = InvoiceNumber
        Case 24
            StartRow.Offset(23, -1) = InvoiceNumber
        Case 25
            StartRow.Offset(24, -1) = InvoiceNumber
        Case 26
            StartRow.Offset(25, -1) = InvoiceNumber
    End Select
End Function

Function PasteValues(Billed As Range, _
    PaidOut As Range, _
    BilledStart As Range, _
    PaidOutStart As Range)
    
    ThisWorkbook.Activate
    PaidOut.Copy
    YTDBook.Activate
    Select Case PayPeriodNumber
        Case 1
            PaidOutStart.PasteSpecial (xlPasteValues)
        Case 2
            PaidOutStart.Offset(1).PasteSpecial (xlPasteValues)
        Case 3
            PaidOutStart.Offset(2).PasteSpecial (xlPasteValues)
        Case 4
            PaidOutStart.Offset(3).PasteSpecial (xlPasteValues)
        Case 5
            PaidOutStart.Offset(4).PasteSpecial (xlPasteValues)
        Case 6
            PaidOutStart.Offset(5).PasteSpecial (xlPasteValues)
        Case 7
            PaidOutStart.Offset(6).PasteSpecial (xlPasteValues)
        Case 8
            PaidOutStart.Offset(7).PasteSpecial (xlPasteValues)
        Case 9
            PaidOutStart.Offset(8).PasteSpecial (xlPasteValues)
        Case 10
            PaidOutStart.Offset(9).PasteSpecial (xlPasteValues)
        Case 11
            PaidOutStart.Offset(10).PasteSpecial (xlPasteValues)
        Case 12
            PaidOutStart.Offset(11).PasteSpecial (xlPasteValues)
        Case 13
            PaidOutStart.Offset(12).PasteSpecial (xlPasteValues)
        Case 14
            PaidOutStart.Offset(13).PasteSpecial (xlPasteValues)
        Case 15
            PaidOutStart.Offset(14).PasteSpecial (xlPasteValues)
        Case 16
            PaidOutStart.Offset(15).PasteSpecial (xlPasteValues)
        Case 17
            PaidOutStart.Offset(16).PasteSpecial (xlPasteValues)
        Case 18
            PaidOutStart.Offset(17).PasteSpecial (xlPasteValues)
        Case 19
            PaidOutStart.Offset(18).PasteSpecial (xlPasteValues)
        Case 20
            PaidOutStart.Offset(19).PasteSpecial (xlPasteValues)
        Case 21
            PaidOutStart.Offset(20).PasteSpecial (xlPasteValues)
        Case 22
            PaidOutStart.Offset(21).PasteSpecial (xlPasteValues)
        Case 23
            PaidOutStart.Offset(22).PasteSpecial (xlPasteValues)
        Case 24
            PaidOutStart.Offset(23).PasteSpecial (xlPasteValues)
        Case 25
            PaidOutStart.Offset(24).PasteSpecial (xlPasteValues)
        Case 26
            PaidOutStart.Offset(25).PasteSpecial (xlPasteValues)
    End Select
    
    ThisWorkbook.Activate
    Billed.Copy
    YTDBook.Activate
    Select Case PayPeriodNumber
        Case 1
            BilledStart.PasteSpecial (xlPasteValues)
        Case 2
            BilledStart.Offset(1).PasteSpecial (xlPasteValues)
        Case 3
            BilledStart.Offset(2).PasteSpecial (xlPasteValues)
        Case 4
            BilledStart.Offset(3).PasteSpecial (xlPasteValues)
        Case 5
            BilledStart.Offset(4).PasteSpecial (xlPasteValues)
        Case 6
            BilledStart.Offset(5).PasteSpecial (xlPasteValues)
        Case 7
            BilledStart.Offset(6).PasteSpecial (xlPasteValues)
        Case 8
            BilledStart.Offset(7).PasteSpecial (xlPasteValues)
        Case 9
            BilledStart.Offset(8).PasteSpecial (xlPasteValues)
        Case 10
            BilledStart.Offset(9).PasteSpecial (xlPasteValues)
        Case 11
            BilledStart.Offset(10).PasteSpecial (xlPasteValues)
        Case 12
            BilledStart.Offset(11).PasteSpecial (xlPasteValues)
        Case 13
            BilledStart.Offset(12).PasteSpecial (xlPasteValues)
        Case 14
            BilledStart.Offset(13).PasteSpecial (xlPasteValues)
        Case 15
            BilledStart.Offset(14).PasteSpecial (xlPasteValues)
        Case 16
            BilledStart.Offset(15).PasteSpecial (xlPasteValues)
        Case 17
            BilledStart.Offset(16).PasteSpecial (xlPasteValues)
        Case 18
            BilledStart.Offset(17).PasteSpecial (xlPasteValues)
        Case 19
            BilledStart.Offset(18).PasteSpecial (xlPasteValues)
        Case 20
            BilledStart.Offset(19).PasteSpecial (xlPasteValues)
        Case 21
            BilledStart.Offset(20).PasteSpecial (xlPasteValues)
        Case 22
            BilledStart.Offset(21).PasteSpecial (xlPasteValues)
        Case 23
            BilledStart.Offset(22).PasteSpecial (xlPasteValues)
        Case 24
            BilledStart.Offset(23).PasteSpecial (xlPasteValues)
        Case 25
            BilledStart.Offset(24).PasteSpecial (xlPasteValues)
        Case 26
            BilledStart.Offset(25).PasteSpecial (xlPasteValues)
    End Select
    
End Function

Function PasteTotalsToYTD(SourceRange As Range, _
    SourceStartCell As String, _
    DestRange As Range, _
    DestWorksheet As Worksheet, _
    PaidOutStart As Range, _
    BilledStart As Range, _
    Name As String)
    
    'for detecting range length
    Dim RangeYTD As Range
    Dim LastCellYTD As String
    Dim Range As Range
    Dim LastCell As String
    Dim RangeLength As Long
    Dim RangeLengthYTD As Long
    'For copy from current book
    Dim Billed As Range
    Dim PaidOut As Range
    Dim InvoiceNumber As String
    
    'Find Length of range current sheet
    With SourceRange
        Set Range = .Find(What:="Company Expenses")
        If Not Range Is Nothing Then
            LastCell = Range.Address
        End If
    End With
    Set Range = TotalsSheet.Range(SourceStartCell & ":" & LastCell)
    RangeLength = Range.Count
    
    'Find Length of range in YTD Sheet
    
    With DestRange
        Set RangeYTD = .Find(What:="Company Expenses")
        If Not Range Is Nothing Then
            LastCellYTD = RangeYTD.Address
        End If
    End With
    Set RangeYTD = DestWorksheet.Range("C1" & ":" & LastCellYTD)
    RangeLengthYTD = RangeYTD.Count
    
    'Compare lengths
    If RangeLength <> RangeLengthYTD Then
        MsgBox ("Name Missing in " & Name & " YTD Worksheet")
        Exit Function
    End If
    
    'Copy and paste values only
    'Set Ranges to be copied
    Set Billed = Range.Offset(1)
    Set PaidOut = Range.Offset(2)

    PasteValues Billed:=Billed, _
        PaidOut:=PaidOut, _
        BilledStart:=BilledStart, _
        PaidOutStart:=PaidOutStart

End Function

Function PasteNonAttrExpenseValues()
    Dim Expenses As Range
    Dim Dest As Range
     
    'Set Range to be copied
    Set Expenses = ThisWorkbook.Sheets("Totals").Range("H32")
    
    'Set range for pasting
    Set Dest = YTDBook.Worksheets("Profit Margins").Range("P4")

    
    'Copy and Paste Expenses
    ThisWorkbook.Activate
    Expenses.Copy
    YTDBook.Activate
    Select Case PayPeriodNumber
        Case 1
            Dest.PasteSpecial (xlPasteValues)
        Case 2
            Dest.Offset(1).PasteSpecial (xlPasteValues)
        Case 3
            Dest.Offset(2).PasteSpecial (xlPasteValues)
        Case 4
            Dest.Offset(3).PasteSpecial (xlPasteValues)
        Case 5
            Dest.Offset(4).PasteSpecial (xlPasteValues)
        Case 6
            Dest.Offset(5).PasteSpecial (xlPasteValues)
        Case 7
            Dest.Offset(6).PasteSpecial (xlPasteValues)
        Case 8
            Dest.Offset(9).PasteSpecial (xlPasteValues)
        Case 9
            Dest.Offset(10).PasteSpecial (xlPasteValues)
        Case 10
            Dest.Offset(11).PasteSpecial (xlPasteValues)
        Case 11
            Dest.Offset(12).PasteSpecial (xlPasteValues)
        Case 12
            Dest.Offset(15).PasteSpecial (xlPasteValues)
        Case 13
            Dest.Offset(16).PasteSpecial (xlPasteValues)
        Case 14
            Dest.Offset(17).PasteSpecial (xlPasteValues)
        Case 15
            Dest.Offset(18).PasteSpecial (xlPasteValues)
        Case 16
            Dest.Offset(19).PasteSpecial (xlPasteValues)
        Case 17
            Dest.Offset(20).PasteSpecial (xlPasteValues)
        Case 18
            Dest.Offset(21).PasteSpecial (xlPasteValues)
        Case 19
            Dest.Offset(24).PasteSpecial (xlPasteValues)
        Case 20
            Dest.Offset(25).PasteSpecial (xlPasteValues)
        Case 21
            Dest.Offset(26).PasteSpecial (xlPasteValues)
        Case 22
            Dest.Offset(27).PasteSpecial (xlPasteValues)
        Case 23
            Dest.Offset(28).PasteSpecial (xlPasteValues)
        Case 24
            Dest.Offset(29).PasteSpecial (xlPasteValues)
        Case 25
            Dest.Offset(30).PasteSpecial (xlPasteValues)
        Case 26
            Dest.Offset(31).PasteSpecial (xlPasteValues)
    End Select
    
End Function



