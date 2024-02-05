Attribute VB_Name = "Module6"
Public YTDBook As Workbook
Public PayPeriodNumber As Integer
Public YTDName As String
Public TotalsSheet As Worksheet
Dim FilePath As String
Dim BhamSourceRange As Range
Dim BhamDestRange As Range
Dim BhamStart As String
Dim BhamDestSheet As Worksheet
Dim BhamPaidOutStart As Range
Dim BhamBilledStart As Range
Dim MemphisSourceRange As Range
Dim MemphisDestRange As Range
Dim MemphisStart As String
Dim MemphisDestSheet As Worksheet
Dim MemphisPaidOutStart As Range
Dim MemphisBilledStart As Range


Function IsWorkBookOpen(Name As String) As Boolean
    Dim YTDwb As Workbook
    On Error Resume Next
    Set YTDwb = Application.Workbooks.Item(Name)
    IsWorkBookOpen = (Not YTDwb Is Nothing)
End Function

Public Sub YTDButton()
   'Check if YTD is open and open if it isn't
   FilePath = "https://d.docs.live.net/91c94bc5a0fe7d16/Documents/Sound%20Roots%20YTD/"
   YTDName = ThisWorkbook.Sheets("AutomationData").Range("B1").Value
   Dim FullPath As String
   FullPath = FilePath & YTDName
    
    Dim xRet As Boolean
    xRet = IsWorkBookOpen(FullPath)
        If Not xRet Then
            Workbooks.Open (FullPath)
        End If
        
    Set YTDBook = Application.Workbooks(YTDName)
    PayPeriodNumber = ThisWorkbook.Sheets("Pay Period Dates").Range("S2").Value
    Set TotalsSheet = ThisWorkbook.Sheets("Totals")
    
    'BIRMINGHAM
    Set BhamSourceRange = TotalsSheet.Range("B18:Z18")
    Set BhamDestSheet = YTDBook.Worksheets("Tin Roof Birmingham")
    Set BhamDestRange = BhamDestSheet.Range("B1:Z1")
    'Ranges in YTD for pasting
    Set BhamPaidOutStart = BhamDestSheet.Range("C31")
    Set BhamBilledStart = BhamDestSheet.Range("C2")
    'First cell in Totals Sheet
    BhamStart = "$B$18"
    
    PasteTotalsToYTD SourceRange:=BhamSourceRange, _
        SourceStartCell:=BhamStart, _
        DestRange:=BhamDestRange, _
        DestWorksheet:=BhamDestSheet, _
        PaidOutStart:=BhamPaidOutStart, _
        BilledStart:=BhamBilledStart, _
        Name:="Birmingham"
        
    'MEMPHIS
    Set MemphisSourceRange = TotalsSheet.Range("B14:Z14")
    Set MemphisDestSheet = YTDBook.Worksheets("Tin Roof Memphis")
    Set MemphisDestRange = MemphisDestSheet.Range("B1:Z1")
    'Ranges in YTD for pasting
    Set MemphisPaidOutStart = MemphisDestSheet.Range("C31")
    Set MemphisBilledStart = MemphisDestSheet.Range("C2")
    'First cell in Totals Sheet
    MemphisStart = "$B$14"
    
     PasteTotalsToYTD SourceRange:=MemphisSourceRange, _
        SourceStartCell:=MemphisStart, _
        DestRange:=MemphisDestRange, _
        DestWorksheet:=MemphisDestSheet, _
        PaidOutStart:=MemphisPaidOutStart, _
        BilledStart:=MemphisBilledStart, _
        Name:="Memphis"
        
    
    NashvillePasteTotalsToYTD
    PasteNonAttrExpenseValues
    
    ThisWorkbook.Sheets("Totals").Activate
    
End Sub
