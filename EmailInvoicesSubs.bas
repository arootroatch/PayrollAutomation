Attribute VB_Name = "EmailInvoicesSubs"
Dim SendKings As Boolean
Dim SendMisc As Boolean
Dim SendTRBro As Boolean
Dim SendTRDem As Boolean
Dim SendTRMem As Boolean
Dim SendTRBham As Boolean
Dim KingsInvoice As Range
Dim MiscInvoice As Range
Dim TRBroInvoice As Range
Dim TRDemInvoice As Range
Dim TRMemInvoice As Range
Dim TRBhamInvoice As Range
Dim KingsFileName As String
Dim MiscFIleName As String
Dim TRBroFileName As String
Dim TRDemFileName As String
Dim TRMemFileName As String
Dim TRBhamFileName As String
Dim KingsFolderstring As String
Dim MiscFolderstring As String
Dim TRBroFolderstring As String
Dim TRDemFolderstring As String
Dim TRMemFolderstring As String
Dim TRBhamFolderstring As String
Dim KingsBody As String
Dim MiscBody As String
Dim TRBroBody As String
Dim TRDemBody As String
Dim TRMemBody As String
Dim TRBhamBody As String
Dim KingsTo As String
Dim MiscTo As String
Dim TRBroTo As String
Dim TRDemTo As String
Dim TRMemTo As String
Dim TRBhamTo As String
Dim KingsCC As String
Dim MiscCC As String
Dim TRBroCC As String
Dim TRDemCC As String
Dim TRMemCC As String
Dim TRBhamCC As String
Dim KingsReceipt1 As String
Dim MiscReceipt1 As String
Dim TRBroReceipt1 As String
Dim TRDemReceipt1 As String
Dim TRMemReceipt1 As String
Dim TRBhamReceipt1 As String
Dim KingsReceipt2 As String
Dim MiscReceipt2 As String
Dim TRBroReceipt2 As String
Dim TRDemReceipt2 As String
Dim TRMemReceipt2 As String
Dim TRBhamReceipt2 As String
Dim KingsReceipt3 As String
Dim MiscReceipt3 As String
Dim TRBroReceipt3 As String
Dim TRDemReceipt3 As String
Dim TRMemReceipt3 As String
Dim TRBhamReceipt3 As String
Dim KingsReceipt4 As String
Dim MiscReceipt4 As String
Dim TRBroReceipt4 As String
Dim TRDemReceipt4 As String
Dim TRMemReceipt4 As String
Dim TRBhamReceipt4 As String
Dim KingsReceipt5 As String
Dim MiscReceipt5 As String
Dim TRBroReceipt5 As String
Dim TRDemReceipt5 As String
Dim TRMemReceipt5 As String
Dim TRBhamReceipt5 As String
Dim KingsReceiptFolder As String
Dim MiscReceiptFolder As String
Dim TRBroReceiptFolder As String
Dim TRDemReceiptFolder As String
Dim TRMemReceiptFolder As String
Dim TRBhamReceiptFolder As String
Dim EmailSubject As String


Public Sub EmailKingsInvoice()
    Set KingsInvoice = ThisWorkbook.Sheets("Kings").Range("A1:L235")
    KingsFileName = ThisWorkbook.Sheets("AutomationData").Range("C4").Value
    KingsFolderstring = ThisWorkbook.Sheets("AutomationData").Range("B4").Value
    KingsBody = ThisWorkbook.Sheets("AutomationData").Range("G4").Value
    KingsTo = ThisWorkbook.Sheets("AutomationData").Range("D4").Value
    KingsCC = ThisWorkbook.Sheets("AutomationData").Range("E4").Value
    
    EmailSubject = ThisWorkbook.Sheets("AutomationData").Range("F4").Value
    
    KingsReceipt1 = ThisWorkbook.Sheets("AutomationData").Range("I4").Value
    KingsReceipt2 = ThisWorkbook.Sheets("AutomationData").Range("J4").Value
    KingsReceipt3 = ThisWorkbook.Sheets("AutomationData").Range("K4").Value
    KingsReceipt4 = ThisWorkbook.Sheets("AutomationData").Range("L4").Value
    KingsReceipt5 = ThisWorkbook.Sheets("AutomationData").Range("M4").Value
    KingsReceiptFolder = ThisWorkbook.Sheets("AutomationData").Range("N4").Value
    
    SendKings = ThisWorkbook.Sheets("AutomationData").Range("H4")
    
    If SendKings = True Then
        SaveMailInvoiceAsPDF Invoice:=KingsInvoice, _
        FileName:=KingsFileName, _
        Folderstring:=KingsFolderstring, _
        strbody:=KingsBody, _
        RecipientEmail:=KingsTo, _
        CCRecipients:=KingsCC, _
        EmailSubject:=EmailSubject, _
        ReceiptFolder:=KingsReceiptFolder, _
        Receipt1:=KingsReceipt1, _
        Receipt2:=KingsReceipt2, _
        Receipt3:=KingsReceipt3, _
        Receipt4:=KingsReceipt4, _
        Receipt5:=KingsReceipt5
    End If
    
End Sub

Public Sub EmailMiscInvoice()
    Set MiscInvoice = ThisWorkbook.Sheets("Misc").Range("A1:L235")
    MiscFIleName = ThisWorkbook.Sheets("AutomationData").Range("C5").Value
    MiscFolderstring = ThisWorkbook.Sheets("AutomationData").Range("B5").Value
    MiscBody = ThisWorkbook.Sheets("AutomationData").Range("G5").Value
    MiscTo = ThisWorkbook.Sheets("AutomationData").Range("D5").Value
    MiscCC = ThisWorkbook.Sheets("AutomationData").Range("E5").Value
    EmailSubject = ThisWorkbook.Sheets("AutomationData").Range("F5").Value
    
    
    MiscReceipt1 = ThisWorkbook.Sheets("AutomationData").Range("I5").Value
    MiscReceipt2 = ThisWorkbook.Sheets("AutomationData").Range("J5").Value
    MiscReceipt3 = ThisWorkbook.Sheets("AutomationData").Range("K5").Value
    MiscReceipt4 = ThisWorkbook.Sheets("AutomationData").Range("L5").Value
    MiscReceipt5 = ThisWorkbook.Sheets("AutomationData").Range("M5").Value
    MiscReceiptFolder = ThisWorkbook.Sheets("AutomationData").Range("N5").Value
    
    SendMisc = ThisWorkbook.Sheets("AutomationData").Range("H5")
    

    If SendMisc = True Then
        SaveMailInvoiceAsPDF Invoice:=MiscInvoice, _
        FileName:=MiscFIleName, _
        Folderstring:=MiscFolderstring, _
        strbody:=MiscBody, _
        RecipientEmail:=MiscTo, _
        CCRecipients:=MiscCC, _
        EmailSubject:=EmailSubject, _
        ReceiptFolder:=MiscReceiptFolder, _
        Receipt1:=MiscReceipt1, _
        Receipt2:=MiscReceipt2, _
        Receipt3:=MiscReceipt3, _
        Receipt4:=MiscReceipt4, _
        Receipt5:=MiscReceipt5
    End If
    
End Sub

Public Sub EmailTRBroInvoice()
    Set TRBroInvoice = ThisWorkbook.Sheets("Tin Roof Broadway").Range("A1:O237")
    TRBroFileName = ThisWorkbook.Sheets("AutomationData").Range("C6").Value
    TRBroFolderstring = ThisWorkbook.Sheets("AutomationData").Range("B6").Value
    TRBroBody = ThisWorkbook.Sheets("AutomationData").Range("G6").Value
    TRBroTo = ThisWorkbook.Sheets("AutomationData").Range("D6").Value
    TRBroCC = ThisWorkbook.Sheets("AutomationData").Range("E6").Value
    
    EmailSubject = ThisWorkbook.Sheets("AutomationData").Range("F6").Value
    
    TRBroReceipt1 = ThisWorkbook.Sheets("AutomationData").Range("I6").Value
    TRBroReceipt2 = ThisWorkbook.Sheets("AutomationData").Range("J6").Value
    TRBroReceipt3 = ThisWorkbook.Sheets("AutomationData").Range("K6").Value
    TRBroReceipt4 = ThisWorkbook.Sheets("AutomationData").Range("L6").Value
    TRBroReceipt5 = ThisWorkbook.Sheets("AutomationData").Range("M6").Value

    TRBroReceiptFolder = ThisWorkbook.Sheets("AutomationData").Range("N6").Value
    
    SendTRBro = ThisWorkbook.Sheets("AutomationData").Range("H6")
    
    If SendTRBro = True Then
        SaveMailInvoiceAsPDF Invoice:=TRBroInvoice, _
        FileName:=TRBroFileName, _
        Folderstring:=TRBroFolderstring, _
        strbody:=TRBroBody, _
        RecipientEmail:=TRBroTo, _
        CCRecipients:=TRBroCC, _
        EmailSubject:=EmailSubject, _
        ReceiptFolder:=TRBroReceiptFolder, _
        Receipt1:=TRBroReceipt1, _
        Receipt2:=TRBroReceipt2, _
        Receipt3:=TRBroReceipt3, _
        Receipt4:=TRBroReceipt4, _
        Receipt5:=TRBroReceipt5
    End If
    
End Sub

Public Sub EmailTRDemInvoice()
    Set TRDemInvoice = ThisWorkbook.Sheets("Tin Roof Demonbreun").Range("A1:L235")
    TRDemFileName = ThisWorkbook.Sheets("AutomationData").Range("C7").Value
    TRDemFolderstring = ThisWorkbook.Sheets("AutomationData").Range("B7").Value
    TRDemBody = ThisWorkbook.Sheets("AutomationData").Range("G7").Value
    TRDemTo = ThisWorkbook.Sheets("AutomationData").Range("D7").Value
    TRDemCC = ThisWorkbook.Sheets("AutomationData").Range("E7").Value
    EmailSubject = ThisWorkbook.Sheets("AutomationData").Range("F7").Value
    
    TRDemReceipt1 = ThisWorkbook.Sheets("AutomationData").Range("I7").Value
    TRDemReceipt2 = ThisWorkbook.Sheets("AutomationData").Range("J7").Value
    TRDemReceipt3 = ThisWorkbook.Sheets("AutomationData").Range("K7").Value
    TRDemReceipt4 = ThisWorkbook.Sheets("AutomationData").Range("L7").Value
    TRDemReceipt5 = ThisWorkbook.Sheets("AutomationData").Range("M7").Value

    TRDemReceiptFolder = ThisWorkbook.Sheets("AutomationData").Range("N7").Value
    
    SendTRDem = ThisWorkbook.Sheets("AutomationData").Range("H7")
    
    If SendTRDem = True Then
        SaveMailInvoiceAsPDF Invoice:=TRDemInvoice, _
        FileName:=TRDemFileName, _
        Folderstring:=TRDemFolderstring, _
        strbody:=TRDemBody, _
        RecipientEmail:=TRDemTo, _
        CCRecipients:=TRDemCC, _
        EmailSubject:=EmailSubject, _
        ReceiptFolder:=TRDemReceiptFolder, _
        Receipt1:=TRDemReceipt1, _
        Receipt2:=TRDemReceipt2, _
        Receipt3:=TRDemReceipt3, _
        Receipt4:=TRDemReceipt4, _
        Receipt5:=TRDemReceipt5
    End If
    
End Sub


Public Sub EmailTRMemInvoice()
    Set TRMemInvoice = ThisWorkbook.Sheets("TR Memphis").Range("A1:L235")
    TRMemFileName = ThisWorkbook.Sheets("AutomationData").Range("C8").Value
    TRMemFolderstring = ThisWorkbook.Sheets("AutomationData").Range("B8").Value
    TRMemBody = ThisWorkbook.Sheets("AutomationData").Range("G8").Value
    TRMemTo = ThisWorkbook.Sheets("AutomationData").Range("D8").Value
    TRMemCC = ThisWorkbook.Sheets("AutomationData").Range("E8").Value
    
    EmailSubject = ThisWorkbook.Sheets("AutomationData").Range("F8").Value
    
    TRMemReceipt1 = ThisWorkbook.Sheets("AutomationData").Range("I8").Value
    TRMemReceipt2 = ThisWorkbook.Sheets("AutomationData").Range("J8").Value
    TRMemReceipt3 = ThisWorkbook.Sheets("AutomationData").Range("K8").Value
    TRMemReceipt4 = ThisWorkbook.Sheets("AutomationData").Range("L8").Value
    TRMemReceipt5 = ThisWorkbook.Sheets("AutomationData").Range("M8").Value
    TRMemReceiptFolder = ThisWorkbook.Sheets("AutomationData").Range("N8").Value
    
    SendTRMem = ThisWorkbook.Sheets("AutomationData").Range("H8")
    
   
    If SendTRMem = True Then
        SaveMailInvoiceAsPDF Invoice:=TRMemInvoice, _
        FileName:=TRMemFileName, _
        Folderstring:=TRMemFolderstring, _
        strbody:=TRMemBody, _
        RecipientEmail:=TRMemTo, _
        CCRecipients:=TRMemCC, _
        EmailSubject:=EmailSubject, _
        ReceiptFolder:=TRMemReceiptFolder, _
        Receipt1:=TRMemReceipt1, _
        Receipt2:=TRMemReceipt2, _
        Receipt3:=TRMemReceipt3, _
        Receipt4:=TRMemReceipt4, _
        Receipt5:=TRMemReceipt5
    End If

    
End Sub

Public Sub EmailTRBhamInvoice()
    Set TRBhamInvoice = ThisWorkbook.Sheets("TR Birmingham").Range("A1:L235")
    TRBhamFileName = ThisWorkbook.Sheets("AutomationData").Range("C9").Value
    TRBhamFolderstring = ThisWorkbook.Sheets("AutomationData").Range("B9").Value
    TRBhamBody = ThisWorkbook.Sheets("AutomationData").Range("G9").Value
    TRBhamTo = ThisWorkbook.Sheets("AutomationData").Range("D9").Value
    TRBhamCC = ThisWorkbook.Sheets("AutomationData").Range("E9").Value
    
    EmailSubject = ThisWorkbook.Sheets("AutomationData").Range("F9").Value
    
    TRBhamReceipt1 = ThisWorkbook.Sheets("AutomationData").Range("I9").Value
    TRBhamReceipt2 = ThisWorkbook.Sheets("AutomationData").Range("J9").Value
    TRBhamReceipt3 = ThisWorkbook.Sheets("AutomationData").Range("K9").Value
    TRBhamReceipt4 = ThisWorkbook.Sheets("AutomationData").Range("L9").Value
    TRBhamReceipt5 = ThisWorkbook.Sheets("AutomationData").Range("M9").Value
    TRBhamReceiptFolder = ThisWorkbook.Sheets("AutomationData").Range("N9").Value
    
    SendTRBham = ThisWorkbook.Sheets("AutomationData").Range("H9")
    
   
    If SendTRBham = True Then
        SaveMailInvoiceAsPDF Invoice:=TRBhamInvoice, _
        FileName:=TRBhamFileName, _
        Folderstring:=TRBhamFolderstring, _
        strbody:=TRBhamBody, _
        RecipientEmail:=TRBhamTo, _
        CCRecipients:=TRBhamCC, _
        EmailSubject:=EmailSubject, _
        ReceiptFolder:=TRBhamReceiptFolder, _
        Receipt1:=TRBhamReceipt1, _
        Receipt2:=TRBhamReceipt2, _
        Receipt3:=TRBhamReceipt3, _
        Receipt4:=TRBhamReceipt4, _
        Receipt5:=TRBhamReceipt5
    End If

    
End Sub


Function SaveMailInvoiceAsPDF(Invoice As Range, _
    FileName As String, _
    Folderstring As String, _
    strbody As String, _
    RecipientEmail As String, _
    CCRecipients As String, _
    EmailSubject As String, _
    ReceiptFolder As String, _
    Receipt1 As String, _
    Receipt2 As String, _
    Receipt3 As String, _
    Receipt4 As String, _
    Receipt5 As String)
    
    'Ron de Bruin, 27-Dec-2022
    'Do not forget to add the custom functions into your workbook
    'More Information : https://macexcel.com/examples/mailpdf/macmailexamples/
    'You can test the macro without changing anything
    
    Dim FilePathName As String
    Dim Receipt1Path As String
    Dim Receipt2Path As String
    Dim Receipt3Path As String
    Dim Receipt4Path As String
    Dim Receipt5Path As String
    Dim AllReceipts As String
    Dim PathArray(5) As String
    Dim RequestAccess As Boolean
    

    'Check for AppleScriptTask script file that we must use to create the mail
    If CheckAppleScriptTaskExcelScriptFile(ScriptFileName:="RDBMacMail2.scpt") = False Then
        MsgBox "Sorry the RDBMacMail2.scpt is not in the correct location"
        Exit Function
    End If
    
    
    'And create the file path string. Do not change this 2 lines
    FilePathName = Folderstring & Application.PathSeparator & FileName & ".pdf"
    Receipt1Path = ReceiptFolder & Application.PathSeparator & Receipt1
    Receipt2Path = ReceiptFolder & Application.PathSeparator & Receipt2
    Receipt3Path = ReceiptFolder & Application.PathSeparator & Receipt3
    Receipt4Path = ReceiptFolder & Application.PathSeparator & Receipt4
    Receipt5Path = ReceiptFolder & Application.PathSeparator & Receipt5
    
    AllReceipts = ""
    If Receipt1 <> "" Then
        AllReceipts = Receipt1Path
        PathArray(0) = Receipt1Path
    End If
    
    If Receipt2 <> "" Then
        AllReceipts = Receipt1Path & "," & Receipt2Path
        PathArray(1) = Receipt2Path
    End If
    
    If Receipt3 <> "" Then
        AllReceipts = Receipt1Path & "," & Receipt2Path & "," & Receipt3Path
        PathArray(2) = Receipt3Path
    End If
    
    If Receipt4 <> "" Then
        AllReceipts = Receipt1Path & "," & Receipt2Path & "," & Receipt3Path & "," & Receipt4Path
        PathArray(3) = Receipt4Path
    End If
    
    If Receipt5 <> "" Then
        AllReceipts = Receipt1Path & "," & Receipt2Path & "," & Receipt3Path & "," & Receipt4Path & "," & Receipt5Path
        PathArray(4) = Receipt5Path
    End If
    
    
    RequestAccess = GrantAccessToMultipleFiles(PathArray)
    
    
    'Create the PDF, you not have to edit this code block
    Invoice.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
    FilePathName, Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, IgnorePrintAreas:=False

    'Function call to create the mail, see this page for more info about the arguments
    'https://macexcel.com/examples/mailpdf/macmailexamples/
    'Do not change the attachmentname argument in this PDF function call
    MacExcelWithMacMailPDFCatalinaAndUp subject:=EmailSubject, _
    mailbody:=strbody, _
    toaddress:=RecipientEmail, _
    ccaddress:=CCRecipients, _
    bccaddress:="", _
    attachmentname:=FilePathName, _
    pathotherattachments:=AllReceipts, _
    displaymail:="yes", _
    thesignature:="Signature1", _
    thesender:="alex@soundrootsproductions.com"

    
End Function



