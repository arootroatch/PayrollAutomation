Attribute VB_Name = "EmailPayStubsSub"
Option Explicit


Sub EmailPayStubs()
    'My additions to make it a for loop,15-Feb-2023
    
    Dim FirstPayStubStart As Integer
    FirstPayStubStart = 44
    Dim FirstPayStubEnd As Integer
    FirstPayStubEnd = 85
    Dim SecondPayStubStart As Integer
    SecondPayStubStart = 87
    Dim PayStubLength As Integer
    PayStubLength = SecondPayStubStart - FirstPayStubStart
    Dim NumberMailed As Integer
    NumberMailed = 0
    Dim RecipientEmail As String
    Dim EmailSubject As String
    EmailSubject = Range("R2").Value
    Dim i As Integer
    
    
    
    'Ron de Bruin, 27-Dec-2022
    'Do not forget to add the custom functions into your workbook
    'More Information : https://macexcel.com/examples/mailpdf/macmailexamples/
    'You can test the macro without changing anything
    Dim FileName As String
    Dim FolderName As String
    Dim Folderstring As String
    Dim FilePathName As String
    Dim strbody As String

    'Check for AppleScriptTask script file that we must use to create the mail
    If CheckAppleScriptTaskExcelScriptFile(ScriptFileName:="RDBMacMail2.scpt") = False Then
        MsgBox "Sorry the RDBMacMail2.scpt is not in the correct location"
        Exit Sub
    End If
    
    'Create folder in the Office folder that we use to save the PDF
    Folderstring = CreateFolderinMacOffice(NameFolder:="RDBMailTempFolder")
    
    'Create the body text in the strbody string
    strbody = "Paystub for the most recent pay period. See attached."

    For i = 1 To 19
        If (Range("A" & FirstPayStubStart).Value = 0 Or Range("A" & FirstPayStubStart).Value = "") = False Then
        
            RecipientEmail = Range("A" & FirstPayStubStart + 1).Value
            
            'Enter the name of the pdf file without extension, Date and Time in this example
            FileName = "Pay Stub " & i
        
            'And create the file path string. Do not change this 2 lines
            FilePathName = Folderstring & Application.PathSeparator & FileName & ".pdf"
        
            'Create the PDF, you not have to edit this code block
            Range("A" & FirstPayStubStart & ":O" & FirstPayStubEnd).ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
            FilePathName, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False
        
            'Function call to create the mail, see this page for more info about the arguments
            'https://macexcel.com/examples/mailpdf/macmailexamples/
            'Do not change the attachmentname argument in this PDF function call
            MacExcelWithMacMailPDFCatalinaAndUp subject:=EmailSubject, _
            mailbody:=strbody, _
            toaddress:=RecipientEmail, _
            ccaddress:="", _
            bccaddress:="", _
            attachmentname:=FilePathName, _
            pathotherattachments:="", _
            displaymail:="no", _
            thesignature:="", _
            thesender:="alex@soundrootsproductions.com"
            
            NumberMailed = NumberMailed + 1
        End If
        FirstPayStubStart = FirstPayStubStart + PayStubLength
        FirstPayStubEnd = FirstPayStubEnd + PayStubLength
    Next i
    
    If NumberMailed = 1 Then
        MsgBox (NumberMailed & " pay stub sent via email!")
    Else
        MsgBox (NumberMailed & " pay stubs sent via email!")
    End If
End Sub


