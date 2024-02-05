Attribute VB_Name = "MacMailFunctions"
Function MacExcelWithMacMailCatalinaAndUp(subject As String, mailbody As String, _
    toaddress As String, ccaddress As String, _
    bccaddress As String, displaymail As String, _
    attachmentname As String, pathotherattachments As String, _
    thesignature As String, thesender As String, Optional FileFormat As Long)
    ' Ron de Bruin 27-Dec-2022
    ' More Information : https://macexcel.com/examples/mailpdf/macmailexamples/
    
    Dim FileExtStr As String, FileFormatNum As Long
    Dim TempFileFolder As String, FilePath As String
    Dim OtherFileAtachments As String
    Dim ScriptStr As String, RunMyScript As String
    Dim AttachmentsArr() As String, OtherAttachmentFilePath As Variant
    
    'If you mail the ActiveSheet or more then one sheet we run this code block
    If attachmentname <> "" Then
        ' Save the new temporary workbook and close it
        Select Case FileFormat
            Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If ActiveWorkbook.HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56: FileExtStr = ".xls": FileFormatNum = 56
            Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
        End Select
  
        TempFileFolder = CreateFolderinMacOffice(NameFolder:="RDBMailTempFolder")
        FilePath = TempFileFolder & Application.PathSeparator & attachmentname & FileExtStr
        With ActiveWorkbook
            .SaveAs FilePath, FileFormat:=FileFormatNum
            .Close SaveChanges:=False
        End With
    End If

    'Check if the file or files exists in pathotherattachments
    If pathotherattachments <> "" Then
        AttachmentsArr = Split(pathotherattachments, ",")
        For Each OtherAttachmentFilePath In AttachmentsArr
            If Not FileOrFolderExistsOnYourMac(CStr(OtherAttachmentFilePath), 1) = True Then
                pathotherattachments = ""
            End If
        Next
        If pathotherattachments = "" Then MsgBox "One or more File path(s) in the pathotherattachments argument not exist on your Mac"
    End If

    'Build the AppleScriptTask string
    ScriptStr = subject & ";" & mailbody & ";" & toaddress & ";" & ccaddress & ";" & _
        bccaddress & ";" & displaymail & ";" & FilePath & ";" & pathotherattachments & ";" & thesignature & ";" & thesender
 
    'Call the RDBMacmail2.scpt script file with the AppleScriptTask function to run the script
    RunMyScript = AppleScriptTask("RDBMacMail2.scpt", "CreateMailInCatalinaAndUp", CStr(ScriptStr))
   
    'Delete the file after we create the mail if filepath <>""
    If attachmentname <> "" Then Kill FilePath
End Function


Function CheckAppleScriptTaskExcelScriptFile(ScriptFileName As String) As Boolean
    'Check if the AppleScriptTask script file exists in the com.microsoft.Excel folder
    'Ron de Bruin : 27 Dec-2022
    Dim AppleScriptTaskFolder As String
    Dim TestStr As String

    AppleScriptTaskFolder = MacScript("return POSIX path of (path to desktop folder) as string")
    AppleScriptTaskFolder = Replace(AppleScriptTaskFolder, "/Desktop", "") & _
        "Library/Application Scripts/com.microsoft.Excel/"

    On Error Resume Next
    TestStr = Dir(AppleScriptTaskFolder & ScriptFileName)
    On Error GoTo 0
    If TestStr = vbNullString Then
        CheckAppleScriptTaskExcelScriptFile = False
    Else
        CheckAppleScriptTaskExcelScriptFile = True
    End If
End Function


Function CreateFolderinMacOffice(NameFolder As String) As String
    'Function to create folder if it not exists in the Microsoft Office Folder
    'Ron de Bruin : 27-Dec-2022
    Dim OfficeFolder As String
    Dim PathToFolder As String
    Dim TestStr As String

    OfficeFolder = MacScript("return POSIX path of (path to desktop folder) as string")
    OfficeFolder = Replace(OfficeFolder, "/Desktop", "") & _
        "Library/Group Containers/UBF8T346G9.Office/"

    PathToFolder = OfficeFolder & NameFolder

    On Error Resume Next
    TestStr = Dir(PathToFolder & "*", vbDirectory)
    On Error GoTo 0
    If TestStr = vbNullString Then
        MkDir PathToFolder
    End If
    CreateFolderinMacOffice = PathToFolder
End Function


Function FileOrFolderExistsOnYourMac(FileOrFolderstr As String, FileOrFolder As Long) As Boolean
    'Ron de Bruin : 27-Dec-2022, for Excel 2016 and higher
    'Function to test if a file or folder exist on your Mac
    'Use 1 as second argument for File and 2 for Folder
    'In the earlier version of this function I used Dir(FileOrFolderstr & "*") to get it working
    'This bug seems to be fixed now so I removed the & "*"
    Dim ScriptToCheckFileFolder As String
    Dim FileOrFolderPath As String
    
    If FileOrFolder = 1 Then
        'File test
        On Error Resume Next
        FileOrFolderPath = Dir(FileOrFolderstr)
        On Error GoTo 0
        If Not FileOrFolderPath = vbNullString Then FileOrFolderExistsOnYourMac = True
    Else
        'folder test
        On Error Resume Next
        FileOrFolderPath = Dir(FileOrFolderstr, vbDirectory)
        On Error GoTo 0
        If Not FileOrFolderPath = vbNullString Then FileOrFolderExistsOnYourMac = True
    End If
End Function


Function MacExcelWithMacMailPDFCatalinaAndUp(subject As String, mailbody As String, _
    toaddress As String, ccaddress As String, _
    bccaddress As String, displaymail As String, _
    attachmentname As String, pathotherattachments As String, _
    thesignature As String, thesender As String)
    ' Ron de Bruin 27-Dec-2022
    ' More Information : https://macexcel.com/examples/mailpdf/macmailexamples/
    
    Dim AttachmentsArr() As String, OtherAttachmentFilePath As Variant
    Dim ScriptStr As String, RunMyScript As String
    
    'Check if the file or files exists in pathotherattachments
    If pathotherattachments <> "" Then
        AttachmentsArr = Split(pathotherattachments, ",")
        For Each OtherAttachmentFilePath In AttachmentsArr
            If Not FileOrFolderExistsOnYourMac(CStr(OtherAttachmentFilePath), 1) = True Then
                pathotherattachments = ""
            End If
        Next
        If pathotherattachments = "" Then MsgBox "One or more File path(s) in the pathotherattachments argument not exist on your Mac"
    End If
    
    'Build the AppleScriptTask string
    ScriptStr = subject & ";" & mailbody & ";" & toaddress & ";" & ccaddress & ";" & _
        bccaddress & ";" & displaymail & ";" & attachmentname & ";" & pathotherattachments & ";" & thesignature & ";" & thesender
        
    'Call the RDBMacmail2.scpt script file with the AppleScriptTask function to run the script
    RunMyScript = AppleScriptTask("RDBMacMail2.scpt", "CreateMailInCatalinaAndUp", CStr(ScriptStr))
    
    
    'Delete the pdf file we just mailed--I added the InString function so it wouldn't delete invoices
    Dim isTemp As Integer
    isTemp = InStr(attachmentname, "RDBMailTempFolder")
    If isTemp > 0 Then
        If attachmentname <> "" Then Kill attachmentname
    End If
End Function




