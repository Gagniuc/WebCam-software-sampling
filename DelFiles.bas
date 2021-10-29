Attribute VB_Name = "DelFiles"
'Api for Moving Files
Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

'Api for creating directory and deleting folders/files
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long
Declare Function CreateDirectory& Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpnewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES)
Declare Function GetDesktopWindow& Lib "user32" ()

Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Const FO_MOVE As Long = &H1
Public Const FO_COPY As Long = &H2
Public Const FO_DELETE As Long = &H3
Public Const FO_RENAME As Long = &H4

Public Const FOF_MULTIDESTFILES As Long = &H1
Public Const FOF_CONFIRMMOUSE As Long = &H2
Public Const FOF_SILENT As Long = &H4
Public Const FOF_RENAMEONCOLLISION As Long = &H8
Public Const FOF_NOCONFIRMATION As Long = &H10
Public Const FOF_WANTMAPPINGHANDLE As Long = &H20
Public Const FOF_CREATEPROGRESSDLG As Long = &H0
Public Const FOF_ALLOWUNDO As Long = &H40
Public Const FOF_FILESONLY As Long = &H80
Public Const FOF_SIMPLEPROGRESS As Long = &H100
Public Const FOF_NOCONFIRMMKDIR As Long = &H200

Public Function File_Delete(path As String) As Boolean
On Error Resume Next
    Dim FileOperation As SHFILEOPSTRUCT
    Dim sTempFilename As String * 100
    Dim sSendMeToTheBin As String
    sSendMeToTheBin = path
    With FileOperation
        .wFunc = FO_DELETE
        .pFrom = sSendMeToTheBin
        .fFlags = FOF_ALLOWUNDO
    End With
    SHFileOperation FileOperation
    File_Delete = FileOperation.fAnyOperationsAborted
End Function

