Attribute VB_Name = "mDECL"
'######################################
'WonderHTML 1.2 Deluxe Edition: 2001 BETA release
'(C) Sushant S. Pandurangi, [sushant@phreaker.net]
'######################################
'For more software, visit http://sushantshome.tripod.com
'######################################
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Any) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Const EM_UNDO = &HC7 ':( As Integer ?
Public Const WM_CLEAR = &H303 ':( As Integer ?
Public Const WM_KILLFOCUS = &H8 ':( As Integer ?

Public Const WM_USER = &H400 ':( As Integer ?
Public Const EM_LINEINDEX = &HBB ':( As Integer ?
Public Const EM_SETTARGETDEVICE = (WM_USER + 72) ':( As Integer ?
Public Const EM_GETLINECOUNT = &HBA ':( As Integer ?
Public Const EM_LINEFROMCHAR = &HC9 ':( As Integer ?

Public Const FLAG_RO = cdlOFNPathMustExist + cdlOFNFileMustExist + cdlOFNOverwritePrompt + cdlOFNReadOnly ':( As Integer ?
Public Const FLAG_RW = cdlOFNPathMustExist + cdlOFNFileMustExist + cdlOFNOverwritePrompt ':( As Integer ?

Public ReturnedPath As String

Public frmMain As frmMDI

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long

Public Enum ERECViewModes
    ercDefault = 0
    ercWordWrap = 1
    ercWYSIWYG = 2
End Enum

'Public Type RECT
'    left As Long
'    top As Long
''    right As Long
'    bottom As Long
'End Type

Public Const EM_CHARFROMPOS = &HD7 ':( As Integer ?

Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type

Public Const FO_MOVE = &H1 ':( As Integer ?
Public Const FO_RENAME = &H4 ':( As Integer ?
Public Const FO_COPY = &H2 ':( As Integer ?
Public Const FO_DELETE = &H3 ':( As Integer ?
Public Const FOF_ALLOWUNDO = &H40 ':( As Integer ?
Public Const FOF_NOCONFIRMATION = &H10 ':( As Integer ?

Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Const Domains = "http:// www. .com .net .org .edu .ac .mil .gov ftp:// gopher:// telnet:// news: mailto: wais: .in .pk .au .uk"

Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:52:42 PM) 87 + 0 = 87 Lines
