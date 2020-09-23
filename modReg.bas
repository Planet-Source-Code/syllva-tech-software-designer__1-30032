Attribute VB_Name = "modReg"
Option Explicit


Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public CheckCh(9) As Integer

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Public Const REG_SZ = 1 ' Unicode nul terminated String
    Public Const REG_DWORD = 4 ' 32-bit number
'MakeFileType "txt", "Text Document", "C:\windows\notepad.exe,0", "open", "C:\windows\notepad.exe %1", False, True

Public Function ReplaceChars(ByVal Text As String, ByVal Char As String, ReplaceChar As String) As String
    Dim counter As Integer
    
    counter = 1
    Do
        counter = InStr(counter, Text, Char)
        If counter <> 0 Then
            Mid(Text, counter, Len(ReplaceChar)) = ReplaceChar
          Else
            ReplaceChars = Text
            Exit Do
        End If
    Loop

    ReplaceChars = Text
End Function




Public Sub savekey(hKey As Long, strPath As String)
    Dim keyhand&
    r = RegCreateKey(hKey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub


Private Function GetString(hKey As Long, strPath As String, strValue As String, DefaultStr As Long)
    'EXAMPLE:
    '
    'text1.text = getstring(HKEY_CURRENT_USE
    '     R, "Software\VBW\Registry", "String")
    '
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim datas1 As String, datas2 As String
    Dim fle As Integer
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                GetString = left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
    If strBuf = "" Then GetString = DefaultStr
End Function


Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub


Function GetDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    'EXAMPLE:
    '
    'text1.text = getdword(HKEY_CURRENT_USER
    '     , "Software\VBW\Registry", "Dword")
    '
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    ' Get length\data type
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)


    If lResult = ERROR_SUCCESS Then


        If lValueType = REG_DWORD Then
            GetDWord = lBuf
        End If
        'Else
        'Call errlog("GetDWORD-" & strPath, Fals
        '     e)
    End If
    r = RegCloseKey(keyhand)
End Function


Function SaveDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    'EXAMPLE"
    '
    'Call SaveDword(HKEY_CURRENT_USER, "Soft
    '     ware\VBW\Registry", "Dword", text1.text)
    '
    '
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then
    '     Call errlog("SetDWORD", False)
    r = RegCloseKey(keyhand)
End Function


Public Function DeleteKey(ByVal hKey As Long, ByVal strKey As String)
    'EXAMPLE:
    '
    'Call DeleteKey(HKEY_CURRENT_USER, "Soft
    '     ware\VBW")
    '
    Dim r As Long
    r = RegDeleteKey(hKey, strKey)
End Function


Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '     ftware\VBW\Registry", "Dword")
    '
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function




'This is the section to read\write all the options :)

Public Function WriteOptions()
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "comment", frmDoc.rt.GetColor(cmClrComment)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "bookmark", frmDoc.rt.GetColor(cmClrBookmark)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "bookmarkbk", frmDoc.rt.GetColor(cmClrBookmarkBk)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "commentbk", frmDoc.rt.GetColor(cmClrCommentBk)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "divider", frmDoc.rt.GetColor(cmClrHDividerLines)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "highlight", frmDoc.rt.GetColor(cmClrHighlightedLine)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "keyword", frmDoc.rt.GetColor(cmClrKeyword)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "keywordbk", frmDoc.rt.GetColor(cmClrKeywordBk)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "left", frmDoc.rt.GetColor(cmClrLeftMargin)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "linenum", frmDoc.rt.GetColor(cmClrLineNumber)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "linenumbk", frmDoc.rt.GetColor(cmClrLineNumberBk)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "number", frmDoc.rt.GetColor(cmClrNumber)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "numberbk", frmDoc.rt.GetColor(cmClrNumberBk)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "operator", frmDoc.rt.GetColor(cmClrOperator)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "operatorbk", frmDoc.rt.GetColor(cmClrOperatorBk)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "scope", frmDoc.rt.GetColor(cmClrScopeKeyword)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "scopebk", frmDoc.rt.GetColor(cmClrScopeKeywordBk)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "string", frmDoc.rt.GetColor(cmClrString)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "stringbk", frmDoc.rt.GetColor(cmClrStringBk)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "tagattrib", frmDoc.rt.GetColor(cmClrTagAttributeName)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "tagattribbk", frmDoc.rt.GetColor(cmClrTagAttributeNameBk)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "tagele", frmDoc.rt.GetColor(cmClrTagElementName)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "tagelebk", frmDoc.rt.GetColor(cmClrTagElementNameBk)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "tagent", frmDoc.rt.GetColor(cmClrTagEntity)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "tagentbk", frmDoc.rt.GetColor(cmClrTagEntityBk)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "tagtxt", frmDoc.rt.GetColor(cmClrTagText)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "tagtxtbk", frmDoc.rt.GetColor(cmClrTagTextBk)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "text", frmDoc.rt.GetColor(cmClrText)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "textbk", frmDoc.rt.GetColor(cmClrTextBk)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "vdivider", frmDoc.rt.GetColor(cmClrVDividerLines)
  SaveString HKEY_CLASSES_ROOT, "Designer\colors\", "window", frmDoc.rt.GetColor(cmClrWindow)
  SaveString HKEY_CLASSES_ROOT, "Designer\options\", "selbounds", frmDoc.rt.SelBounds
  SaveString HKEY_CLASSES_ROOT, "Designer\data\", "numbering", frmDoc.rt.LineNumbering
  SaveString HKEY_CLASSES_ROOT, "Designer\data\", "lttips", frmDoc.rt.LineToolTips
  SaveString HKEY_CLASSES_ROOT, "Designer\data\", "numberingstyle", frmDoc.rt.LineNumberStyle
  SaveString HKEY_CLASSES_ROOT, "Designer\data\", "numberingstart", frmDoc.rt.LineNumberStart
  SaveString HKEY_CLASSES_ROOT, "Designer\data\", "leftmargin", frmDoc.rt.DisplayLeftMargin
  'savestring HKEY_CLASSES_ROOT, "Designer\data\", "leftmargin", frmDoc.rt.
End Function

Public Function ReadOptions(rt As CodeMax)
  Call rt.SetColor(cmClrComment, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "comment", 32896))
  Call rt.SetColor(cmClrBookmark, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "bookmark", -1))
  Call rt.SetColor(cmClrBookmarkBk, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "bookmarkbk", -1))
  Call rt.SetColor(cmClrCommentBk, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "commentbk", -1))
  Call rt.SetColor(cmClrHDividerLines, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "divider", -1))
  Call rt.SetColor(cmClrHighlightedLine, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "highlight", 65535))
  Call rt.SetColor(cmClrKeyword, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "keyword", 16711680))
  Call rt.SetColor(cmClrKeywordBk, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "keywordbk", -1))
  Call rt.SetColor(cmClrLeftMargin, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "left", 8421504))
  Call rt.SetColor(cmClrLineNumber, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "linenum", 16777215))
  Call rt.SetColor(cmClrLineNumberBk, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "linenumbk", 8421504))
  Call rt.SetColor(cmClrNumber, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "number", 0))
  Call rt.SetColor(cmClrNumberBk, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "numberbk", -1))
  Call rt.SetColor(cmClrOperator, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "operator", 255))
  Call rt.SetColor(cmClrOperatorBk, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "operatorbk", -1))
  Call rt.SetColor(cmClrScopeKeyword, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "scope", 16711680))
  Call rt.SetColor(cmClrScopeKeywordBk, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "scopebk", -1))
  Call rt.SetColor(cmClrString, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "string", 8388736))
  Call rt.SetColor(cmClrStringBk, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "stringbk", -1))
  Call rt.SetColor(cmClrTagAttributeName, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "tagattrib", 16711680))
  Call rt.SetColor(cmClrTagAttributeNameBk, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "tagattribbk", -1))
  Call rt.SetColor(cmClrTagElementName, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "tagele", 128))
  Call rt.SetColor(cmClrTagElementNameBk, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "tagelebk", -1))
  Call rt.SetColor(cmClrTagEntity, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "tagent", 255))
  Call rt.SetColor(cmClrTagEntityBk, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "tagentbk", -1))
  Call rt.SetColor(cmClrTagText, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "tagtxt", 0))
  Call rt.SetColor(cmClrTagTextBk, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "tagtxtbk", -1))
  Call rt.SetColor(cmClrText, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "text", 0))
  Call rt.SetColor(cmClrTextBk, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "textbk", -1))
  Call rt.SetColor(cmClrVDividerLines, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "vdivider", -1))
  Call rt.SetColor(cmClrWindow, GetString(HKEY_CLASSES_ROOT, "Designer\colors\", "window", -1))
  rt.SelBounds = GetString(HKEY_CLASSES_ROOT, "Designer\options\", "selbounds", True)
  rt.DisplayLeftMargin = GetString(HKEY_CLASSES_ROOT, "Designer\options\", "leftmargin", True)
  rt.LineNumbering = GetString(HKEY_CLASSES_ROOT, "Designer\data\", "numbering", True)
  rt.LineToolTips = GetString(HKEY_CLASSES_ROOT, "Designer\options\", "lttips", True)
  rt.LineNumberStyle = GetString(HKEY_CLASSES_ROOT, "Designer\data\", "numberingstyle", 1)
  rt.LineNumberStart = GetString(HKEY_CLASSES_ROOT, "Designer\data\", "numberingstart", 1)

End Function

'Window Data

Public Function WriteData()
  SaveString HKEY_CLASSES_ROOT, "Designer\window\", "windowstate", frmMain.WindowState
  frmMain.WindowState = vbNormal
  SaveString HKEY_CLASSES_ROOT, "Designer\window\", "left", frmMain.left
  SaveString HKEY_CLASSES_ROOT, "Designer\window\", "top", frmMain.top
  SaveString HKEY_CLASSES_ROOT, "Designer\window\", "width", frmMain.Width
  SaveString HKEY_CLASSES_ROOT, "Designer\window\", "height", frmMain.Height
  'SaveString HKEY_CLASSES_ROOT, "Designer\window\", "toolbar", frmMain.tBar.Visible
  'SaveString HKEY_CLASSES_ROOT, "Designer\window\", "statusbar", frmMain.stBar.Visible
End Function

Public Function ReadData()
  Dim m As Boolean
  frmMain.left = GetString(HKEY_CLASSES_ROOT, "Designer\window\", "left", 1980)
  frmMain.top = GetString(HKEY_CLASSES_ROOT, "Designer\window\", "top", 1980)
  frmMain.Width = GetString(HKEY_CLASSES_ROOT, "Designer\window\", "width", 10080)
  frmMain.Height = GetString(HKEY_CLASSES_ROOT, "Designer\window\", "height", 5640)
  frmMain.WindowState = GetString(HKEY_CLASSES_ROOT, "Designer\window\", "windowstate", 0)
  m = GetString(HKEY_CLASSES_ROOT, "Designer\window\", "toolbar", True)
  'frmMain.tBar.Visible = m
  frmMain.ToolBar.Checked = m
  m = GetString(HKEY_CLASSES_ROOT, "Designer\window\", "statusbar", True)
  'frmMain.stBar.Visible = m
  'frmMain.statusbar2.Checked = m
End Function

Public Function WriteFile()
  Dim X As Integer
  For X = 0 To 9
    SaveString HKEY_CLASSES_ROOT, "Designer\file\", "chk" & X, frmFile.chkFile(X).value
  Next
End Function

Public Function ReadFile()
  Dim X As Integer
  For X = 0 To 9
    frmFile.chkFile(X).value = GetString(HKEY_CLASSES_ROOT, "Designer\file\", "chk" & X, 0)
  Next X
End Function

Public Function WriteInput()
  SaveString HKEY_CLASSES_ROOT, "Designer\options\", "whitespace", frmMain.whitespace.Checked
  SaveString HKEY_CLASSES_ROOT, "Designer\options\", "hlline", frmMain.hlline.Checked
End Function

Public Function ReadInput()
  WhiteSpaced = GetString(HKEY_CLASSES_ROOT, "Designer\options\", "whitespace", False)
  frmMain.whitespace.Checked = WhiteSpaced
  frmDoc.rt.DisplayWhitespace = WhiteSpaced
  HighLight = GetString(HKEY_CLASSES_ROOT, "Designer\options\", "hlline", False)
  frmMain.hlline.Checked = HighLight
  If m = True Then HighLight = True
End Function

Private Sub CreateKey(ByVal hKey As Long, ByVal Key As String, Optional SubKey As Variant)

    Dim hHnd As Long
    
    If Not IsMissing(SubKey) Then
        Temp = RegCreateKey(hKey, Key & "\" & SubKey, hHnd)
        Temp = RegCloseKey(hHnd)
    Else
        Temp = RegCreateKey(hKey, Key, hHnd)
        Temp = RegCloseKey(hHnd)
    End If

End Sub

Public Sub SaveString2(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueTitle As String, ByVal ValueData As String)

    Dim hHnd As Long
    Dim Temp As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegCreateKey(hKey, KeyPath, hHnd)
    Temp = RegSetValueEx(hHnd, ValueTitle, 0, REG_SZ, ByVal ValueData, Len(ValueData))
    Temp = RegCloseKey(hHnd)

End Sub

Public Function EnableShellNew(ByVal Extension As String, ByVal ShellNew As Boolean) As Boolean
    On Error GoTo OopsShellN
    Dim dotExtension As String
    dotExtension = "." & Extension
    
    If ShellNew = True Then
        'enable
        CreateKey HKEY_CLASSES_ROOT, dotExtension, "ShellNew"
        SaveString2 HKEY_CLASSES_ROOT, dotExtension, "ShellNew", "NullFile", " "
      Else
        'disable
        DeleteKey HKEY_CLASSES_ROOT, dotExtension & "\ShellNew"
    End If
    EnableShellNew = True
    Exit Function
    
OopsShellN:
    EnableShellNew = False
    Exit Function
    Resume Next
End Function

Public Function EnableQuickView(ByVal Extension As String, ByVal QuickView As Boolean) As Boolean
    On Error GoTo QuickViewOops
    Dim Extensionfile As String
    Extensionfile = Extension & "file"
    
    If QuickView = True Then
        'enable QuickView
        CreateKey HKEY_CLASSES_ROOT, Extensionfile, "QuickView"
        SaveString2 HKEY_CLASSES_ROOT, Extensionfile, "QuickView", " ", "*"
      Else
        'disable QuickView
        DeleteKey HKEY_CLASSES_ROOT, Extensionfile & "\QuickView"
    End If
    
    EnableQuickView = True
    Exit Function
    
QuickViewOops:
    EnableQuickView = False
    Exit Function
    Resume Next
    
End Function
