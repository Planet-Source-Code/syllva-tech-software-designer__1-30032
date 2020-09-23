Attribute VB_Name = "mMRU"
'######################################
'WonderHTML 1.2 Deluxe Edition: 2001 BETA release
'(C) Sushant S. Pandurangi, [sushant@phreaker.net]
'######################################
'For more software, visit http://sushantshome.tripod.com
'######################################
'Thanks to Andrea Batina for MRU code
Option Explicit
Dim i As Integer
Public Const SPI_SETBORDER = 6 ':( As Integer ?

Sub ShowWebMRU()

    For i = 1 To 4
        If i > frmMain.WebMRU.Count Then Exit For ':( Expand Structure
        ' Set menu caption
        frmMain.mnuWebMRU(i).Caption = "&" & i & "  " & GetFile(frmMain.WebMRU(i))
        ' Set menu tag to file name
        frmMain.mnuWebMRU(i).tag = frmMain.WebMRU(i)
        ' Show menu
        frmMain.mnuWebMRU(i).Visible = True
    Next ':( Repeat For-Variable: I
    
    For i = frmMain.WebMRU.Count + 1 To 4
        frmMain.mnuWebMRU(i).Visible = False 'Hide empty menus
    Next i

End Sub

Sub GetWebMRU()

  Dim fileName As String
    
    Set frmMain.WebMRU = New Collection 'Create new collection
    For i = 1 To 4
        fileName = ReadValue("WebMRU" & i, , "MRU Webs")
        If Len(fileName) > 2 Then
            frmMain.WebMRU.Add fileName  'Add file name to collection
        End If
    Next ':( Repeat For-Variable: I
    ShowWebMRU 'Call DisplayMRUList sub

End Sub

Sub AddWebMRU(fileName As String)

    For i = 1 To 4
        If i > frmMain.WebMRU.Count Then Exit For ':( Expand Structure
        If LCase$(frmMain.WebMRU(i)) = LCase$(fileName) Then 'If filename exist in the
            frmMain.WebMRU.Remove i                     'collection exit sub
            Exit For '>---> Next
        End If
    Next i
    
    If frmMain.WebMRU.Count > 0 Then 'If the collection is not empty
        frmMain.WebMRU.Add fileName, , 1  'add file to begining of the collecton
      Else 'else
        frmMain.WebMRU.Add fileName  'just add it
    End If
    
    If frmMain.WebMRU.Count > 4 Then 'If there are more items than 8 remove the last one
        frmMain.WebMRU.Remove 5
    End If
    
    For i = 1 To 4
        If i > frmMain.WebMRU.Count Then 'If no more files then leave it empty
            fileName = ""
          Else 'else
            fileName = frmMain.WebMRU(i) 'add it
        End If
        ' Add file to the INI
        SaveValue "WebMRU" & i, fileName, "MRU Webs"
    Next i
    GetWebMRU

End Sub

Public Sub GetFileMRU()

  Dim i As Integer ':( Duplicated Name
  Dim fileName As String
    
    Set frmMain.FileMRU = New Collection 'Create new collection
    For i = 1 To 8
        fileName = ReadValue("FileMRU" & i, , "MRU Files")
        If Len(fileName) > 2 Then
            frmMain.FileMRU.Add fileName 'Add file name to collection
        End If
    Next ':( Repeat For-Variable: I
    ShowFileMRU 'Call DisplayMRUList sub

End Sub

Public Sub ShowFileMRU()

  Dim i As Integer ':( Duplicated Name

    For i = 1 To 8
        If i > frmMain.FileMRU.Count Then Exit For ':( Expand Structure
        ' Set menu caption
        frmMain.mnuFileMRU(i).Caption = "&" & i & "  " & GetFile(frmMain.FileMRU(i))
        ' Set menu tag to file name
        frmMain.mnuFileMRU(i).tag = frmMain.FileMRU(i)
        ' Show menu
        frmMain.mnuFileMRU(i).Visible = True
    Next ':( Repeat For-Variable: I
    
    For i = frmMain.FileMRU.Count + 1 To 8
        frmMain.mnuFileMRU(i).Visible = False 'Hide empty menus
    Next ':( Repeat For-Variable: I
    
End Sub

Public Sub AddFileMRU(ByVal fileName As String)

  Dim i As Integer ':( Duplicated Name

    For i = 1 To 8
        If i > frmMain.FileMRU.Count Then Exit For ':( Expand Structure
        If LCase$(frmMain.FileMRU(i)) = LCase$(fileName) Then 'If filename exist in the
            frmMain.FileMRU.Remove i                     'collection exit sub
            Exit For '>---> Next
        End If
    Next i
    
    If frmMain.FileMRU.Count > 0 Then 'If the collection is not empty
        frmMain.FileMRU.Add fileName, , 1  'add file to begining of the collecton
      Else 'else
        frmMain.FileMRU.Add fileName 'just add it
    End If
    
    If frmMain.FileMRU.Count > 8 Then 'If there are more items than 8 remove the last one
        frmMain.FileMRU.Remove 9
    End If
    
    For i = 1 To 8
        If i > frmMain.FileMRU.Count Then 'If no more files then leave it empty
            fileName = ""
          Else 'else
            fileName = frmMain.FileMRU(i) 'add it
        End If
        ' Add file to the registry
        SaveValue "FileMRU" & i, fileName, "MRU Files"
    Next i
    GetFileMRU

End Sub

Sub Outline(lpForm As Form, lpTreeViewCtl As TreeView, lpStringHTML As String, bExitFlags As Boolean)

    On Error Resume Next 'just in case
      If bExitFlags Then Exit Sub ':( Expand Structure
      frmMain.MousePointer = 11
    Dim InPos As Long, InPos2 As Long, ThisTag As String ':( Move line to top of current Sub
    Dim SpacePos As Long, sImg As Long, Whole As String ':( Move line to top of current Sub
    Dim pParent As String, InScript As Boolean ':( Move line to top of current Sub
      InScript = False

      lpTreeViewCtl.Nodes.Clear 'I tried using WM_CLEAR, but it hangs
      If lpForm.Icon = lpForm.p2.Picture Then GoTo bye: Exit Sub 'its not an HTML doc':( Expand Structure

      AddMainNodes lpTreeViewCtl

      Do

          InPos = InStr(InPos2 + 1, lpStringHTML, "<")
          InPos2 = InStr(InPos + 1, lpStringHTML, ">")
          If InPos = 0 Then Exit Do ':( Expand Structure
          If InPos2 = 0 Then Exit Do ':( Expand Structure

          ThisTag = Mid$(lpStringHTML, InPos + 1, InPos2 - 1 - InPos)
          Whole = ThisTag 'at this time
          SpacePos = InStr(1, ThisTag, " ")
          If SpacePos > 0 Then ThisTag = left$(ThisTag, SpacePos) ':( Expand Structure
          ThisTag = Trim$(ThisTag)

          If left$(ThisTag, 2) = "!-" Then GoTo comments 'comments':( Expand Structure
          If LCase$(ThisTag) = "/script" Then InScript = False 'we're out of it':( Expand Structure
          If LCase$(ThisTag) = "/style" Then InScript = False 'we're out of it':( Expand Structure
          If left$(ThisTag, 1) = "/" Then GoTo n 'don't want end tags':( Expand Structure
          If left$(ThisTag, 1) = " " Then GoTo n 'is a bogus tag':( Expand Structure

          Select Case LCase$(ThisTag)
            Case "body", "head"
              pParent = "document"
              ThisTag = UCase$(ThisTag)
            Case "img", "bgsound"
              pParent = "images"
              ThisTag = ReadAttrib("src", Whole)
            Case "a"
              pParent = "links"
              ThisTag = ReadAttrib("href", Whole)
              If ThisTag = Whole Then
                  ThisTag = ReadAttrib("name", Whole)
                  pParent = "bookmarks" 'put in bookmarks
              End If
            Case "table"
              pParent = "tables"
            Case "form", "input", "select", "textarea"
              pParent = "forms"
              ThisTag = ReadAttrib("name", Whole)
              If ThisTag = Whole Then ThisTag = ReadAttrib("id", Whole) ':( Expand Structure
              If ThisTag = Whole Then ThisTag = ReadAttrib("value", Whole) ':( Expand Structure
              If ThisTag = Whole Then ThisTag = ReadAttrib("type", Whole) ':( Expand Structure
            Case "font"
              pParent = "styles"
              ThisTag = ReadAttrib("face", Whole)
              If ThisTag = Whole Then ThisTag = ReadAttrib("color", Whole) ':( Expand Structure
              If ThisTag = Whole Then ThisTag = ReadAttrib("size", Whole) ':( Expand Structure
              If ThisTag = Whole Then ThisTag = "FONT" ':( Expand Structure
            Case "div"
              pParent = "divisions"
              ThisTag = ReadAttrib("id", Whole)
              If ThisTag = Whole Then ThisTag = "DIV" ':( Expand Structure
            Case "script"
              pParent = "scripts"
              ThisTag = ReadAttrib("src", Whole)
              InScript = True
              If ThisTag = Whole Then ThisTag = ReadAttrib("language", Whole) ':( Expand Structure
              If ThisTag = Whole Then ThisTag = ReadAttrib("id", Whole) ':( Expand Structure
              If ThisTag = Whole Then ThisTag = ReadAttrib("type", Whole) ':( Expand Structure
            Case "layer"
              pParent = "layers"
              ThisTag = ReadAttrib("name", Whole)
              If ThisTag = Whole Then ThisTag = ReadAttrib("id", Whole) ':( Expand Structure
            Case "title"
    Dim i As Integer, t As String ':( Duplicated Name':( Move line to top of current Sub
              i = InStr(InPos2 + 1, lpStringHTML, "</")
              t = Mid$(lpStringHTML, InPos, i - InPos + 9)  '9 len of </title>
              t = Trim$(t)
              Whole = Mid$(t, 2, Len(t) - 3)
              ThisTag = Mid$(Whole, 7, Len(Whole) - 13) '7 after <title> and 13 counting </title>
              pParent = "title"
            Case "h1", "h2", "h3", "h4", "h5", "h6"
              pParent = "headings"
              ThisTag = UCase$(ThisTag)
            Case "ol", "ul", "d", "dt"
              pParent = "lists"
              ThisTag = UCase$(ThisTag)
            Case "p", "br", "li", "html", "b", "i", "u", "em", "blockquote", _
                 "", "tr", "td", "th", "center"
              GoTo n 'don't want
            Case "meta", "!doctype"
              pParent = "declare"
              ThisTag = ReadAttrib("name", Whole)
              If ThisTag = Whole Then ThisTag = ReadAttrib("http-equiv", Whole) ':( Expand Structure
              If ThisTag = Whole And left$(ThisTag, 1) = "!" Then ThisTag = "!DOCTYPE" ':( Expand Structure
            Case "!--"
comments:
              If InScript Then GoTo n 'inscript helps to not add comments which':( Expand Structure
              'are inside script or style tags
              pParent = "comments"
              ThisTag = Mid$(Whole, 4, Len(Whole) - 6): ThisTag = Trim$(ThisTag)
              If Len(ThisTag) > 9 Then ThisTag = left$(ThisTag, 9) & "..." ':( Expand Structure
            Case "style"
              InScript = True
              pParent = "other"
            Case Else
              pParent = "other"
          End Select
          lpTreeViewCtl.Nodes.Add pParent, tvwChild, "temp", ThisTag, 2
          lpTreeViewCtl.Nodes.Item("temp").tag = Whole
          lpTreeViewCtl.Nodes.Item("temp").key = ""
n:           'next
      Loop
      lpTreeViewCtl.Nodes.Item("Main").Expanded = True
      lpTreeViewCtl.Nodes.Item("document").Expanded = True
      lpTreeViewCtl.Nodes.Item("declare").Expanded = True
      lpTreeViewCtl.Nodes.Item("title").Expanded = True
      lpTreeViewCtl.SelectedItem = lpTreeViewCtl.Nodes.Item("Main")
bye:
      frmMain.MousePointer = 0

End Sub ':( On Error Resume still active

Sub AddMainNodes(lpTreeViewCtl As TreeView)

    lpTreeViewCtl.Nodes.Add , , "Main", "HTML Outline", 1
    lpTreeViewCtl.Nodes.Add "Main", tvwChild, "document", "Document", 3
    lpTreeViewCtl.Nodes.Add "Main", tvwChild, "declare", "Declaration", 3
    lpTreeViewCtl.Nodes.Add "Main", tvwChild, "title", "Page Title", 3
    lpTreeViewCtl.Nodes.Add "document", tvwChild, "bookmarks", "Anchors", 3
    lpTreeViewCtl.Nodes.Add "document", tvwChild, "comments", "Comments", 3
    lpTreeViewCtl.Nodes.Add "document", tvwChild, "divisions", "Divisions", 3
    lpTreeViewCtl.Nodes.Add "document", tvwChild, "forms", "Forms", 3
    lpTreeViewCtl.Nodes.Add "document", tvwChild, "headings", "Headings", 3
    lpTreeViewCtl.Nodes.Add "document", tvwChild, "images", "Images", 3
    lpTreeViewCtl.Nodes.Add "document", tvwChild, "layers", "Layers", 3
    lpTreeViewCtl.Nodes.Add "document", tvwChild, "links", "Links", 3
    lpTreeViewCtl.Nodes.Add "document", tvwChild, "lists", "Lists", 3
    lpTreeViewCtl.Nodes.Add "document", tvwChild, "scripts", "Scripts", 3
    lpTreeViewCtl.Nodes.Add "document", tvwChild, "styles", "Styles", 3
    lpTreeViewCtl.Nodes.Add "document", tvwChild, "tables", "Tables", 3
    lpTreeViewCtl.Nodes.Add "document", tvwChild, "other", "Others", 3

End Sub

Function ReadAttrib(ByVal lpAttribName As String, ByVal lpString As String) As String

  Dim i As Integer, ii As Integer, tmp As String, ln As Long ':( Duplicated Name

    ln = Len(lpAttribName) + 1 'for the =
    i = InStr(1, UCase$(lpString), UCase$(lpAttribName) & "=")
    If i = 0 Then ReadAttrib = lpString: Exit Function ':( Expand Structure
    ii = InStr(i + 1, lpString, " ")
    If ii = 0 Then ii = Len(lpString) 'didn't find space...':( Expand Structure
    If ii = 0 Then ReadAttrib = lpString: Exit Function ':( Expand Structure
    tmp = Mid$(lpString, i + ln, ii - i - ln + 1)
    tmp = Trim$(tmp)
    If left$(tmp, 1) = Chr$(34) Or left$(tmp, 1) = "'" Then tmp = right$(tmp, Len(tmp) - 1) ':( Expand Structure
    If right$(tmp, 1) = Chr$(34) Or right$(tmp, 1) = "'" Then tmp = left$(tmp, Len(tmp) - 1) ':( Expand Structure
    ReadAttrib = tmp

End Function

Public Function ParseInt(Expression As Variant) As Long

  'This will return only the Integer portion
  
  Dim pos As Integer, TEMP As String

    For pos = 1 To Len(Expression)
        If IsNumeric(Mid$(Expression, pos, 1)) = True Then TEMP = TEMP & CStr(Mid$(Expression, pos, 1)) ':( Expand Structure
    Next pos
    ParseInt = CLng(TEMP)

End Function

Sub AddScripts(lpText As String, lpTree As TreeView, bExitFlags As Boolean)

    On Error Resume Next
      If bExitFlags Then Exit Sub ':( Expand Structure

    Dim InPos1 As Long, InPos2 As Long, i As Long ':( Duplicated Name':( Move line to top of current Sub
    Dim VarCount, FunCount As Long, diff As Long ':( As Variant ?':( Move line to top of current Sub
    Dim ThisFun As String, ThisVar As String ':( Move line to top of current Sub

      lpTree.Nodes.Remove lpTree.Nodes("Document").index
      lpTree.Nodes.Add , , "Document", GetFile(frmMain.ActiveForm.Caption), "fp2"
      lpTree.Nodes.Add "Document", tvwChild, "functions", "Functions", "c"
      lpTree.Nodes("functions").ExpandedImage = "o"
      lpTree.Nodes.Add "Document", tvwChild, "variables", "Variables", "c"
      lpTree.Nodes("variables").ExpandedImage = "o"

      FunCount = StrCount(lpText, "function ")
      VarCount = StrCount(lpText, "var ")

      For i = 1 To FunCount
          InPos1 = InStr(InPos1 + 1, lpText, "function ")
          InPos2 = InStr(InPos1 + 1, lpText, "{")
          If InStr(InPos1 + 1, lpText, vbNewLine) < InPos2 Then InPos2 = InStr(InPos1 + 1, lpText, vbNewLine) ':( Expand Structure
          ThisFun = Mid$(lpText, InPos1 + 9, InPos2 - InPos1 - 9)
          If InStr(1, ThisFun, "(") = 0 Then GoTo n ':( Expand Structure
          lpTree.Nodes.Add "functions", tvwChild, "function " & ThisFun, GetFunctionName(ThisFun), "function"
n:
      Next i

      InPos1 = 0: InPos2 = 0

      For i = 1 To VarCount
          InPos1 = InStr(InPos1 + 1, lpText, "var ")
          InPos2 = InStr(InPos1 + 1, lpText, ";")
          If InStr(InPos1 + 1, lpText, vbNewLine) < InPos2 Then InPos2 = InStr(InPos1 + 1, lpText, vbNewLine) ':( Expand Structure
          lpTree.Nodes.Add "variables", tvwChild, "var " & Mid$(lpText, InPos1 + 4, InPos2 - InPos1 - 4), GetVarName(Mid$(lpText, InPos1 + 4, InPos2 - InPos1 - 4)), "variable"
      Next i

      lpTree.Nodes("Document").Expanded = True
      lpTree.Nodes("functions").Expanded = True

End Sub ':( On Error Resume still active

Function FileIcon(lpFileName As String) As String

    Select Case LCase$(right$(lpFileName, 3))
      Case "tml", "htm", "xml", "asp"
        FileIcon = "file"
      Case "gif", "bmp", "jpg"
        FileIcon = "image"
      Case Else
        FileIcon = "!file"
    End Select

End Function

Function DeleteFile(lpFileName As String) As Boolean

    If MsgBox("Send " & (GetFile(lpFileName)) & " to the Recycle Bin?", vbExclamation + vbOKCancel, "Delete") = vbCancel Then DeleteFile = False: Exit Function ':( Expand Structure
  Dim lpSH As SHFILEOPSTRUCT ':( Move line to top of current Function
    With lpSH
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
        .hWnd = frmMain.hWnd
        .lpszProgressTitle = "Designer"
        .pFrom = lpFileName
        .wFunc = FO_DELETE
    End With 'LPSH
    DeleteFile = SHFileOperation(lpSH) = 0

End Function

Function MoveFile(lpFileName As String, lpNewDest As String) As Boolean

  Dim lpSH As SHFILEOPSTRUCT

    With lpSH
        .fFlags = FOF_ALLOWUNDO
        .hWnd = frmMain.hWnd
        .lpszProgressTitle = "Designer"
        .pFrom = lpFileName
        .pTo = lpNewDest
        .wFunc = FO_MOVE
    End With 'LPSH
    MoveFile = SHFileOperation(lpSH) = 0

End Function

Function CopyFile(lpFileName As String, lpNewDest As String) As Boolean

  Dim lpSH As SHFILEOPSTRUCT

    With lpSH
        .fFlags = FOF_ALLOWUNDO
        .hWnd = frmMain.hWnd
        .lpszProgressTitle = "Designer"
        .pFrom = lpFileName
        .pTo = lpNewDest
        .wFunc = FO_COPY
    End With 'LPSH
    CopyFile = SHFileOperation(lpSH) = 0

End Function

Function RenameFile(lpFileName As String, lpNewName As String) As Boolean

  Dim lpSH As SHFILEOPSTRUCT

    With lpSH
        .fFlags = FOF_ALLOWUNDO
        .hWnd = frmMain.hWnd
        .lpszProgressTitle = "Designer"
        .pFrom = lpFileName
        .pTo = lpNewName
        .wFunc = FO_RENAME
    End With 'LPSH
    RenameFile = SHFileOperation(lpSH) = 0

End Function

Function GetVarName(VarName As String) As String

    On Error Resume Next
    Dim EqPos As Long ':( Move line to top of current Function
      EqPos = InStr(1, VarName, "=")
      If EqPos = 0 Then GetVarName = VarName: Exit Function ':( Expand Structure
      GetVarName = left$(VarName, EqPos - 1)
      GetVarName = Trim$(GetVarName)

End Function ':( On Error Resume still active

Function GetFunctionName(FunctionName As String) As String

    On Error Resume Next
    Dim EqPos As Long ':( Move line to top of current Function
      EqPos = InStr(1, FunctionName, "(")
      If EqPos = 0 Then GetFunctionName = FunctionName: Exit Function ':( Expand Structure
      GetFunctionName = left$(FunctionName, EqPos - 1)
      GetFunctionName = Trim$(GetFunctionName)

End Function ':( On Error Resume still active

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:52:36 PM) 10 + 460 = 470 Lines
