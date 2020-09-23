Attribute VB_Name = "mHTML"
Option Explicit

Sub Main()

    On Error Resume Next
      Screen.MousePointer = 11

    Dim lpSp As Boolean ':( Move line to top of current Sub
      lpSp = ReadValue("NoSplash", False)
      If Not lpSp Then
          Load frmAbout
      End If

      Set frmMain = New frmMDI
      Load frmMain

      frmMain.Show
      frmMain.MDIForm_Resize
      Screen.MousePointer = 0

      If ReadValue("StartupTips", 0) = 1 Then frmTip.Show vbModal ':( Expand Structure
ReDim Cmds(gintCommandCount)
PopulateCommands


'frmCleanup.Show

End Sub ':( On Error Resume still active

Function CleanUp(rtf As RichTextBox) As String

    On Error GoTo hell
  Dim Second, First As Long, InPos As Long ':( As Variant ?':( Move line to top of current Function
  Dim ThisTag, Whole, ThisText As String ':( As Variant ?':( Move line to top of current Function
    Do
one:
        First = InStr(First + 1, rtf.text, "<")
        rtf.SelStart = First - 1
        Second = InStr(rtf.SelStart + 1, rtf.text, ">")
        rtf.SelLength = Second - rtf.SelStart
        Whole = Mid$(rtf.SelText, 2, Len(rtf.SelText) - 2)
        InPos = InStr(1, Whole, " ")
        If InPos > 0 Then ThisTag = left$(Whole, InPos) Else ThisTag = Whole ':( Expand Structure
        Select Case LCase$(ThisTag)
          Case "p", "table", "/h1", "/h2", "/h3", "/h4", "/h5", "/h6"
            ThisText = vbNewLine & vbNewLine
          Case "br", "li", "/dt", "/dl"
            ThisText = vbNewLine
          Case "hr"
            ThisText = vbNewLine & String$(30, "-") & vbNewLine
          Case Else
            ThisText = ""
        End Select
        rtf.SelText = ThisText
        If InStr(1, rtf.text, "<") = 0 Then
            First = 0: Second = 0: Exit Do
          Else
            First = 0: Second = 0
        End If
    Loop
hell:
    ReplaceStuff rtf

End Function

Public Function StrCount(sString As String, sChar As String) As Long

  Dim pStep, pCount As Long ':( As Variant ?

    'Somehow I havent seen such a function in the object
    'browser till today. VB needed this one badly.
    pStep = InStr(1, sString, sChar)
    'pStep is the first occurence
    If pStep = 0 Then Exit Function 'Char dosent exist':( Expand Structure
looper:
    Do
        pStep = InStr(pStep + 1, sString, sChar)
        pCount = pCount + 1
    Loop Until pStep = 0
    StrCount = pCount

End Function

Function Words(lpText As String) As Long

  Dim lpWords As Long

    lpWords = StrCount(lpText, " ") + 1 + StrCount(lpText, ";") + StrCount(lpText, "(") + StrCount(lpText, ")")

End Function

Sub ReplaceStuff(lpRTF As RichTextBox)

    lpRTF.text = Replace$(lpRTF.text, "&nbsp;", " ")
    lpRTF.text = Replace$(lpRTF.text, "&lt;", "<")
    lpRTF.text = Replace$(lpRTF.text, "&gt;", ">")
    lpRTF.text = Replace$(lpRTF.text, "&quot;", Chr$(34))
    lpRTF.text = Replace$(lpRTF.text, "&amp;", "&")
    lpRTF.text = Replace$(lpRTF.text, "&reg;", "®")
    lpRTF.text = Replace$(lpRTF.text, "&copy;", "©")

    lpRTF.text = Replace$(lpRTF.text, "&agrave;", "à")
    lpRTF.text = Replace$(lpRTF.text, "&aacute;", "á")
    lpRTF.text = Replace$(lpRTF.text, "&acirc;", "â")
    lpRTF.text = Replace$(lpRTF.text, "&atilde;", "ã")
    lpRTF.text = Replace$(lpRTF.text, "&auml;", "ä")
    lpRTF.text = Replace$(lpRTF.text, "&aring;", "å")
    lpRTF.text = Replace$(lpRTF.text, "&aelig;", "æ")

End Sub

Function GetDay(lpDate As Date) As String

  Dim lpW As Integer, lpS As String

    lpW = Weekday(lpDate)
    Select Case lpW
      Case 1
        lpS = "Sunday"
      Case 2
        lpS = "Monday"
      Case 3
        lpS = "Tuesday"
      Case 4
        lpS = "Wednesday"
      Case 5
        lpS = "Thursday"
      Case 6
        lpS = "Friday"
      Case 7
        lpS = "Saturday"
    End Select
    GetDay = lpS

End Function

Sub ConvEntities(rtf As RichTextBox)

  'convert HTML numeric entities
  
  Dim iPos As Long, SCPos As Long, Asc_ As Long, a As String

    Do Until InStr(1, rtf.text, "&#") = 0
        iPos = InStr(iPos + 1, rtf.text, "&#")
        rtf.SelStart = iPos - 1
        SCPos = InStr(iPos + 1, rtf.text, ";")
        a = Trim$(Mid$(rtf.text, iPos + 2, SCPos - iPos - 2))
        Asc_ = CLng(a)
        rtf.SelLength = SCPos - iPos + 1 'for semicolon
        rtf.SelText = Chr$(Asc_)
    Loop

End Sub

Public Function ReadValue(key As String, Optional Default As String, Optional Section As String = "Designer")

  ' Read from INI file
  
  Dim sReturn As String

    sReturn = String$(255, Chr$(0))
    ReadValue = left$(sReturn, GetPrivateProfileString(Section, key, Default, sReturn, Len(sReturn), App.Path & "\settings.ini"))

End Function

Public Sub SaveValue(key As String, value As String, Optional Section As String = "Designer")

  ' Write to INI file

    WritePrivateProfileString Section, key, value, App.Path & "\settings.ini"

End Sub

Function SelectDir(Optional NoName As Boolean) As String

    Load frmWeb
    frmWeb.txName.Visible = Not NoName
    frmWeb.Label4.Visible = Not NoName
    frmWeb.Show vbModal
    SelectDir = ReturnedPath
    ReturnedPath = ""

End Function

Public Function Reverse(sString As String) As String

  'VB6 has this as an in-built function called
  'StrReverse(String) but I am not sure of VB5.
  
  Dim i As Integer, s As String

    For i = 1 To Len(sString)
        s = s & Mid$(sString, Len(sString) + 1 - i, 1)
    Next i
    Reverse = s

End Function

Public Function CBinary(Expression As Boolean) As Integer

  'Useful for converting BOOLs to 0 or 1. CByte() would
  'return 255 for true, which wont be useful for setting the
  'values of, for instance, a checkbox; as it uses 0 and 1.

    If Expression = False Then CBinary = 0 Else CBinary = 1 ':( Expand Structure

End Function

Function WdCount(pString As String) As Long

  'Number of words; decided using number of spaces and other characters

    WdCount = StrCount(pString, " ") + 1 + StrCount(pString, "=") + StrCount(pString, "-") + StrCount(pString, "+") + StrCount(pString, "\") + StrCount(pString, "/") + StrCount(pString, ".")

End Function

Function LnCount(pTextBox As Object) As Integer

  'Number of lines

    LnCount = SendMessage(pTextBox.hWnd, &HBA, 0, 0&)

End Function

Function SnCount(pText As String) As Integer

  'Number of sentences

    SnCount = StrCount(pText, ".")

End Function

Function Up1Level(sPath As String) As String

    On Error Resume Next
      'Name of directory up one level from that given
    Dim pos As Long, i As Integer, Dummy As String ':( Move line to top of current Function
      If right$(sPath, 1) = "\" Then sPath = left$(sPath, Len(sPath) - 1) ':( Expand Structure
      Dummy = Reverse(sPath)
      pos = InStr(1, Dummy, "\")
      Up1Level = right$(Dummy, Len(Dummy) - pos)
      Up1Level = Reverse(Up1Level)
      If right$(Up1Level, 1) = ":" Then Up1Level = Up1Level & "\" ':( Expand Structure
      If right$(Up1Level, 1) = "\" And Len(Up1Level) > 3 Then Up1Level = left$(Up1Level, Len(Up1Level) - 1) ':( Expand Structure

End Function ':( On Error Resume still active

Function GetFile(sPath As String) As String

    On Error Resume Next
      If right$(sPath, 1) = "\" Then sPath = left$(sPath, Len(sPath) - 1) ':( Expand Structure
      'Returns only file title
    Dim i, j As Integer ':( As Variant ?':( Move line to top of current Function
      i = InStr(1, Reverse(sPath), "\")
      GetFile = right$(sPath, i - 1)
      If GetFile = "" Then GetFile = sPath ':( Expand Structure

End Function ':( On Error Resume still active

Function GetPath(sPath As String) As String

  'Returns only path name without file title

    GetPath = Up1Level(sPath) & "\"

End Function

Function InitCap(sString As String) As String

  'First letter caps

    InitCap = UCase$(left$(sString, 1)) & LCase$(right$(sString, Len(sString) - 1))

End Function

Function FullPath(lpPath As String, lpFile As String) As String

    If right$(lpPath, 1) <> "\" Then lpPath = lpPath & "\" ':( Expand Structure
    FullPath = lpPath & lpFile

End Function

Function HTML() As String

  Dim AppDesc, lpHTMLStart, lpHTMLEnd, bAt As String ':( As Variant ?

    AppDesc = ReadValue("Comments")
    If AppDesc <> "" Then AppDesc = AppDesc & vbNewLine ':( Expand Structure
    lpHTMLStart = "<!DOCTYPE HTML PUBLIC " & Chr$(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr$(34) & ">" & vbNewLine & vbNewLine & "<html>" & vbNewLine & "<head>" & vbNewLine & AppDesc
    bAt = ReadValue("BodyAttrib")
    If bAt = "" Then GoTo n ':( Expand Structure
    If left$(bAt, 1) <> " " Then bAt = " " & bAt ':( Expand Structure
n:
    lpHTMLEnd = "<title>Untitled</title>" & vbNewLine & "</head>" & vbNewLine & "<body" & bAt & ">" & vbNewLine & vbNewLine & "</body>" & vbNewLine & "</html>"
    HTML = lpHTMLStart & "<meta name=" & Chr$(34) & "Author" & Chr$(34) & " content=" & Chr$(34) & ReadValue("Author") & Chr$(34) & ">" & vbNewLine & lpHTMLEnd

End Function

Sub LoadImage(lpImageFileName As String)

  Dim frmImag As New frmImage

    Load frmImag
    frmImag.tag = (lpImageFileName)
    frmImag.pB.Picture = LoadPicture(lpImageFileName)
    frmImag.Caption = (lpImageFileName)
    frmImag.Form_Resize
    AddFileMRU lpImageFileName

End Sub

Function FormsLeft()

  Dim FormsCount As Long, lpF As Form

    For Each lpF In Forms
        If lpF.BackColor = &H8000000F Then
            FormsCount = FormsCount + 1
        End If
    Next lpF
    FormsLeft = FormsCount

End Function

Public Sub SetViewMode(ByVal eViewMode As ERECViewModes, rtf As RichTextBox)

    Select Case eViewMode 'Set View Mode
      Case 0 'to No Wrap
        SendMessageLong rtf.hWnd, EM_SETTARGETDEVICE, 0, 1
      Case 1 'to Word Wrap
        SendMessageLong rtf.hWnd, EM_SETTARGETDEVICE, 0, 0
      Case 2 'to WYSIWYG
        On Error Resume Next
          SendMessageLong rtf.hWnd, EM_SETTARGETDEVICE, Printer.hdc, Printer.Width
      End Select

End Sub ':( On Error Resume still active

Public Function GetTotalLines(RichTextBox As RichTextBox)

  Dim TotalLines As Long

    TotalLines = SendMessage(RichTextBox.hWnd, EM_GETLINECOUNT, 0, 0&)
    GetTotalLines = TotalLines

End Function

Public Function GetCurrentLine(RichTextBox As RichTextBox)

  Dim CurrentLine As Long

    CurrentLine = SendMessage(RichTextBox.hWnd, EM_LINEFROMCHAR, -1, 0&) + 1
    GetCurrentLine = CurrentLine

End Function

Sub FileInfo(lpFileSize, lpFileName)

    On Error Resume Next
      Load frmFile
      frmFile.tag = lpFileName
      frmFile.Show vbModal
      frmMain.tvW.SetFocus 'on unload

End Sub ':( On Error Resume still active

Sub SetFont(lpForm As Form)

    On Error Resume Next
    Dim lpC As Control, lpFont As String ':( Move line to top of current Sub
      lpFont = ReadValue("DisplayFont")
      If Not IsFont(lpFont) Then lpFont = "Arial" ':( Expand Structure
      If Not IsFont(lpFont) Then lpFont = "MS Sans Serif" ':( Expand Structure
      For Each lpC In lpForm.Controls
          lpC.Font.Name = lpFont
      Next lpC

End Sub ':( On Error Resume still active

Function IsFont(lpFontName As String) As Boolean

  Dim i As Integer

    For i = 0 To Screen.FontCount - 1
        If Screen.Fonts(i) = lpFontName Then IsFont = True: Exit Function ':( Expand Structure
    Next i
    IsFont = False

End Function

Function StripPath(lpString As String) As String

    StripPath = Replace$(lpString, frmMain.tvW.Nodes(1).key, "")
    If left$(StripPath, 1) = "/" Or left$(StripPath, 1) = "\" Then StripPath = right$(StripPath, Len(StripPath) - 1) ':( Expand Structure

End Function

Sub AddScriptFiles()

    On Error Resume Next
    Dim i As Integer, pPar As String ':( Move line to top of current Sub
      frmMain.fSC.Path = App.Path & "\Scripts"
      frmMain.tvS.Nodes.Clear
      frmMain.tvS.Nodes.Add , , "ScriptView", "ScriptView", "fp"
      frmMain.tvS.Nodes.Add "ScriptView", tvwChild, "HTML", "HTML files", "c"
      frmMain.tvS.Nodes.Add "ScriptView", tvwChild, "Scripts", "Scripts", "c"
      frmMain.tvS.Nodes.Add "ScriptView", tvwChild, "Images", "Images", "c"
      frmMain.tvS.Nodes.Add "ScriptView", tvwChild, "Others", "Others", "c"
      frmMain.tvS.Nodes("HTML").ExpandedImage = "o"
      frmMain.tvS.Nodes("Scripts").ExpandedImage = "o"
      frmMain.tvS.Nodes("Images").ExpandedImage = "o"
      frmMain.tvS.Nodes("Others").ExpandedImage = "o"

      For i = 0 To frmMain.fSC.ListCount - 1
          Select Case LCase$(Ext(frmMain.fSC.List(i)))
            Case "xhtml", "xhtm", "xsl", "xml"
              pPar = "XML"
            Case "html", "htm", "asp", "shtml", "xml"
              pPar = "HTML"
            Case "bmp", "jpg", "gif", "ico"
              pPar = "Images"
            Case "js", "vbs"
              pPar = "Scripts"
            Case Else
              pPar = "Others"
          End Select
          frmMain.tvS.Nodes.Add pPar, tvwChild, FullPath(frmMain.fSC.Path, frmMain.fSC.List(i)), NoExt(frmMain.fSC.List(i)), FileIcon(frmMain.fSC.List(i))
      Next i
      frmMain.tvS.Nodes("ScriptView").Expanded = True

End Sub ':( On Error Resume still active

Function NoExt(lpFileName As String) As String

    On Error Resume Next
    Dim InPos As Long ':( Move line to top of current Function
      lpFileName = GetFile(lpFileName)
      InPos = InStr(1, StrReverse(lpFileName), ".")
      If InPos = 0 Then NoExt = lpFileName: Exit Function ':( Expand Structure
      NoExt = left$(lpFileName, Len(lpFileName) - InPos)

End Function ':( On Error Resume still active

Public Function RichWordOver(rch As RichTextBox, x As Single, y As Single) As String

  Dim pt As POINTAPI
  Dim pos As Long
  Dim start_pos As Long
  Dim end_pos As Long
  Dim ch As String
  Dim txt As String
  Dim txtlen As Long

    ' Convert the position to pixels.
    pt.x = x \ Screen.TwipsPerPixelX
    pt.y = y \ Screen.TwipsPerPixelY

    ' Get the character number
    pos = SendMessage(rch.hWnd, EM_CHARFROMPOS, 0&, pt)
    If pos <= 0 Then Exit Function ':( Expand Structure

    ' Find the start of the word.
    txt = rch.text
    For start_pos = pos To 1 Step -1
        ch = Mid$(rch.text, start_pos, 1)
        ' Allow digits, letters, and underscores.
        If Not ((ch >= "0" And ch <= "9") Or (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or ch = "_" Or ch = ":" Or ch = "/" Or ch = ".") Then Exit For ':( Expand Structure
    Next start_pos
    start_pos = start_pos + 1

    ' Find the end of the word.
    txtlen = Len(txt)
    For end_pos = pos To txtlen
        ch = Mid$(txt, end_pos, 1)
        ' Allow digits, letters, and underscores.
        If Not ((ch >= "0" And ch <= "9") Or (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or ch = "_" Or ch = ":" Or ch = "/" Or ch = ".") Then Exit For ':( Expand Structure
    Next end_pos
    end_pos = end_pos - 1

    If start_pos <= end_pos Then _
       RichWordOver = Mid$(txt, start_pos, end_pos - start_pos + 1) ':( Expand Structure

End Function

Function Ext(lpFileName As String) As String

  Dim iPos As Long

    iPos = InStr(1, StrReverse(lpFileName), ".")
    If iPos = 0 Then Ext = lpFileName: Exit Function ':( Expand Structure
    Ext = right$(lpFileName, iPos - 1)

End Function

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:52:40 PM) 8 + 489 = 497 Lines
