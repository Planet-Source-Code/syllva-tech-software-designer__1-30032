Attribute VB_Name = "HTMLFormat"
Option Explicit

Type tag
    text As String
    start As Double
    length As Double
End Type

'*********************************************************************
Public Function SimpleFormat(target As String) As String

    SimpleFormat = ReplaceSubString(CompactFormat(target), "><", ">" & vbCrLf & "<")

End Function

'*********************************************************************
Public Function CompactFormat(target As String) As String

  Dim a As String

    a = ReplaceSubString(target, vbCrLf, "")

    a = ReplaceSubString(a, Chr$(9), " ")

    a = ReplaceSubString(a, "     ", " ")
    a = ReplaceSubString(a, "    ", " ")
    a = ReplaceSubString(a, "   ", " ")
    a = ReplaceSubString(a, "  ", " ")

    a = Clean(a)

    CompactFormat = a

End Function

'*********************************************************************
Public Function HierarchalFormat(target As String) As String
    
    target = ReplaceSubString(target, vbCrLf, "")
    target = ReplaceSubString(target, vbTab, "")
    
    target = Eformat(target)
    
    HierarchalFormat = Clean(target)

End Function

'*********************************************************************
'this lines denotes separation from public access and inner workings
'*********************************************************************

Private Function Clean(targ As String) As String

    targ = ReplaceSubString(targ, " >", ">")
    targ = ReplaceSubString(targ, "< ", "<")
    targ = ReplaceSubString(targ, "> <", "><")
    'targ = ReplaceSubString(targ, "> ", ">")
    'targ = ReplaceSubString(targ, " <", "<")

    Clean = targ

End Function

Public Function ReplaceSubString(str As String, ByVal substr As String, ByVal newsubstr As String)

  Dim pos As Double
  Dim startPos As Double
  Dim new_str As String

    startPos = 1
    pos = InStr(str, substr)
    Do While pos > 0
        new_str = new_str & Mid$(str, startPos, pos - startPos) & newsubstr
        startPos = pos + Len(substr)
        pos = InStr(startPos, str, substr)
    Loop
    new_str = new_str & Mid$(str, startPos)
    ReplaceSubString = new_str
    
End Function

Private Function Eformat(str As String) As String

    On Error Resume Next

    Dim startPos As Double ':( Move line to top of current Function
    Dim endPos As Double ':( Move line to top of current Function

    Dim indentationLevel As Double ':( Move line to top of current Function

    Dim new_str As String ':( Move line to top of current Function

      indentationLevel = 0
      startPos = 0
      endPos = 0

      If (Mid$(str, 1, 1) <> "<") Then
        
    Dim tempEnd As Double ':( Move line to top of current Function
          tempEnd = InStr(1, str, "<")
          If tempEnd = 0 Then
              tempEnd = Len(str)
          End If
        
          new_str = Mid$(str, 1, tempEnd)
    
      End If

      Do

          DoEvents

          If InStr(startPos + 1, str, "</") <> 0 And InStr(startPos + 1, str, "</") <= InStr(startPos + 1, str, "<") Then

              startPos = InStr(startPos + 1, str, "</")
              endPos = InStr(startPos + 1, str, "<")

              If endPos = 0 Then
                  endPos = Len(str) + 1
              End If

              indentationLevel = indentationLevel - 1
              new_str = new_str & vbCrLf & String$(indentationLevel, vbTab) & Mid$(str, startPos, endPos - startPos)

            Else

              startPos = InStr(startPos + 1, str, "<")
              endPos = InStr(startPos + 1, str, "<")

              If endPos = 0 Then
                  endPos = Len(str) + 1
              End If

              new_str = new_str & vbCrLf & String$(indentationLevel, vbTab) & Mid$(str, startPos, endPos - startPos)
            
    Dim tagName As String ':( Move line to top of current Function
              tagName = LCase$(returnNameOfTag(returnNextTag(str, startPos)))
              If tagName <> "br" And tagName <> "hr" And tagName <> "img" And tagName <> "meta" And tagName <> "applet" And tagName <> "p" And tagName <> "!--" And tagName <> "input" And tagName <> "!doctype" And tagName <> "area" Then
                  'If isPairedTag(tagName) Then
                  indentationLevel = indentationLevel + 1
              End If
        
          End If

      Loop While startPos > 0

      Eformat = new_str

End Function ':( On Error Resume still active

Public Function returnNextTag(ByRef str As String, ByVal start As Double) As tag

    On Error Resume Next

    Dim endPos As Double ':( Move line to top of current Function

      start = InStr(start + 1, str, "<")
      endPos = InStr(start + 1, str, ">")

      returnNextTag.text = Mid$(str, start, endPos - start + 1)
      returnNextTag.start = start
      returnNextTag.length = endPos - start

End Function ':( On Error Resume still active

Public Function returnNameOfTag(ByRef str As tag) As String

    On Error Resume Next

    Dim endPos As Double ':( Move line to top of current Function
    Dim start As Double ':( Move line to top of current Function

      start = 2
      endPos = InStr(1, str.text, " ")
      If Mid$(str.text, 2, 3) = "!--" Then
          endPos = 5
        ElseIf endPos = 0 Then
          endPos = InStr(1, str.text, ">")
      End If

      returnNameOfTag = Mid$(str.text, start, endPos - start)

End Function ':( On Error Resume still active

'
'Public Function isPairedTag(ByVal tagName As String) As Boolean
'On Error Resume Next
'
'    isPairedTag = False
'
'    Dim rcTagSearch As Recordset
'    Dim base As Database
'    Set base = frmEdit.dataAccess.Database
'    Set rcTagSearch = base.OpenRecordset("SELECT tag.tag, tag.paired From tag WHERE (((tag.tag)='" & tagName & "'))")
'
'    isPairedTag = rcTagSearch("paired")
'
'End Function

Public Function fileExist(fileName As String) As Boolean

  Dim l As Long
    
    On Error Resume Next
    
      l = FileLen(fileName)
    
      fileExist = Not (Err.Number > 0)
    
    On Error GoTo 0

End Function

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:52:44 PM) 7 + 207 = 214 Lines
