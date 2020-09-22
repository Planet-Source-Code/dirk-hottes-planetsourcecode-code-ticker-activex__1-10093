Attribute VB_Name = "basPSC"
Public Const psTickerLink = "http://www.planet-source-code.com/vb/linktous/ScrollingCode.asp?lngWId=1"
Public Const psStartTicker = "<marquee behavior=scroll direction=up"
Public Const psStartItem = "<font face=verdana,arial><font size=1><b>"
Private psStartLink As String

Public Type TickerItem
  Caption As String
  HasScreenShot As Boolean
  Info As String
  Link As String
  Screenshot As String
End Type

Public pscItems() As TickerItem

Public Sub ExtractItemFromHTML(sContent As String)
  Dim sItems() As String, iItem As Integer
  Dim iStartLink As String, iEndLink As Integer
  Dim iStartDescription As Integer
  
  psStartLink = "<a target=" & Chr$(34) & "_top" & Chr$(34) & " href=" & Chr$(34)
  sItems() = Split(sContent, psStartItem)
  ReDim pscItems(UBound(sItems()))
  For iItem = 1 To UBound(sItems())
    sItems(iItem) = Mid$(sItems(iItem), InStr(1, sItems(iItem), "<"))
    If Mid$(sItems(iItem), 1, 8) = "<a href=" Then
      pscItems(iItem).HasScreenShot = True
      pscItems(iItem).Screenshot = Mid$(sItems(iItem), 10, InStr(10, sItems(iItem), Chr$(34)) - 10)
    Else
      pscItems(iItem).HasScreenShot = False
      pscItems(iItem).Screenshot = ""
    End If
    iStartLink = InStr(1, sItems(iItem), psStartLink) + Len(psStartLink)
    iEndLink = InStr(iStartLink, sItems(iItem), Chr(34) & ">") + 2
    iStartDescription = iEndLink
    iEndLink = iEndLink - iStartLink - 2
    pscItems(iItem).Link = Mid$(sItems(iItem), iStartLink, iEndLink)
    iEndLink = InStr(iStartDescription, sItems(iItem), "</a>")
    iEndLink = iEndLink - iStartDescription
    pscItems(iItem).Caption = CleanString(Mid$(sItems(iItem), iStartDescription, iEndLink))
    iStartLink = InStr(iStartDescription, sItems(iItem), "<BR>") + 4
    iEndLink = InStr(iStartLink, sItems(iItem), "</b>")
    iEndLink = iEndLink - iStartLink
    pscItems(iItem).Info = CleanString(Mid$(sItems(iItem), iStartLink, iEndLink))
  Next
End Sub

Public Sub CreateTicker(ctlTicker As ctlPSCTicker)
  Dim iItem As Integer
  
  ctlTicker.Clear
  For iItem = 1 To UBound(pscItems())
    ctlTicker.AddItem pscItems(iItem).Caption & vbCrLf & pscItems(iItem).Info, pscItems(iItem).Link, pscItems(iItem).Screenshot
  Next
  ctlTicker.StartTicker
End Sub

Private Function CleanString(sText As String)
  Dim iChar As Integer, sBuffer As String
  
  sBuffer = ""
  For iChar = 1 To Len(sText)
    Select Case Asc(Mid$(sText, iChar, 1))
    Case 9, 34
    Case Else
      sBuffer = sBuffer & Mid$(sText, iChar, 1)
    End Select
  Next
  CleanString = sBuffer
End Function

