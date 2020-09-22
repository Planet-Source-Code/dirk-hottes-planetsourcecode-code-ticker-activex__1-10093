VERSION 5.00
Begin VB.UserControl ctlPSCTicker 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1980
   ScaleHeight     =   3600
   ScaleWidth      =   1980
   Begin PSCTicker.ctlPSCTickerItem ctlPSCTickerItem 
      Height          =   555
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   979
   End
   Begin PSCTicker.ctlFileDownload ctlTickerSrc 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
   End
   Begin VB.Timer ctlTimer 
      Enabled         =   0   'False
      Interval        =   125
      Left            =   720
      Top             =   1920
   End
End
Attribute VB_Name = "ctlPSCTicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Event LinkClick(sLink As String)

Const lStep = 30

Property Get hwnd() As Long
  hwnd = UserControl.hwnd
End Property

Private Sub ctlPSCTickerItem_LinkClick(Index As Integer, sLink As String)
  RaiseEvent LinkClick(sLink)
End Sub

Private Sub SetStartPosition()
  Dim iItem As Integer, lCurrentTop As Long
  
  ctlPSCTickerItem(1).Top = Height
  For iItem = 2 To ctlPSCTickerItem.UBound
    ctlPSCTickerItem(iItem).Top = ctlPSCTickerItem(iItem - 1).Top + ctlPSCTickerItem(iItem - 1).Height + 360
  Next
End Sub

Private Sub ctlTimer_Timer()
  Dim iItem As Integer, lCurrentTop As Long
  
  For iItem = 1 To ctlPSCTickerItem.UBound
    ctlPSCTickerItem(iItem).Top = ctlPSCTickerItem(iItem).Top - lStep
    If ctlPSCTickerItem(iItem).Top <= -1 * Abs(ctlPSCTickerItem(iItem).Height + 360) Then
      If iItem > 1 Then
        ctlPSCTickerItem(iItem).Top = ctlPSCTickerItem(iItem - 1).Top + ctlPSCTickerItem(iItem - 1).Height + 360
      Else
        ctlPSCTickerItem(iItem).Top = ctlPSCTickerItem(ctlPSCTickerItem.UBound).Top + ctlPSCTickerItem(ctlPSCTickerItem.UBound).Height + 360
      End If
    End If
  Next
End Sub

Private Sub UserControl_Resize()
  Dim iItem As Integer
  
  For iItem = 0 To ctlPSCTickerItem.UBound
    ctlPSCTickerItem(iItem).Width = Width - 60
  Next
End Sub

Public Sub AddItem(sCaption As String, sLink As String, sPreview As String)
  Load ctlPSCTickerItem(ctlPSCTickerItem.UBound + 1)
  With ctlPSCTickerItem(ctlPSCTickerItem.UBound)
    .SetCaption sCaption, sLink
    .SetScreenShot IIf(Len(sPreview) > 0, True, False), sPreview
    .Width = Width - 60
    .Visible = False
  End With
End Sub

Public Sub Clear()
  Dim iItem As Integer
  
  For iItem = ctlPSCTickerItem.UBound To 1 Step -1
    Unload ctlPSCTickerItem(iItem)
  Next
  StopTicker
End Sub

Public Sub StartTicker()
  Dim iItem As Integer
  
  SetStartPosition
  For iItem = ctlPSCTickerItem.UBound To 1 Step -1
    ctlPSCTickerItem(iItem).Visible = True
  Next
  ctlTimer.Enabled = True
End Sub

Public Sub StopTicker()
  ctlTimer.Enabled = False
End Sub

Public Sub RefreshTicker()
  ctlTickerSrc.IUrl = psTickerLink
  ctlTickerSrc.IAsFile = False
  ctlTickerSrc.IDownload
End Sub

Private Sub ctlTickerSrc_IReadComplete()
  ExtractItemFromHTML ctlTickerSrc.IContent
  CreateTicker Me
End Sub

Private Sub UserControl_Terminate()
  StopTicker
End Sub
