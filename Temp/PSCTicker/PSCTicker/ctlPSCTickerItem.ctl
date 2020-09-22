VERSION 5.00
Begin VB.UserControl ctlPSCTickerItem 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   ScaleHeight     =   1320
   ScaleWidth      =   1950
   Begin VB.Image imgScreenshot 
      Height          =   480
      Left            =   0
      Picture         =   "ctlPSCTickerItem.ctx":0000
      Top             =   840
      Width           =   480
   End
   Begin VB.Label lblScreenShot 
      BackStyle       =   0  'Transparent
      Caption         =   "(View Screenshot)"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "ctlPSCTickerItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event LinkClick(sLink As String)

Public Sub SetScreenShot(fVisible As Boolean, sLink As String)
  imgScreenshot.Visible = fVisible
  lblScreenShot.Visible = fVisible
  lblScreenShot.Tag = sLink
  ReAlign
End Sub

Public Sub SetCaption(sText As String, sLink As String)
  With lblText
    .Caption = sText
    .Tag = sLink
  End With
  ReAlign
End Sub

Private Sub imgScreenshot_Click()
  RaiseEvent LinkClick(lblScreenShot.Tag)
End Sub

Private Sub lblScreenShot_Click()
  RaiseEvent LinkClick(lblScreenShot.Tag)
End Sub

Private Sub lblText_Click()
  RaiseEvent LinkClick(lblText.Tag)
End Sub

Private Sub ReAlign()
  Dim lWidth As Long, lRows As Long
  
  lblText.Width = Width
  lWidth = TextWidth(lblText.Caption)
  lRows = lWidth \ lblText.Width
  If lWidth Mod lblText.Width > 0 Then
    lRows = lRows + 1
  End If
  If InStr(lblText.Caption, vbCrLf) > 0 Then
    lRows = lRows + 1
  End If
  lblText.Height = lRows * 195
  
  lblScreenShot.Top = lblText.Height + 160
  imgScreenshot.Top = lblText.Height + 30
  If Not imgScreenshot.Visible Then
    Height = lblText.Height + 30
  Else
    Height = imgScreenshot.Top + imgScreenshot.Height + 30
  End If
End Sub

Private Sub UserControl_Resize()
  ReAlign
End Sub
