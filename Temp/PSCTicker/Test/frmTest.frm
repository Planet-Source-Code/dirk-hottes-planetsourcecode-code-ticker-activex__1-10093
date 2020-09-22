VERSION 5.00
Object = "*\A..\PSCTicker\prjPSCTicker.vbp"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Planetsourcecode VB Ticker ActiveX Control"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame fraBottomBack 
      BorderStyle     =   0  'Kein
      Height          =   705
      Left            =   0
      TabIndex        =   5
      Top             =   4800
      Width           =   7455
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   375
         Left            =   6240
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Default         =   -1  'True
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   7560
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   7560
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   900
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7455
      Begin VB.Line Line4 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   8560
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Image imgLogo 
         Height          =   735
         Left            =   6600
         Picture         =   "frmTest.frx":0000
         Top             =   60
         Width           =   735
      End
      Begin VB.Label lblScrnSubTitle 
         BackColor       =   &H80000005&
         Caption         =   "This ticker shows the last code submissions on Planetsourcecode"
         Height          =   255
         Left            =   675
         TabIndex        =   4
         Top             =   405
         Width           =   4785
      End
      Begin VB.Label lblScrnTitle 
         BackColor       =   &H80000005&
         Caption         =   "Planetsourcecode Ticker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   345
         TabIndex        =   3
         Top             =   180
         Width           =   6315
      End
   End
   Begin VB.PictureBox picTicker 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   480
      ScaleHeight     =   3315
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
      Begin PSCTicker.ctlPSCTicker ctlPSCTicker 
         Height          =   3135
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   5530
      End
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Press 'Refresh' to show the latest code submissions on Planetsourcecode."
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label lblVBTicker 
      BackStyle       =   0  'Transparent
      Caption         =   "VB Ticker"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Private Const HWND_NOTOPMOST = -2

Public Sub ThinBorder(ByVal hwnd As Long, ByVal bState As Boolean)
  Dim lStyle As Long
  
  lStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
  If bState Then
    lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
  Else
    lStyle = lStyle Or WS_EX_CLIENTEDGE And Not WS_EX_STATICEDGE
  End If
  SetWindowLong hwnd, GWL_EXSTYLE, lStyle
  SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

Private Sub cmdClose_Click()
  ctlPSCTicker.StopTicker
  Unload Me
End Sub

Private Sub cmdRefresh_Click()
  ctlPSCTicker.RefreshTicker
End Sub

Private Sub Form_Load()
  ThinBorder picTicker.hwnd, True
End Sub
