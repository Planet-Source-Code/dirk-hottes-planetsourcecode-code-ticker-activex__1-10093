VERSION 5.00
Begin VB.UserControl ctlFileDownload 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   375
   ScaleHeight     =   390
   ScaleWidth      =   375
End
Attribute VB_Name = "ctlFileDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "KERNEL32" () As Long

Public IUrl As String
Public IContent As String
Public IKey As Long
Public IAsFile As Boolean

Public Event IReadComplete()

Public Sub ICancel()
  On Error Resume Next
  CancelAsyncRead IKey
End Sub

Public Sub IDownload()
  IKey = GetTickCount
  AsyncRead IUrl, vbAsyncTypeFile, IKey
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
  Dim sBuffer As String, iFile As Integer
  
  On Error Resume Next
  IContent = ""
  If IAsFile Then
    IContent = AsyncProp.Value
  Else
    iFile = FreeFile
    Open AsyncProp.Value For Input As iFile
    Do Until EOF(iFile)
      Line Input #iFile, sBuffer
      IContent = IContent & sBuffer
    Loop
    Close iFile
    Kill AsyncProp.Value
  End If
  
  RaiseEvent IReadComplete
End Sub

Private Sub UserControl_Initialize()
  IAsFile = False
End Sub
