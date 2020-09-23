VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmSpooky 
   Caption         =   "Form1"
   ClientHeight    =   825
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3945
   Icon            =   "Spooky.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   825
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1020
      Top             =   450
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   210
      TabIndex        =   0
      Top             =   135
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "frmSpooky"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    PlayAVIFrom MMControl1, 0

End Sub

Private Sub Form_Load()
    LoadAVI MMControl1, Me, App.Path & "\mitewalk.avi"

End Sub

Private Sub MMControl1_StatusUpdate()

If MMControl1.Mode = mciModeStop Then PlayAVIFrom MMControl1, 0

End Sub

Private Sub Timer1_Timer()

WHwnd& = FindWindow("ComboLBox", vbNullString)
If MMControl1.hWndDisplay <> WHwnd& Then
    MMControl1.hWndDisplay = WHwnd&
End If

End Sub
