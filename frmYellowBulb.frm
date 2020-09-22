VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmYellowBulb 
   BorderStyle     =   0  'None
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmYellowBulb.frx":0000
   ScaleHeight     =   90
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   3240
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   360
      Top             =   3240
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      Picture         =   "frmYellowBulb.frx":16722
      ScaleHeight     =   540
      ScaleWidth      =   12750
      TabIndex        =   3
      Top             =   2160
      Width           =   12750
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      Picture         =   "frmYellowBulb.frx":2CE44
      ScaleHeight     =   540
      ScaleWidth      =   12750
      TabIndex        =   2
      Top             =   480
      Width           =   12750
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      Picture         =   "frmYellowBulb.frx":43566
      ScaleHeight     =   540
      ScaleWidth      =   12750
      TabIndex        =   1
      Top             =   960
      Width           =   12750
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   360
      Top             =   2760
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   2760
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      Picture         =   "frmYellowBulb.frx":59C88
      ScaleHeight     =   540
      ScaleWidth      =   12750
      TabIndex        =   0
      Top             =   1680
      Width           =   12750
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmYellowBulb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The following code is only used to allow form draging
'from any part of it

Option Explicit


Private hRgn As Long

'Constants declaration needed for the CommonDialog
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const CC_FULLOPEN = &H2
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_RGBINIT = &H1
Private Const CC_ANYCOLOR = &H100

Private Sub Form_Deactivate()
SetWindowPos hwnd, conHwndTopmost, 0, 0, 1000, 50, conSwpNoActivate Or conSwpShowWindow

End Sub

Private Sub Form_LostFocus()
SetWindowPos hwnd, conHwndTopmost, 0, 0, 1000, 50, conSwpNoActivate Or conSwpShowWindow

End Sub



Private Sub Form_Load()
    
'transparent color is white..
CommonDialog1.Color = vbWhite
SetRegion
frmOptions.Show



If frmOptions.Option3.Value = True Then
frmYellowBulb.Timer1.Enabled = False
frmYellowBulb.Timer2.Enabled = False
frmYellowBulb.Timer3.Enabled = False
frmYellowBulb.Timer4.Enabled = False
frmYellowBulb.Picture = frmYellowBulb.Picture4.Picture
End If

If frmOptions.Option4.Value = True Then
frmYellowBulb.Timer1.Enabled = True
frmYellowBulb.Timer3.Enabled = False
frmYellowBulb.Timer4.Enabled = False
End If

If frmOptions.Option5.Value = True Then
frmYellowBulb.Timer1.Enabled = False
frmYellowBulb.Timer2.Enabled = False
frmYellowBulb.Picture = frmYellowBulb.Picture3.Picture
frmYellowBulb.Timer3.Enabled = True
End If

If frmOptions.Option6.Value = True Then

End If

    
End Sub

Private Sub Form_Paint()
SetWindowPos hwnd, conHwndTopmost, 0, 0, 1000, 50, conSwpNoActivate Or conSwpShowWindow
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Free the used memory by the Region and unload the
'Form
    If hRgn Then DeleteObject hRgn
    
End Sub





Private Sub SetRegion()
'Free the memory set
    If hRgn Then DeleteObject hRgn
'Scan the Bitmap and remove all transparent pixels from
'it, creating a new region
    hRgn = GetBitmapRegion(frmYellowBulb.Picture, CommonDialog1.Color)
'Set the Forms new Region
    SetWindowRgn frmYellowBulb.hwnd, hRgn, True
End Sub

Private Sub Image1_Click()
Unload Me
'Free the used memory by the Region and unload the
'Form
    If hRgn Then DeleteObject hRgn
End Sub

Private Sub Timer1_Timer()
Me.Picture = Picture1.Picture
Timer1.Enabled = False
Timer2.Enabled = True

End Sub

Private Sub Timer2_Timer()
Me.Picture = Picture2.Picture
Timer2.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub Timer3_Timer()
Me.Picture = Picture3.Picture
Timer4.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Me.Picture = Picture4.Picture
Timer4.Enabled = False
Timer3.Enabled = True
End Sub
