VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Power"
      Height          =   1935
      Left            =   4440
      TabIndex        =   29
      Top             =   2520
      Width           =   1815
      Begin VB.Label Label1 
         Caption         =   "Function Not Added"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   1455
      End
   End
   Begin MSComctlLib.ImageCombo ColCombo 
      Height          =   330
      Left            =   2160
      TabIndex        =   27
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      OLEDropMode     =   1
      Locked          =   -1  'True
      ImageList       =   "ImageList1"
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4560
      Top             =   6360
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4080
      Top             =   6360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3000
      Top             =   6360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2520
      Top             =   6360
   End
   Begin VB.PictureBox pre3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4080
      Picture         =   "frmOptions.frx":0CCA
      ScaleHeight     =   270
      ScaleWidth      =   1440
      TabIndex        =   26
      Top             =   5640
      Width           =   1440
   End
   Begin VB.PictureBox pre4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4080
      Picture         =   "frmOptions.frx":214C
      ScaleHeight     =   270
      ScaleWidth      =   1440
      TabIndex        =   25
      Top             =   6000
      Width           =   1440
   End
   Begin VB.PictureBox pre2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   2520
      Picture         =   "frmOptions.frx":35CE
      ScaleHeight     =   270
      ScaleWidth      =   1440
      TabIndex        =   24
      Top             =   6000
      Width           =   1440
   End
   Begin VB.PictureBox pre1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   2520
      Picture         =   "frmOptions.frx":4A50
      ScaleHeight     =   270
      ScaleWidth      =   1440
      TabIndex        =   23
      Top             =   5640
      Width           =   1440
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "frmOptions.frx":5ED2
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   21
      Top             =   5640
      Width           =   495
   End
   Begin VB.Frame Frame8 
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   4560
      Width           =   6135
      Begin VB.CommandButton Command1 
         Caption         =   "Apply"
         Height          =   375
         Left            =   5160
         TabIndex        =   19
         Top             =   280
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Exit Lights"
         Height          =   375
         Left            =   1080
         TabIndex        =   18
         Top             =   280
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Help"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   280
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Ok"
         Height          =   375
         Left            =   3240
         TabIndex        =   16
         Top             =   280
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Top             =   280
         Width           =   855
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Music"
      Height          =   1815
      Left            =   4440
      TabIndex        =   13
      Top             =   600
      Width           =   1815
      Begin VB.Label Label3 
         Caption         =   "Function Not Added"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Misc"
      Height          =   855
      Left            =   2040
      TabIndex        =   8
      Top             =   3600
      Width           =   2295
      Begin VB.CommandButton Command9 
         Caption         =   "Calendar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Color"
      Height          =   975
      Left            =   2040
      TabIndex        =   7
      Top             =   2520
      Width           =   2295
      Begin MSComctlLib.ImageCombo ColCombo1 
         Height          =   330
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         OLEDropMode     =   1
         Locked          =   -1  'True
         ImageList       =   "ImageList2"
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bulbs"
      Height          =   1815
      Left            =   2040
      TabIndex        =   6
      Top             =   600
      Width           =   2295
      Begin VB.OptionButton Option10 
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Missing"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   34
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Broken"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   33
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Burnt Out"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   32
         Top             =   960
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         X1              =   1080
         X2              =   1080
         Y1              =   960
         Y2              =   1680
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   1080
         X2              =   1080
         Y1              =   960
         Y2              =   1680
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Flash Rate"
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
      Begin VB.PictureBox pre 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   180
         Picture         =   "frmOptions.frx":6B9C
         ScaleHeight     =   270
         ScaleWidth      =   1440
         TabIndex        =   22
         Top             =   1320
         Width           =   1440
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   1515
         TabIndex        =   3
         Top             =   360
         Width           =   1575
         Begin MSComctlLib.Slider Slider1 
            Height          =   495
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            _Version        =   393216
            LargeChange     =   100
            Min             =   1
            Max             =   5
            SelStart        =   1
            Value           =   1
         End
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Very Slow"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Flash Mode"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1815
      Begin VB.OptionButton Option6 
         Caption         =   "Chase"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton Option5 
         Caption         =   "All Together"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Alternate"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "None"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Text            =   "500"
      Top             =   5760
      Width           =   1575
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":801E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":8CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":99D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":A6B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":B38E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":C06A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":CD46
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":DA22
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":E6FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":F3DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":100B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":10D92
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":11A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1274A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":13426
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":14102
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":14DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":15ABA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1440
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":16796
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":17472
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1814E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":18E2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":19B06
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1A7E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1B4BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1C19A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1CE76
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   -120
      Picture         =   "frmOptions.frx":1DB52
      Top             =   0
      Width           =   6720
   End
   Begin VB.Menu menuBar 
      Caption         =   "MenuBar"
      Visible         =   0   'False
      Begin VB.Menu mnuMain 
         Caption         =   "About Holiday Lights"
         Index           =   0
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Options..."
         Index           =   1
      End
      Begin VB.Menu mnuMain 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Show Calendar Events"
         Index           =   3
      End
      Begin VB.Menu mnuMain 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Remove Lights"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim t As NOTIFYICONDATA

Private Sub ColCombo_Click()
Command1.SetFocus
If ColCombo.ComboItems(2).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(3).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(4).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(5).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(6).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(7).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(8).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(9).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(10).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(11).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(12).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(13).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(14).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(15).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(16).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(17).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
If ColCombo.ComboItems(18).Selected = True Then
ColCombo1.ComboItems(1).Selected = True
End If
End Sub

Private Sub ColCombo1_Click()
Command1.SetFocus
If ColCombo1.ComboItems(2).Selected = True Then
ColCombo.ComboItems(1).Selected = True
End If
If ColCombo1.ComboItems(3).Selected = True Then
ColCombo.ComboItems(1).Selected = True
End If
If ColCombo1.ComboItems(4).Selected = True Then
ColCombo.ComboItems(1).Selected = True
End If
If ColCombo1.ComboItems(5).Selected = True Then
ColCombo.ComboItems(1).Selected = True
End If
If ColCombo1.ComboItems(6).Selected = True Then
ColCombo.ComboItems(1).Selected = True
End If
If ColCombo1.ComboItems(7).Selected = True Then
ColCombo.ComboItems(1).Selected = True
End If
If ColCombo1.ComboItems(8).Selected = True Then
ColCombo.ComboItems(1).Selected = True
End If
If ColCombo1.ComboItems(9).Selected = True Then
ColCombo.ComboItems(1).Selected = True
End If
End Sub

Private Sub Command1_Click()
' Begin Snow Bulb Code
'======================================
If ColCombo.Text = "SnowFlakes" Then
Unload frmRedBulb
Unload frmYellowBulb
Unload frmBlueBulb
Unload frmRoundBulb
frmSnowBulb.Picture = frmSnowBulb.Picture4.Picture
Load frmSnowBulb
frmSnowBulb.Show
'======================================
If Slider1.Value = "5" Then
frmSnowBulb.Timer1.Enabled = True
frmSnowBulb.Timer1.Interval = "100"
frmSnowBulb.Timer2.Interval = "100"
frmSnowBulb.Timer3.Interval = "100"
frmSnowBulb.Timer4.Interval = "100"
End If
If Slider1.Value = "4" Then
frmSnowBulb.Timer1.Enabled = True
frmSnowBulb.Timer1.Interval = "200"
frmSnowBulb.Timer2.Interval = "200"
frmSnowBulb.Timer3.Interval = "200"
frmSnowBulb.Timer4.Interval = "200"
End If
If Slider1.Value = "3" Then
frmSnowBulb.Timer1.Enabled = True
frmSnowBulb.Timer1.Interval = "300"
frmSnowBulb.Timer2.Interval = "300"
frmSnowBulb.Timer3.Interval = "300"
frmSnowBulb.Timer4.Interval = "300"
End If
If Slider1.Value = "2" Then
frmSnowBulb.Timer1.Enabled = True
frmSnowBulb.Timer1.Interval = "400"
frmSnowBulb.Timer2.Interval = "400"
frmSnowBulb.Timer3.Interval = "400"
frmSnowBulb.Timer4.Interval = "400"
End If
If Slider1.Value = "1" Then
frmSnowBulb.Timer1.Enabled = True
frmSnowBulb.Timer1.Interval = "500"
frmSnowBulb.Timer2.Interval = "500"
frmSnowBulb.Timer3.Interval = "500"
frmSnowBulb.Timer4.Interval = "500"
End If

If Option3.Value = True Then
frmSnowBulb.Timer1.Enabled = False
frmSnowBulb.Timer2.Enabled = False
frmSnowBulb.Timer3.Enabled = False
frmSnowBulb.Timer4.Enabled = False
Unload frmRedBulb
Unload frmYellowBulb
Unload frmBlueBulb
Unload frmRoundBulb
frmSnowBulb.Picture = frmSnowBulb.Picture4.Picture
End If
If Option4.Value = True Then
frmSnowBulb.Timer1.Enabled = True
frmSnowBulb.Timer3.Enabled = False
frmSnowBulb.Timer4.Enabled = False
End If
If Option5.Value = True Then
frmSnowBulb.Timer1.Enabled = False
frmSnowBulb.Timer2.Enabled = False
frmSnowBulb.Picture = frmSnowBulb.Picture4.Picture
frmSnowBulb.Timer3.Enabled = True
End If

If Option6.Value = True Then
End If
End If
' End Snow Bulb Code
'======================================
' Begin Round Bulb Code
'======================================
If ColCombo.Text = "Round" Then
Unload frmRedBulb
Unload frmYellowBulb
Unload frmBlueBulb
Unload frmSnowBulb
frmRoundBulb.Picture = frmRoundBulb.Picture4.Picture
Load frmRoundBulb
frmRoundBulb.Show
'======================================
If Slider1.Value = "5" Then
frmRoundBulb.Timer1.Enabled = True
frmRoundBulb.Timer1.Interval = "100"
frmRoundBulb.Timer2.Interval = "100"
frmRoundBulb.Timer3.Interval = "100"
frmRoundBulb.Timer4.Interval = "100"
End If
If Slider1.Value = "4" Then
frmRoundBulb.Timer1.Enabled = True
frmRoundBulb.Timer1.Interval = "200"
frmRoundBulb.Timer2.Interval = "200"
frmRoundBulb.Timer3.Interval = "200"
frmRoundBulb.Timer4.Interval = "200"
End If
If Slider1.Value = "3" Then
frmRoundBulb.Timer1.Enabled = True
frmRoundBulb.Timer1.Interval = "300"
frmRoundBulb.Timer2.Interval = "300"
frmRoundBulb.Timer3.Interval = "300"
frmRoundBulb.Timer4.Interval = "300"
End If
If Slider1.Value = "2" Then
frmRoundBulb.Timer1.Enabled = True
frmRoundBulb.Timer1.Interval = "400"
frmRoundBulb.Timer2.Interval = "400"
frmRoundBulb.Timer3.Interval = "400"
frmRoundBulb.Timer4.Interval = "400"
End If
If Slider1.Value = "1" Then
frmRoundBulb.Timer1.Enabled = True
frmRoundBulb.Timer1.Interval = "500"
frmRoundBulb.Timer2.Interval = "500"
frmRoundBulb.Timer3.Interval = "500"
frmRoundBulb.Timer4.Interval = "500"
End If

If Option3.Value = True Then
frmRoundBulb.Timer1.Enabled = False
frmRoundBulb.Timer2.Enabled = False
frmRoundBulb.Timer3.Enabled = False
frmRoundBulb.Timer4.Enabled = False
Unload frmRedBulb
Unload frmYellowBulb
Unload frmBlueBulb
Unload frmSnowBulb
frmRoundBulb.Picture = frmRoundBulb.Picture4.Picture
End If
If Option4.Value = True Then
frmRoundBulb.Timer1.Enabled = True
frmRoundBulb.Timer3.Enabled = False
frmRoundBulb.Timer4.Enabled = False
End If
If Option5.Value = True Then
frmRoundBulb.Timer1.Enabled = False
frmRoundBulb.Timer2.Enabled = False
frmRoundBulb.Picture = frmRoundBulb.Picture4.Picture
frmRoundBulb.Timer3.Enabled = True
End If

If Option6.Value = True Then
End If
End If
' End Round Bulb Code
'======================================
' Begin Blue Bulb Code
'======================================
If ColCombo1.Text = "Blue" Then
Unload frmRedBulb
Unload frmYellowBulb
Unload frmBlueBulb
Unload frmRoundBulb
Unload frmSnowBulb
Load frmBlueBulb
frmBlueBulb.Show
'======================================
If Slider1.Value = "5" Then
frmBlueBulb.Timer1.Enabled = True
frmBlueBulb.Timer1.Interval = "100"
frmBlueBulb.Timer2.Interval = "100"
frmBlueBulb.Timer3.Interval = "100"
frmBlueBulb.Timer4.Interval = "100"
End If
If Slider1.Value = "4" Then
frmBlueBulb.Timer1.Enabled = True
frmBlueBulb.Timer1.Interval = "200"
frmBlueBulb.Timer2.Interval = "200"
frmBlueBulb.Timer3.Interval = "200"
frmBlueBulb.Timer4.Interval = "200"
End If
If Slider1.Value = "3" Then
frmBlueBulb.Timer1.Enabled = True
frmBlueBulb.Timer1.Interval = "300"
frmBlueBulb.Timer2.Interval = "300"
frmBlueBulb.Timer3.Interval = "300"
frmBlueBulb.Timer4.Interval = "300"
End If
If Slider1.Value = "2" Then
frmBlueBulb.Timer1.Enabled = True
frmBlueBulb.Timer1.Interval = "400"
frmBlueBulb.Timer2.Interval = "400"
frmBlueBulb.Timer3.Interval = "400"
frmBlueBulb.Timer4.Interval = "400"
End If
If Slider1.Value = "1" Then
frmBlueBulb.Timer1.Enabled = True
frmBlueBulb.Timer1.Interval = "500"
frmBlueBulb.Timer2.Interval = "500"
frmBlueBulb.Timer3.Interval = "500"
frmBlueBulb.Timer4.Interval = "500"
End If
If Option3.Value = True Then
frmBlueBulb.Timer1.Enabled = False
frmBlueBulb.Timer2.Enabled = False
frmBlueBulb.Timer3.Enabled = False
frmBlueBulb.Timer4.Enabled = False
frmBlueBulb.Picture = frmBlueBulb.Picture4.Picture
End If
If Option4.Value = True Then
frmBlueBulb.Timer1.Enabled = True
frmBlueBulb.Timer3.Enabled = False
frmBlueBulb.Timer4.Enabled = False
End If
If Option5.Value = True Then
frmBlueBulb.Timer1.Enabled = False
frmBlueBulb.Timer2.Enabled = False
frmBlueBulb.Picture = frmBlueBulb.Picture3.Picture
frmBlueBulb.Timer3.Enabled = True
End If

If Option6.Value = True Then
End If
End If
' End Blue Bulb Code
'======================================
' Begin red Bulb Code
'======================================
If ColCombo1.Text = "Red" Then
Unload frmSnowBulb
Unload frmYellowBulb
Unload frmBlueBulb
Unload frmRoundBulb
Load frmRedBulb
frmRedBulb.Show
'======================================
If Slider1.Value = "5" Then
frmRedBulb.Timer1.Enabled = True
frmRedBulb.Timer1.Interval = "100"
frmRedBulb.Timer2.Interval = "100"
frmRedBulb.Timer3.Interval = "100"
frmRedBulb.Timer4.Interval = "100"
End If
If Slider1.Value = "4" Then
frmRedBulb.Timer1.Enabled = True
frmRedBulb.Timer1.Interval = "200"
frmRedBulb.Timer2.Interval = "200"
frmRedBulb.Timer3.Interval = "200"
frmRedBulb.Timer4.Interval = "200"
End If
If Slider1.Value = "3" Then
frmRedBulb.Timer1.Enabled = True
frmRedBulb.Timer1.Interval = "300"
frmRedBulb.Timer2.Interval = "300"
frmRedBulb.Timer3.Interval = "300"
frmRedBulb.Timer4.Interval = "300"
End If
If Slider1.Value = "2" Then
frmRedBulb.Timer1.Enabled = True
frmRedBulb.Timer1.Interval = "400"
frmRedBulb.Timer2.Interval = "400"
frmRedBulb.Timer3.Interval = "400"
frmRedBulb.Timer4.Interval = "400"
End If
If Slider1.Value = "1" Then
frmRedBulb.Timer1.Enabled = True
frmRedBulb.Timer1.Interval = "500"
frmRedBulb.Timer2.Interval = "500"
frmRedBulb.Timer3.Interval = "500"
frmRedBulb.Timer4.Interval = "500"
End If


If Option3.Value = True Then
frmRedBulb.Timer1.Enabled = False
frmRedBulb.Timer2.Enabled = False
frmRedBulb.Timer3.Enabled = False
frmRedBulb.Timer4.Enabled = False
frmRedBulb.Picture = frmRedBulb.Picture4.Picture
End If

If Option4.Value = True Then
frmRedBulb.Timer1.Enabled = True
frmRedBulb.Timer3.Enabled = False
frmRedBulb.Timer4.Enabled = False
End If

If Option5.Value = True Then
frmRedBulb.Timer1.Enabled = False
frmRedBulb.Timer2.Enabled = False
frmRedBulb.Picture = frmRedBulb.Picture3.Picture
frmRedBulb.Timer3.Enabled = True
End If

If Option6.Value = True Then

End If
End If
' End Red Bulb Code
'======================================
' Begin Yellow Bulb Code
'======================================
If ColCombo1.Text = "Yellow" Then
Unload frmRedBulb
Unload frmBlueBulb
Unload frmRoundBulb
Unload frmSnowBulb
Load frmYellowBulb
frmYellowBulb.Show
'======================================
If Slider1.Value = "5" Then
frmYellowBulb.Timer1.Enabled = True
frmYellowBulb.Timer1.Interval = "100"
frmYellowBulb.Timer2.Interval = "100"
frmYellowBulb.Timer3.Interval = "100"
frmYellowBulb.Timer4.Interval = "100"
End If
If Slider1.Value = "4" Then
frmYellowBulb.Timer1.Enabled = True
frmYellowBulb.Timer1.Interval = "200"
frmYellowBulb.Timer2.Interval = "200"
frmYellowBulb.Timer3.Interval = "200"
frmYellowBulb.Timer4.Interval = "200"
End If
If Slider1.Value = "3" Then
frmYellowBulb.Timer1.Enabled = True
frmYellowBulb.Timer1.Interval = "300"
frmYellowBulb.Timer2.Interval = "300"
frmYellowBulb.Timer3.Interval = "300"
frmYellowBulb.Timer4.Interval = "300"
End If
If Slider1.Value = "2" Then
frmYellowBulb.Timer1.Enabled = True
frmYellowBulb.Timer1.Interval = "400"
frmYellowBulb.Timer2.Interval = "400"
frmYellowBulb.Timer3.Interval = "400"
frmYellowBulb.Timer4.Interval = "400"
End If
If Slider1.Value = "1" Then
frmYellowBulb.Timer1.Enabled = True
frmYellowBulb.Timer1.Interval = "500"
frmYellowBulb.Timer2.Interval = "500"
frmYellowBulb.Timer3.Interval = "500"
frmYellowBulb.Timer4.Interval = "500"
End If



If Option3.Value = True Then
frmYellowBulb.Timer1.Enabled = False
frmYellowBulb.Timer2.Enabled = False
frmYellowBulb.Timer3.Enabled = False
frmYellowBulb.Timer4.Enabled = False
frmYellowBulb.Picture = frmYellowBulb.Picture4.Picture
End If

If Option4.Value = True Then
frmYellowBulb.Timer1.Enabled = True
frmYellowBulb.Timer3.Enabled = False
frmYellowBulb.Timer4.Enabled = False
End If

If Option5.Value = True Then
frmYellowBulb.Timer1.Enabled = False
frmYellowBulb.Timer2.Enabled = False
frmYellowBulb.Picture = frmYellowBulb.Picture3.Picture
frmYellowBulb.Timer3.Enabled = True
End If

If Option6.Value = True Then

End If
End If
' End Yellow Bulb Code
'======================================



End Sub








Private Sub Command5_Click()
    t.cbSize = Len(t)
    t.hwnd = Picture1.hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t

End

End Sub

Private Sub Command7_Click()
Command1_Click
Me.Hide
End Sub

Private Sub Command8_Click()
Me.Hide
End Sub

Private Sub Form_Load()
    t.cbSize = Len(t)
    t.hwnd = Picture1.hwnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    'this is where the Form's Icon gets called
          t.hIcon = Me.Icon
    'this is where the tool tip goes
          t.szTip = "Lights 2001" & Chr$(0)
    
    Shell_NotifyIcon NIM_ADD, t
    
    '-----------------------------------
ColCombo.ComboItems.Add 1, , "Normal", 1, 1
ColCombo.ComboItems.Add 2, , "Round", 2, 2
ColCombo.ComboItems.Add 3, , "SnowFlakes", 3, 3
ColCombo.ComboItems.Add 4, , "Bells", 4, 4
ColCombo.ComboItems.Add 5, , "CandyCane", 5, 5
ColCombo.ComboItems.Add 6, , "Hearts", 6, 6
ColCombo.ComboItems.Add 7, , "Flags", 7, 7
ColCombo.ComboItems.Add 8, , "Easter", 8, 8
ColCombo.ComboItems.Add 9, , "Pumpkin", 9, 9
ColCombo.ComboItems.Add 10, , "Smiley", 10, 10
ColCombo.ComboItems.Add 11, , "Shamrock", 11, 11
ColCombo.ComboItems.Add 12, , "Lantern", 12, 12
ColCombo.ComboItems.Add 13, , "Hot Peppers", 13, 13
ColCombo.ComboItems.Add 14, , "Transparent", 14, 14
ColCombo.ComboItems.Add 15, , "X-mas Trees", 15, 15
ColCombo.ComboItems.Add 16, , "Stockings", 16, 16
ColCombo.ComboItems.Add 17, , "Holly", 17, 17
ColCombo.ComboItems.Add 18, , "Snowman", 18, 18
ColCombo.ComboItems(1).Selected = True
ColCombo.Refresh
'------------------------------------------
ColCombo1.ComboItems.Add 1, , "None", 9, 9
ColCombo1.ComboItems.Add 2, , "Blue", 2, 2
ColCombo1.ComboItems.Add 3, , "Purple", 1, 1
ColCombo1.ComboItems.Add 4, , "Red", 3, 3
ColCombo1.ComboItems.Add 5, , "Green", 4, 4
ColCombo1.ComboItems.Add 6, , "Yellow", 5, 5
ColCombo1.ComboItems.Add 7, , "White", 6, 6
ColCombo1.ComboItems.Add 8, , "Red & Green", 7, 7
ColCombo1.ComboItems.Add 9, , "Multi Colors", 8, 8
ColCombo1.ComboItems(2).Selected = True
ColCombo1.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    t.cbSize = Len(t)
    t.hwnd = Picture1.hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub

Private Sub Form_Terminate()
   t.cbSize = Len(t)
    t.hwnd = Picture1.hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub

Private Sub Form_Unload(Cancel As Integer)
    t.cbSize = Len(t)
    t.hwnd = Picture1.hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub

Private Sub Option3_Click()

If Option3.Value = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
pre.Picture = pre2.Picture
End If
End Sub

Private Sub Option4_Click()


If Option4.Value = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = True
End If

End Sub

Private Sub Option5_Click()
If Option5.Value = True Then
Timer3.Enabled = False
Timer4.Enabled = False
Timer1.Enabled = True
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static rec As Boolean, msg As Long
    Dim RetVal As String
    Dim returnstring
    Dim retvalue
    msg = X / Screen.TwipsPerPixelX
    If rec = False Then
    rec = True
    Select Case msg
    Case WM_LBUTTONDOWN:
    Me.PopupMenu menuBar
    'This is what yo put here when the icon in the tray
    'is pressed!
    Case WM_LBUTTONDBLCLK 'not used in this program
    Case WM_LBUTTONUP: 'not used in this program
    Case WM_RBUTTONDBLCLK: 'not used in this program
    Case WM_RBUTTONDOWN:   'not used in this program
    Case WM_RBUTTONUP:
    'if Right Mouse Button is down then
    'Bring up the Popup Menu
    Me.PopupMenu menuBar
    End Select
    rec = False
    End If
End Sub

Private Sub mnuMain_Click(Index As Integer)
'allow user to remove Icon from System Tray

Select Case Index
    Case 0
    MsgBox "About"
       
    Case 1
    Load frmOptions
    frmOptions.Show
           
    Case 3
    MsgBox "Show Calendar Events"
    
    Case 5
    End
    
    Case Else
    
End Select

End Sub
Private Sub Slider1_Change()
If Slider1.Value = "5" Then
Text1.Text = 100
Timer1.Interval = 100
Timer2.Interval = 100
Timer3.Interval = 100
Timer4.Interval = 100
Label2.Caption = "Very Fast"
End If
If Slider1.Value = "4" Then
Text1.Text = 200
Timer1.Interval = 200
Timer2.Interval = 200
Timer3.Interval = 200
Timer4.Interval = 200
Label2.Caption = "Fast"
End If
If Slider1.Value = "3" Then
Text1.Text = 300
Timer1.Interval = 300
Timer2.Interval = 300
Timer3.Interval = 300
Timer4.Interval = 300
Label2.Caption = "Medium"
End If
If Slider1.Value = "2" Then
Text1.Text = 400
Timer1.Interval = 400
Timer2.Interval = 400
Timer3.Interval = 400
Timer4.Interval = 400
Label2.Caption = "Slow"
End If
If Slider1.Value = "1" Then
Text1.Text = 500
Timer1.Interval = 500
Timer2.Interval = 500
Timer3.Interval = 500
Timer4.Interval = 500
Label2.Caption = "Very Slow"
End If
End Sub

Private Sub Timer1_Timer()
pre.Picture = pre2.Picture
Timer1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
pre.Picture = pre1.Picture
Timer2.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub Timer3_Timer()
pre.Picture = pre3.Picture
Timer3.Enabled = False
Timer4.Enabled = True
End Sub

Private Sub Timer4_Timer()
pre.Picture = pre4.Picture
Timer4.Enabled = False
Timer3.Enabled = True
End Sub


