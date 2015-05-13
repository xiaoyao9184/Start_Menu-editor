VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于..."
   ClientHeight    =   3225
   ClientLeft      =   5670
   ClientTop       =   4620
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4905
   Begin VB.Timer TmrScroll 
      Interval        =   100
      Left            =   1320
      Top             =   1920
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000F&
      Height          =   1935
      Left            =   1920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmAbout.frx":0000
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label LblScroll 
      Caption         =   $"frmAbout.frx":0248
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   2400
      Width           =   12495
   End
   Begin VB.Label warn 
      Caption         =   "警告：目前仅在程序在实验阶段，作者不负责使用此程序所产生的任何后果"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Image Img1 
      Height          =   600
      Left            =   0
      Picture         =   "frmAbout.frx":02D9
      Top             =   0
      Width           =   210
   End
   Begin VB.Image Img2 
      Height          =   1680
      Left            =   0
      Picture         =   "frmAbout.frx":06A3
      Top             =   360
      Width           =   1785
   End
   Begin VB.Label name 
      Caption         =   "Start_Menu编辑器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
 Unload Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TmrScroll.Interval = 100
End Sub
Private Sub LblScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TmrScroll.Interval = 15
End Sub
Private Sub TmrScroll_Timer()       '定时器开始
If LblScroll.Left > -CInt(LblScroll.Width) Then
 LblScroll.Left = LblScroll.Left - 20 '滚屏字幕的左边从窗口右边向左移动
Else
 LblScroll.Left = frmAbout.Width
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub
