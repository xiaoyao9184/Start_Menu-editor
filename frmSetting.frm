VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   4095
   ClientLeft      =   4560
   ClientTop       =   4125
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdapply 
      Caption         =   "应用"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.CheckBox chkPicture 
      Caption         =   "加载外部图片"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame fraPicture 
      Caption         =   "图片目录"
      Height          =   2775
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   6375
      Begin VB.TextBox txtIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1680
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txtWallpaper 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1320
         Width           =   4455
      End
      Begin VB.CommandButton cmdIcon 
         Caption         =   "状态栏目录"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton optPictureSuff_Other 
         Caption         =   "默认"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optPictureSuff_Other 
         Caption         =   "其他目录"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdWallpaper 
         Caption         =   "背景图片"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "选择目录下任何一个GIF文件即可指定此目录为状态栏目录"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   6015
      End
      Begin VB.Label Label2 
         Caption         =   "程序目录\icon\"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "必须同时指定背景图片和图标目录"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   840
         Width           =   4455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
txtWallpaper.Text = GetINI("Setting", "WallpaperPath", App.Path & "\Config.ini")
txtIcon.Text = GetINI("Setting", "IconPath", App.Path & "\Config.ini")
optPictureSuff_Other(IIf(GetINI("Setting", "PictureSuff_Other", App.Path & "\Config.ini") = 0, 0, 1)).Value = True
optPictureSuff_Other_Click (IIf(GetINI("Setting", "PictureSuff_Other", App.Path & "\Config.ini") = 0, 0, 1))
chkPicture.Value = GetINI("Setting", "PictureFT", App.Path & "\Config.ini")
chkPicture_Click
End Sub
Private Sub chkPicture_Click()
If chkPicture.Value = 1 Then
    optPictureSuff_Other(0).Enabled = True
    optPictureSuff_Other(1).Enabled = True
    optPictureSuff_Other(IIf(GetINI("Setting", "PictureSuff_Other", App.Path & "\Config.ini") = 0, 0, 1)).Value = True
    optPictureSuff_Other_Click (IIf(GetINI("Setting", "PictureSuff_Other", App.Path & "\Config.ini") = 0, 0, 1))
Else
    optPictureSuff_Other(0).Enabled = False
    optPictureSuff_Other(1).Enabled = False
    cmdWallpaper.Enabled = False
    cmdIcon.Enabled = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub

Private Sub optPictureSuff_Other_Click(Index As Integer)
If Index = 0 Then
    cmdWallpaper.Enabled = False
    cmdIcon.Enabled = False
    
    txtWallpaper.Text = App.Path & "\Default\icon\Wallpaper.jpg"
    txtIcon.Text = App.Path & "\Default\icon\"
Else
    cmdWallpaper.Enabled = True
    cmdIcon.Enabled = True
    
    txtWallpaper.Text = GetINI("Setting", "WallpaperPath", App.Path & "\Config.ini")
    txtIcon.Text = GetINI("Setting", "IconPath", App.Path & "\Config.ini")
End If

txtWallpaper.ToolTipText = txtWallpaper.Text
txtIcon.ToolTipText = txtIcon.Text
End Sub
Private Sub cmdWallpaper_Click()
CommonDialog1.Filter = "JPEG(*.jpg)|*.jpg|BMP(*.bmp)|*.bmp|GIF(*.gif)|*.gif"
CommonDialog1.CancelError = True
On Error Resume Next
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
    txtWallpaper.Text = CommonDialog1.FileName
    txtWallpaper.ToolTipText = CommonDialog1.FileName
End Sub

Private Sub cmdIcon_Click()
CommonDialog1.Filter = GetINI("lng", "cmdIcon_CF", App.Path & "\Config.ini")
CommonDialog1.CancelError = True
On Error Resume Next
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
    txtIcon.Text = CurDir()
    txtIcon.ToolTipText = CurDir()
End Sub
Private Sub cmdSave_Click()
WriteINI "Setting", "PictureFT", chkPicture.Value, App.Path & "\Config.ini"
WriteINI "Setting", "PictureSuff_Other", IIf(optPictureSuff_Other(1).Value = True, 1, 0), App.Path & "\Config.ini"
WriteINI "Setting", "WallpaperPath", txtWallpaper.Text, App.Path & "\Config.ini"
WriteINI "Setting", "IconPath", txtIcon.Text, App.Path & "\Config.ini"
cmdapply.Enabled = True
End Sub
Private Sub cmdapply_Click()
apply_Picture
cmdapply.Enabled = False
End Sub
