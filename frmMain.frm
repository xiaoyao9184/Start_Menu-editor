VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Start_Menu �༭��"
   ClientHeight    =   5310
   ClientLeft      =   3390
   ClientTop       =   3165
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPict 
      Height          =   270
      Left            =   3960
      TabIndex        =   6
      Text            =   "PY"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtNum 
      Height          =   270
      Left            =   3960
      TabIndex        =   3
      Text            =   "Num"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtOffset 
      Height          =   270
      Left            =   6480
      TabIndex        =   2
      Text            =   "offset"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtT 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2160
      TabIndex        =   39
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   38
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdRes 
      Caption         =   "ˢ��"
      Height          =   495
      Left            =   1560
      TabIndex        =   20
      Top             =   4200
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ԫ��"
      Height          =   2415
      Left            =   2880
      TabIndex        =   29
      Top             =   2520
      Width           =   5295
      Begin VB.PictureBox picIcons 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   360
         ScaleHeight     =   63
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   143
         TabIndex        =   22
         Top             =   1200
         Width           =   2175
         Begin VB.Image Image1 
            Height          =   375
            Index           =   0
            Left            =   0
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.OptionButton Path_code 
         Caption         =   "����"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   18
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtPath 
         Height          =   375
         Left            =   2760
         TabIndex        =   19
         Text            =   "path"
         Top             =   1680
         Width           =   2415
      End
      Begin VB.OptionButton Path_code 
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   17
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         Text            =   "name"
         Top             =   720
         Width           =   2415
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   2055
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3625
         _Version        =   393217
         Indentation     =   9
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
         OLEDropMode     =   1
      End
      Begin VB.Image imgIcons 
         Height          =   375
         Left            =   2760
         Top             =   240
         Width           =   375
      End
      Begin VB.Label LblName 
         Caption         =   "����"
         Height          =   255
         Left            =   3240
         TabIndex        =   30
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox txtY 
      Height          =   270
      Left            =   7680
      TabIndex        =   5
      Text            =   "Y"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtX 
      Height          =   270
      Left            =   5880
      TabIndex        =   4
      Text            =   "X"
      Top             =   840
      Width           =   495
   End
   Begin VB.ComboBox cbbfontNO 
      Height          =   300
      Left            =   5040
      TabIndex        =   7
      Text            =   "��ѡ��������"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtStart 
      Height          =   270
      Left            =   3960
      TabIndex        =   1
      Text            =   "code"
      Top             =   360
      Width           =   855
   End
   Begin VB.PictureBox Wallpaper 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3300
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   220
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   176
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   2640
      Begin VB.Image Iicon 
         Height          =   255
         Index           =   0
         Left            =   240
         Top             =   2520
         Width           =   255
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   9
         Left            =   2310
         Picture         =   "frmMain.frx":370F
         Top             =   0
         Width           =   330
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   8
         Left            =   2025
         Picture         =   "frmMain.frx":37CF
         Top             =   0
         Width           =   285
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   7
         Left            =   1755
         Top             =   0
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   6
         Left            =   1515
         Picture         =   "frmMain.frx":385F
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   5
         Left            =   1245
         Picture         =   "frmMain.frx":390A
         Top             =   0
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   4
         Left            =   975
         Picture         =   "frmMain.frx":39BA
         Top             =   0
         Width           =   270
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   3
         Left            =   780
         Picture         =   "frmMain.frx":3A45
         Top             =   0
         Width           =   195
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   2
         Left            =   540
         Picture         =   "frmMain.frx":3AB2
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         DataMember      =   "225"
         Height          =   225
         Index           =   1
         Left            =   330
         Picture         =   "frmMain.frx":3B30
         Top             =   0
         Width           =   210
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":3BB4
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   6360
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   9
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblColor 
      Caption         =   "�߿�2"
      Height          =   255
      Index           =   11
      Left            =   4600
      TabIndex        =   35
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label LblRegister 
      Caption         =   "��������"
      Height          =   255
      Left            =   2880
      TabIndex        =   23
      Top             =   360
      Width           =   975
   End
   Begin VB.Label LblEorC 
      Caption         =   "Chinese"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   6840
      TabIndex        =   14
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   5520
      TabIndex        =   13
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   12
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   11
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblColor 
      Caption         =   "�߿�4"
      Height          =   255
      Index           =   13
      Left            =   7240
      TabIndex        =   37
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label lblColor 
      Caption         =   "�߿�3"
      Height          =   255
      Index           =   12
      Left            =   5920
      TabIndex        =   36
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label lblColor 
      Caption         =   "�߿�1"
      Height          =   255
      Index           =   10
      Left            =   3280
      TabIndex        =   34
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label lblColor 
      Caption         =   "�����ɫ"
      Height          =   255
      Index           =   9
      Left            =   6840
      TabIndex        =   33
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblColor 
      Caption         =   "������ɫ"
      Height          =   255
      Index           =   8
      Left            =   5280
      TabIndex        =   32
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label LblBg_P 
      Caption         =   "ͼƬY����"
      Height          =   255
      Left            =   2880
      TabIndex        =   31
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Caption         =   "00"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   8
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblH 
      Caption         =   "�߶�"
      Height          =   255
      Left            =   6840
      TabIndex        =   27
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblW 
      Caption         =   "���"
      Height          =   255
      Left            =   5040
      TabIndex        =   26
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblNum 
      Caption         =   "��ʾ��Ŀ"
      Height          =   255
      Left            =   2880
      TabIndex        =   25
      Top             =   840
      Width           =   975
   End
   Begin VB.Label LblGBKMAP 
      Caption         =   "GBK������ַ"
      Height          =   255
      Left            =   5040
      TabIndex        =   24
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblColor 
      Caption         =   "������ɫ"
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   28
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Menu mFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mNew 
         Caption         =   "�½�(&N)"
      End
      Begin VB.Menu mOpen 
         Caption         =   "��(&O)"
      End
      Begin VB.Menu mSave 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu maSave 
         Caption         =   "���Ϊ(&A)..."
      End
      Begin VB.Menu mf1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mSet 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu mAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mTV 
      Caption         =   "TV�ؼ�"
      Visible         =   0   'False
      Begin VB.Menu mAdd 
         Caption         =   "���(&A)"
         Begin VB.Menu mFront 
            Caption         =   "��ӵ�ǰ��(&F)"
         End
         Begin VB.Menu mBehind 
            Caption         =   "��ӵ�����(&B)"
         End
         Begin VB.Menu mSub 
            Caption         =   "����Ӽ�(&S)"
         End
      End
      Begin VB.Menu mUp 
         Caption         =   "����(&U)"
      End
      Begin VB.Menu mDown 
         Caption         =   "����(&D)"
      End
      Begin VB.Menu mDele 
         Caption         =   "ɾ��(&E)"
      End
   End
   Begin VB.Menu mImage 
      Caption         =   "ͼƬ�ؼ�"
      Visible         =   0   'False
      Begin VB.Menu mEdit 
         Caption         =   "�༭"
         Begin VB.Menu mEditPath 
            Caption         =   "-"
            Index           =   0
         End
      End
      Begin VB.Menu mAddIm 
         Caption         =   "���ͼƬ"
      End
      Begin VB.Menu mDeleIm 
         Caption         =   "ɾ��ͼƬ"
      End
      Begin VB.Menu mOpenIm 
         Caption         =   "��ͼƬĿ¼"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()

'Ĭ�Ͽؼ�״̬
    apply_Picture
    picIcons.Visible = False
    frmMain.Wallpaper.ZOrder 0
    txtT.Tag = "00"
    Frame1.Enabled = False
    cmdRes.Enabled = False
    
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO00", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO01", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO02", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO03", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO04", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO05", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO06", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO07", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO08", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO09", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO0A", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO0B", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO0C", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO0D", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO0E", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO0F", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO10", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO11", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO12", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO13", App.Path & "\Config.ini")
    cbbfontNO.AddItem GetINI("lng", "cbbfontNO14", App.Path & "\Config.ini")


'���GIF�༭���˵�
    Dim j As Variant, i As Byte
    j = Split(GetINI("Setting", "ImageEdit", App.Path & "\Config.ini"), "|")
    For i = 1 To CByte(j(0))
        Load frmMain.mEditPath(frmMain.mEditPath.UBound + 1) '�Ӳ˵���+1
        frmMain.mEditPath(i).Caption = j(i)
    Next
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtT.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub



'****************************************************************


'�˵�
Private Sub mNew_Click()
        Dim fs, f
        Set fs = CreateObject("Scripting.FileSystemObject")
        '����SM_BG_PIC.GIF������Ŀ¼��
        If Len(Dir(Replace(App.Path & "\SM_BG_PIC.GIF", "\\", "\"))) <> 0 Then fs.DeleteFile Replace(App.Path & "\SM_BG_PIC.GIF", "\\", "\"), False
        Set f = fs.GetFile(Replace(App.Path & "\Default\SM_BG_PIC.GIF", "\\", "\"))
        f.Copy Replace(App.Path & "\SM_BG_PIC.GIF", "\\", "\")
        '��ʱͼ���ļ���
        On Error Resume Next
        fs.DeleteFolder Replace(App.Path & "\Temp", "\\", "\"), False
        Set f = fs.GetFolder(Replace(App.Path & "\Default\icons\", "\\", "\"))
        f.Copy (Replace(App.Path & "\Temp", "\\", "\"))
    If Open3(Replace(App.Path & "\Default\", "\\", "\"), "") = False Then Frame1.Enabled = False: Exit Sub
    Frame1.Enabled = True: cmdRes.Enabled = True
    SavePath = ""
End Sub
Private Sub mOpen_Click()
    CommonDialog1.Filter = GetINI("lng", "munOpen_CF", App.Path & "\Config.ini")
    CommonDialog1.CancelError = True
    On Error Resume Next
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
    Dim qwe As Variant
    qwe = Split(CommonDialog1.FileTitle, ".")
    If Open3(Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle)), CStr(qwe(0))) = False Then: Frame1.Enabled = False: Exit Sub

    Frame1.Enabled = True: cmdRes.Enabled = True
    SaveName = qwe(0)
    SavePath = Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle))
End Sub
Private Sub mSave_Click()
    If SavePath = "" Then Call maSave_Click: Exit Sub
    If Save3(SavePath, SaveName) = True Then MsgBox GetINI("lng", "SaveOK_MG", App.Path & "\Config.ini"), 1, GetINI("lng", "title_MG", App.Path & "\Config.ini")
End Sub
Private Sub maSave_Click()
    CommonDialog1.Filter = GetINI("lng", "munSaveAs_CF", App.Path & "\Config.ini")
    CommonDialog1.CancelError = True
    On Error Resume Next
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
    Dim qwe As Variant
    qwe = Split(CommonDialog1.FileTitle, ".")
    If Save3(Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle)), CStr(qwe(0))) = False Then Exit Sub
    '�ƶ�ͼ��Ŀ¼
        Dim fs, f
        Dim i As String
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.GetFolder(Replace(IIf(SavePath = "", App.Path & "\Temp\", SavePath & "icons\"), "\\", "\"))
        f.Copy Mid(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle))
        i = IIf(SavePath = "", "\Temp", "\icons")
    '������ͼ��Ŀ¼
    If Len(Dir(Replace(SavePath & "\icons\", "\\", "\"))) <> 0 Then fs.DeleteFolder Replace(SavePath & "\icons", "\\", "\"), False
    Name Replace(Mid(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle)) & i, "\\", "\") As Replace(Mid(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle)) & "\icons", "\\", "\")
    If Len(Dir(Replace(SavePath & "\icons\", "\\", "\"))) = 0 Then Exit Sub
    
    element(0).image = Replace(SavePath & "\icons\", "\\", "\")
    SaveName = qwe(0)
    SavePath = Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle))
    MsgBox GetINI("lng", "SaveOK_MG", App.Path & "\Config.ini"), 1, GetINI("lng", "title_MG", App.Path & "\Config.ini")
End Sub
Private Sub mExit_Click()
    If MsgBox(GetINI("lng", "mExit_SavePrompt_MG", App.Path & "\Config.ini"), _
    vbOKCancel, GetINI("lng", "title_MG", App.Path & "\Config.ini")) = 1 Then Call mSave_Click
    End
End Sub
Private Sub mSet_Click()
    frmMain.Enabled = False
    frmSetting.Show
End Sub
Private Sub mAbout_Click()
    frmMain.Enabled = False
    frmAbout.Show
End Sub


Private Sub mUp_Click() '��
    Dim ilast As String
    ilast = TreeView1.SelectedItem.Previous.Key
    TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
    TreeView1.Nodes.Add ilast, 3, element(nownum).name, element(nownum).name, element(nownum).image
    TreeView1.SelectedItem = TreeView1.Nodes(ilast).Previous
    Call iPrint(TreeView1.Nodes(ilast).Previous.Key, False, True)
End Sub
Private Sub mDown_Click() '��
    Dim inext As String
    inext = TreeView1.SelectedItem.Next.Key
    TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
    TreeView1.Nodes.Add inext, 2, element(nownum).name, element(nownum).name, element(nownum).image
    TreeView1.SelectedItem = TreeView1.Nodes(inext).Next
    Call iPrint(TreeView1.Nodes(inext).Next.Key, False, True)
End Sub
Private Sub mdele_Click() 'ɾ��
    TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
    element(nownum).name = element(UBound(element)).name
    element(nownum).image = element(UBound(element)).image
    element(nownum).Path_code = element(UBound(element)).Path_code
    element(nownum).TF = element(UBound(element)).TF
    ReDim Preserve element(UBound(element) - 1) As one_element
    nownum = 1
End Sub
 '���
Private Sub mFront_Click() 'ǰ��
    mAddSub (3)
End Sub
Private Sub mBehind_Click() '����
    mAddSub (2)
End Sub
Private Sub mSub_Click() '�Ӽ�
    mAddSub (4)
End Sub
Private Sub mAddSub(ty As Byte)  '���(����)
'������ƣ����ó�������ʾ��Ŀ��
    If ty <> 4 Then '�������ͬ��
        Do Until frmMain.TreeView1.Nodes(nownum).Previous Is Nothing
            nownum = Get_Index(frmMain.TreeView1.Nodes(nownum).Previous)
        Loop
        Dim mun As Byte
        Do Until frmMain.TreeView1.Nodes(nownum).Next Is Nothing
            nownum = Get_Index(frmMain.TreeView1.Nodes(nownum).Next)
            mun = mun + 1
        Loop
        If mun + 1 >= CInt(txtNum.Text) Then MsgBox GetINI("lng", "AddElement_MG", App.Path & "\Config.ini"), 1, GetINI("lng", "warn_MG", App.Path & "\Config.ini"): Exit Sub
    Else '��������Ӽ�
        If frmMain.TreeView1.Nodes(nownum).Children >= CInt(txtNum.Text) Then MsgBox GetINI("lng", "AddElement_MG", App.Path & "\Config.ini"), 1, GetINI("lng", "warn_MG", App.Path & "\Config.ini"): Exit Sub
    End If
    TreeView1.Nodes.Add TreeView1.SelectedItem, ty, "key", "name", 1
ReDim Preserve element(UBound(element) + 1)
    nownum = UBound(element)
    element(nownum).image = element(1).image
    element(nownum).name = "key"
    element(nownum).Path_code = "0"
Call TreeView1_NodeClick(TreeView1.Nodes.Item(nownum))
End Sub
'�Ҽ����༭ͼƬ
Private Sub mEditPath_Click(Index As Integer)
    ChDrive Left(element(0).image, 1)
    ChDir element(0).image
    On Error Resume Next
    Shell GetINI("setting", "ImageEditPath" & Index, App.Path & "\Config.ini") & " " & element(nownum).image, 4
    If Err.Number = 53 Then MsgBox GetINI("lng", "mEditPath_MG", App.Path & "\Config.ini"), 1, GetINI("lng", "title_MG", App.Path & "\Config.ini"): Exit Sub
End Sub
'�Ҽ������ͼƬ
Private Sub mAddIm_Click()
    CommonDialog1.Filter = GetINI("lng", "mAddIm_CF", App.Path & "\Config.ini")
    CommonDialog1.CancelError = True
    On Error Resume Next
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
        Dim fs, f
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.GetFile(CommonDialog1.FileName)
    '����ͼƬ��С
    Image1(Image1.Count - 1).Picture = LoadPicture(CommonDialog1.FileName)
    If Image1(Image1.Count - 1).Height > 20 And Image1(Image1.Count - 1).Width > 20 Then MsgBox GetINI("lng", "mAddIm_MG", App.Path & "\Config.ini"), 1, GetINI("lng", "title_MG", App.Path & "\Config.ini"): Exit Sub
        f.Copy IIf(SavePath = "", App.Path & "\Temp\", SavePath) & CommonDialog1.FileTitle
    ImageList1.ListImages.Add frmMain.ImageList1.ListImages.Count + 1, CommonDialog1.FileTitle, LoadPicture(CommonDialog1.FileName)
    imgIcons_Click
End Sub
'�Ҽ���ɾ��ͼƬ
Private Sub mDeleIm_Click()
    ImageList1.ListImages.Remove CInt(mDeleIm.Tag)
    imgIcons_Click
End Sub
'�Ҽ�����ͼƬĿ¼
Private Sub mOpenIm_Click()
    Dim sTmp As String * 200, Length As Long
    Length = GetWindowsDirectory(sTmp, 200)
    Shell Left(sTmp, Length) & "\explorer.exe " & Replace(IIf(SavePath = "", App.Path & "\Temp\", SavePath & "\icons\"), "\\", "\"), 4
End Sub

'ˢ��
Private Sub cmdRes_Click()
    If txtNum.Text = "" Or txtX.Text = "" Or txtY.Text = "" Then _
    MsgBox GetINI("lng", "cmdRes_noNumXY_MG", App.Path & "\Config.ini"), vbOKOnly, _
    GetINI("lng", "warn_MG", App.Path & "\Config.ini"): Exit Sub
    If TreeView1.Nodes.Count <> 0 Then iPrint element(nownum).name, True, True
End Sub


'****************************************************************


'��������
Private Sub txtNum_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0 '��������
'If KeyAscii = 8 Then Exit Sub
'If CInt(Mid(txtY.Text, 1, txtY.SelStart) & Chr(KeyAscii) & Mid(txtT.Text, txtT.SelStart + txtT.SelLength + 1)) > 10 Then KeyAscii = 0 '���ִ�С����
End Sub
Private Sub txtStart_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or (KeyAscii > 57 And KeyAscii < 65) Or (KeyAscii > 70 And KeyAscii < 97) Or KeyAscii > 102) And KeyAscii <> 120 And KeyAscii <> 88 And KeyAscii <> 8 Then KeyAscii = 0 '���Ƴ����֡�ABCDEFabcdef���롢Xx
If (KeyAscii = 88 Or KeyAscii = 120) And txtOffset.SelStart <> 1 Then KeyAscii = 0
End Sub
Private Sub txtOffset_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or (KeyAscii > 57 And KeyAscii < 65) Or (KeyAscii > 70 And KeyAscii < 97) Or KeyAscii > 102) And KeyAscii <> 8 Then KeyAscii = 0
If txtOffset.SelStart < 2 Or (txtOffset.SelStart = 2 And txtOffset.SelLength = 0 And KeyAscii = 8) Then KeyAscii = 0
End Sub
Private Sub txtX_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0 '��������
If KeyAscii = 8 Then Exit Sub
If CInt(Mid(txtX.Text, 1, txtX.SelStart) & Chr(KeyAscii) & Mid(txtT.Text, txtT.SelStart + txtT.SelLength + 1)) > 176 Then KeyAscii = 0 '���ִ�С����
End Sub
Private Sub txtY_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0 '��������
If KeyAscii = 8 Then Exit Sub
If CInt(Mid(txtY.Text, 1, txtY.SelStart) & Chr(KeyAscii) & Mid(txtT.Text, txtT.SelStart + txtT.SelLength + 1)) > 220 Then KeyAscii = 0 '���ִ�С����
End Sub
Private Sub txtPath_KeyPress(KeyAscii As Integer)
If Path_code(0).Value = True Then '����
    If (KeyAscii < 48 Or (KeyAscii > 57 And KeyAscii < 65) Or (KeyAscii > 70 And KeyAscii < 97) Or KeyAscii > 102) And KeyAscii <> 8 Then KeyAscii = 0 '���Ƴ����֡�ABCDEFabcdef����
    If txtPath.SelStart < 2 Or (txtPath.SelStart = 2 And txtPath.SelLength = 0 And KeyAscii = 8) Then KeyAscii = 0 '���Ƶ�����С��2��ʱ
Else '·��
    'If txtPath.SelStart < 1 And KeyAscii <> 33 Or (txtPath.SelStart = 1 And txtPath.SelLength = 0 And KeyAscii = 8) Then KeyAscii = 0 '���Ƶ�����С��2��ʱ
End If
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
    If txtName.SelStart = 0 And KeyAscii > 47 And KeyAscii < 58 Then KeyAscii = 0 '��һ���ַ�����������
End Sub
Private Sub txtT_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0 '��������
    If KeyAscii = 8 Then Exit Sub
    If CInt(Mid(txtT.Text, 1, txtT.SelStart) & Chr(KeyAscii) & Mid(txtT.Text, txtT.SelStart + txtT.SelLength + 1)) > 256 Then KeyAscii = 0 '���ִ�С����
End Sub


'****************************************************************


'��ʱ�ı�
Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    element(nownum).name = NewString
    txtName.Text = element(nownum).name
    TreeView1.Nodes.Item(nownum).Key = element(nownum).name
    Call iPrint(element(nownum).name, False, True) 'ˢ��ͼ��
End Sub
Private Sub LblColor_Click(Index As Integer)
    If Index > 6 Then Exit Sub
    If iButton(0) = 2 Then
        txtT.Left = iButton(1) + lblColor(Index).Left
        txtT.Top = iButton(2) + lblColor(Index).Top
        txtT.Visible = True
        txtT.SetFocus '�õ����
    Else
        On Error Resume Next
        CommonDialog1.ShowColor
        If Err.Number = 32755 Then Exit Sub
        lblColor(Index).BackColor = CommonDialog1.Color
        lblColor(Index).ForeColor = 16777215 - CommonDialog1.Color '��ɫ��ʾ
        Call iPrint(element(nownum).name, True, True) 'ˢ�±߿�ͼ��
    End If
End Sub
Private Sub txtName_Change()
    If element(nownum).name = txtName.Text Then Exit Sub
    element(nownum).name = txtName.Text
    TreeView1.Nodes.Item(nownum).Text = element(nownum).name
    TreeView1.Nodes.Item(nownum).Key = element(nownum).name
    Call iPrint(element(nownum).name, False, True) 'ˢ������
End Sub
Private Sub Path_code_Click(Index As Integer)
    If Index = IIf(Left(element(nownum).Path_code, 2) = "0x", 0, 1) Then Exit Sub '���ѡ�����ͬ�����˳�
    If Index = 0 Then
        txtPath.Text = "0x0"
    Else
        txtPath.Text = "/b/ELF/"
    End If
End Sub
Private Sub txtPath_Change()
    If element(nownum).Path_code = txtPath.Text Then Exit Sub
    element(nownum).Path_code = IIf(txtPath.Text = "0x", "0x0", txtPath.Text)
End Sub

Private Sub LblBg_P_Click()
    txtPict.Tag = IIf(txtPict.Enabled, txtPict.Text, txtPict.Tag) '���Ҫ�رգ�����ֵ����TAG
    txtPict.Text = IIf(txtPict.Enabled, 0, txtPict.Tag) '���Ҫ�رգ���0д��TXT
    txtPict.Enabled = IIf(txtPict.Enabled, False, True) '���
End Sub
Private Sub LblEorC_Click()
    If txtPict.Enabled = linzi Then LblBg_P_Click
    LblGBKMAP.Visible = IIf(linzi, False, True)
    txtOffset.Visible = IIf(linzi, False, True)
    LblBg_P.Visible = IIf(linzi, False, True)
    txtPict.Visible = IIf(linzi, False, True)
    LblEorC.Caption = GetINI("lng", IIf(linzi, "LblEorC-E", "LblEorC-C"), App.Path & "\Config.ini")
    linzi = IIf(linzi, False, True)
End Sub

'�����ַ�����������TXTBOX��
Private Sub txtPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picIcons.Visible = False
    If Len(txtPath.Text) < 26 Then Exit Sub
    txtPath.Left = 120
    txtPath.Width = 4817
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtPath.Left = 2760
    txtPath.Width = 2415
    picIcons.Visible = False
End Sub


'****************************************************************

'��ɫ͸����
Private Sub lblColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index > 6 Then Exit Sub
    If Button = 2 Then '��ʾΪ�Ҽ�
        iButton(0) = 2: iButton(1) = X: iButton(2) = Y '��¼XY����
        txtT.Tag = Index '��¼��ӦIndex
        txtT.Text = CInt("&H" & lblColor(Index).Caption) '��ʾΪ16����
    ElseIf Button = 1 Then
        iButton(0) = 1 '��ʾΪ���
    End If
End Sub

Private Sub txtT_Change()
    If txtT.Text <> "" Then lblColor(txtT.Tag).Caption = Hex(txtT.Text)
    If Len(lblColor(txtT.Tag).Caption) = 1 Then lblColor(txtT.Tag).Caption = "0" & lblColor(txtT.Tag).Caption
End Sub

'****************************************************************


'ѡ��Ԫ��&�Ҽ��˵�
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
If iButton(0) = 1 Then
    iButton(0) = 0
    'Ԥ��
    Dim a As Boolean
    With frmMain.TreeView1
        If Left(.Nodes(element(nownum).name).FullPath, _
        Len(.Nodes(element(nownum).name).FullPath) - Len(.Nodes(element(nownum).name).Text)) _
        <> Left(.Nodes(element(lastnum).name).FullPath, _
        Len(.Nodes(element(lastnum).name).FullPath) - Len(.Nodes(element(lastnum).name).Text)) _
        Then a = True
    End With
    Call iPrint(Node.Key, False, a)
ElseIf iButton(0) = 2 Then '�Ҽ��˵�
    iButton(0) = 1
    mUp.Enabled = True: mDown.Enabled = True
    If Node.Previous Is Nothing Then mUp.Enabled = False
    If Node.Next Is Nothing Then mDown.Enabled = False
    If Node.Children Then mDown.Enabled = False: mUp.Enabled = False
    frmMain.PopupMenu mTV, vbPopupMenuLeftAlign, iButton(1) + TreeView1.Left + Frame1.Left, iButton(2) + TreeView1.Top + Frame1.Top
End If
End Sub

'2-׼���Ҽ��˵�,1-ѡ��
Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    iButton(0) = 2: iButton(1) = X: iButton(2) = Y '��ʾΪ�Ҽ�
ElseIf Button = 1 Then
    iButton(0) = 1 '��ʾΪ���
End If
'��ʾԪ������
If Not (TreeView1.HitTest(X, Y) Is Nothing) Then
    lastnum = nownum '��һ��
    nownum = Get_Index(TreeView1.HitTest(X, Y).Key) 'ȡ��������
    txtName.Text = element(nownum).name
    If TreeView1.HitTest(X, Y).Children = 0 Then '�����Ӽ�
        Path_code(IIf(Len(element(nownum).Path_code) > 5, 1, 0)).Value = 1
        txtPath.Text = element(nownum).Path_code
        Path_code(0).Enabled = True
        Path_code(1).Enabled = True
        txtPath.Enabled = True
    Else
        element(nownum).TF = True '���Ӽ�
        Path_code(0).Value = True
        txtPath.Text = "0" '�޸�element(nownum).Path_code
        Path_code(0).Enabled = False
        Path_code(1).Enabled = False
        txtPath.Enabled = False
    End If
    imgIcons.Picture = LoadPicture(element(0).image & element(nownum).image)
End If
End Sub
'����ƶ�-�϶�
Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And Not (TreeView1.HitTest(X, Y) Is Nothing) Then 'ָʾһ���϶�������
        TreeView1.SelectedItem = TreeView1.Nodes(element(nownum).name)
        'ʹ��CreateDragImage���������϶�ͼ�ꡣ
        TreeView1.DragIcon = TreeView1.SelectedItem.CreateDragImage
        TreeView1.Drag vbBeginDrag '�϶�������
    End If
    txtPath.Left = 2760
    txtPath.Width = 2415
    picIcons.Visible = False
End Sub
Private Sub TreeView1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    TreeView1.DropHighlight = TreeView1.HitTest(X, Y) '����
    If Not (TreeView1.DropHighlight Is Nothing) Then '���۵���
        TreeView1.DropHighlight.Expanded = True
    End If
End Sub
'�ƶ����������Ӽ�
Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
If TreeView1.DropHighlight Is Nothing Then Exit Sub
If TreeView1.DropHighlight.Key <> TreeView1.SelectedItem Then
    Dim Highlight As String
    Highlight = TreeView1.DropHighlight.Key
    TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
    TreeView1.Nodes.Add Highlight, 4, element(nownum).name, element(nownum).name, element(nownum).image
End If
    Set TreeView1.DropHighlight = Nothing
    TreeView1.SelectedItem = TreeView1.Nodes(element(nownum).name) '������
    Call iPrint(element(nownum).name, False, True) 'ˢ��ͼ��
End Sub


'****************************************************************


'�������ʾ����ͼƬ �Ҽ����༭ͼƬ
Private Sub imgIcons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    picIcons.Left = X + imgIcons.Left
    picIcons.Top = Y + imgIcons.Top
    imgIcons_Click
ElseIf Button = vbRightButton Then
    mEdit.Visible = True
    mAddIm.Visible = False
    mDeleIm.Visible = False
    PopupMenu mImage, vbPopupMenuLeftAlign, X + imgIcons.Left + Frame1.Left, X + Frame1.Top + imgIcons.Top
End If
End Sub
'��ʾ����ͼƬ
Private Sub imgIcons_Click()
Dim i As Integer
For i = 1 To Image1.Count - 1
    Unload Image1(i)
Next
For i = 0 To ImageList1.ListImages.Count - 1
    Load Image1(Image1.UBound + 1)
    Image1(i).Left = 17 * (i Mod 8)
    Image1(i).Top = 18 * (i \ 8)
    Image1(i).Picture = ImageList1.ListImages(i + 1).Picture
    Image1(i).ToolTipText = ImageList1.ListImages(i + 1).Key & " " & Image1(i).Width & "x" & Image1(i).Height
    Image1(i).Visible = True
Next
picIcons.Height = Screen.TwipsPerPixelY * 19 * (i \ 8 + 1)
picIcons.Visible = True '��ʾ����ͼƬ
End Sub
'�ı䵱ǰͼ��
Private Sub Image1_Click(Index As Integer)
    element(nownum).image = ImageList1.ListImages(Index + 1).Key
    imgIcons.Picture = LoadPicture(element(0).image & element(nownum).image)
    TreeView1.Nodes(element(nownum).name).image = element(nownum).image
    picIcons.Visible = False
End Sub
Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mDeleIm.Visible = True
    mAddIm.Visible = False
    mEdit.Visible = False
    mDeleIm.Tag = Index + 1
    PopupMenu mImage, vbPopupMenuLeftAlign, X + Image1(Index).Left * Screen.TwipsPerPixelX + picIcons.Left + Frame1.Left, Y + Image1(Index).Top * Screen.TwipsPerPixelY + Frame1.Top + picIcons.Top
End Sub

'�Ҽ������ͼƬ������
Private Sub picIcons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mAddIm.Visible = True
    mEdit.Visible = False
    mDeleIm.Visible = False
    PopupMenu mImage, vbPopupMenuLeftAlign, X * Screen.TwipsPerPixelX + picIcons.Left + Frame1.Left, Y * Screen.TwipsPerPixelY + Frame1.Top + picIcons.Top
End Sub
