Attribute VB_Name = "Module1"
Option Explicit
Public element() As one_element 'Ԫ��
Public nownum As Integer 'ĿǰԪ��NUM
Public lastnum As Integer 'ǰһ��Ԫ��NUM
Public linzi As Boolean '�������İ�
Public Bg_P As Boolean '����ͼƬ(����2.6��)


'Ԫ��
Public Type one_element
    name As String
    TF As Boolean
    Path_code As String
    image As String
End Type


Public SavePath As String '����·��
Public SaveName As String '��������

Public iButton(2) As Integer '���ز˵�


'Public PictureFT As Byte, PicturePath As Byte
'Public WallpaperPath$, IconPath$

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'��ȡWinsowsϵͳ�ļ���

Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal _
        hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) _
        As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc _
        As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
        ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
        lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
        ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc _
        As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush _
        As LOGBRUSH) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Const SRCPAINT = &HEE0086
Const SRCAND = &H8800C6
Const BS_SOLID = 0
Const gColor = &HFFFFFF
Public Sub Main()
    frmMain.Show
    'ȡ�������в���
    If Len(Command()) <> 0 Then
        Dim iName As String '·��
        Dim qwe As Variant, qwe2 As Variant
        iName = Replace(Command(), Chr(34), "") '�滻"Ϊ��
        qwe = Split(iName, "\")
        qwe2 = Split(qwe(UBound(qwe)), ".")
        If Open3(Left(iName, Len(iName) - Len(qwe(UBound(qwe)))), CStr(qwe(0))) = False Then Exit Sub
        SaveName = qwe2(0)
        SavePath = Left(iName, Len(iName) - Len(qwe(UBound(qwe))))
        frmMain.mSave.Enabled = True
        frmMain.maSave.Enabled = True
    End If
frmMain.Caption = GetINI("lng", "MainCaption", App.Path & "\Config.ini")
frmMain.mFile.Caption = GetINI("lng", "mFile", App.Path & "\Config.ini")
frmMain.mNew.Caption = GetINI("lng", "mNew", App.Path & "\Config.ini")

frmMain.mOpen.Caption = GetINI("lng", "mOpen", App.Path & "\Config.ini")
frmMain.mSave.Caption = GetINI("lng", "mSave", App.Path & "\Config.ini")
frmMain.maSave.Caption = GetINI("lng", "maSave", App.Path & "\Config.ini")
frmMain.mExit.Caption = GetINI("lng", "mExit", App.Path & "\Config.ini")
frmMain.mHelp.Caption = GetINI("lng", "mHelp", App.Path & "\Config.ini")
frmMain.mSet.Caption = GetINI("lng", "mSet", App.Path & "\Config.ini")
frmMain.mAbout.Caption = GetINI("lng", "mAbout", App.Path & "\Config.ini")
frmMain.mAdd.Caption = GetINI("lng", "mAdd", App.Path & "\Config.ini")
frmMain.mFront.Caption = GetINI("lng", "mFront", App.Path & "\Config.ini")
frmMain.mBehind.Caption = GetINI("lng", "mBehind", App.Path & "\Config.ini")
frmMain.mSub.Caption = GetINI("lng", "mSub", App.Path & "\Config.ini")
frmMain.mDown.Caption = GetINI("lng", "mDown", App.Path & "\Config.ini")
frmMain.mDele.Caption = GetINI("lng", "mDele", App.Path & "\Config.ini")

frmMain.cmdRes.Caption = GetINI("lng", "cmdRes", App.Path & "\Config.ini")
frmMain.Frame1.Caption = GetINI("lng", "Frame1", App.Path & "\Config.ini")
frmMain.LblName.Caption = GetINI("lng", "name", App.Path & "\Config.ini")
frmMain.Path_code(0).Caption = GetINI("lng", "Path", App.Path & "\Config.ini")
frmMain.Path_code(1).Caption = GetINI("lng", "App", App.Path & "\Config.ini")

frmMain.LblRegister.Caption = GetINI("lng", "Register", App.Path & "\Config.ini")
frmMain.LblGBKMAP.Caption = GetINI("lng", "GBKMAP", App.Path & "\Config.ini")
frmMain.lblNum.Caption = GetINI("lng", "items_num", App.Path & "\Config.ini")
frmMain.lblW.Caption = GetINI("lng", "Item_width", App.Path & "\Config.ini")
frmMain.lblH.Caption = GetINI("lng", "Item_height", App.Path & "\Config.ini")
frmMain.LblBg_P.Caption = GetINI("lng", "Bg_Picture_YPosition", App.Path & "\Config.ini")
frmMain.LblBg_P.ToolTipText = GetINI("lng", "Lab_Bg_Picture_Tip", App.Path & "\Config.ini")
frmMain.lblColor(7).Caption = GetINI("lng", "BackgroundColor", App.Path & "\Config.ini")
frmMain.lblColor(8).Caption = GetINI("lng", "TextColor", App.Path & "\Config.ini")
frmMain.lblColor(9).Caption = GetINI("lng", "CusorColor", App.Path & "\Config.ini")
frmMain.lblColor(10).Caption = GetINI("lng", "Line1Color", App.Path & "\Config.ini")
frmMain.lblColor(11).Caption = GetINI("lng", "Line2Color", App.Path & "\Config.ini")
frmMain.lblColor(12).Caption = GetINI("lng", "Line3Color", App.Path & "\Config.ini")
frmMain.lblColor(13).Caption = GetINI("lng", "Line4Color", App.Path & "\Config.ini")
frmMain.cbbfontNO.Text = GetINI("lng", "Font", App.Path & "\Config.ini")


frmSetting.Caption = GetINI("lng", "SettingCaption", App.Path & "\Config.ini")
frmSetting.chkPicture.Caption = GetINI("lng", "chkPicture", App.Path & "\Config.ini")
frmSetting.fraPicture.Caption = GetINI("lng", "fraPicture", App.Path & "\Config.ini")
frmSetting.optPictureSuff_Other(0).Caption = GetINI("lng", "optPictureSuff", App.Path & "\Config.ini")
frmSetting.optPictureSuff_Other(1).Caption = GetINI("lng", "optPictureOther", App.Path & "\Config.ini")
frmSetting.Label1.Caption = GetINI("lng", "Label1", App.Path & "\Config.ini")
frmSetting.Label2.Caption = GetINI("lng", "Label2", App.Path & "\Config.ini")
frmSetting.Label3.Caption = GetINI("lng", "Label3", App.Path & "\Config.ini")
frmSetting.cmdWallpaper.Caption = GetINI("lng", "cmdWallpaper", App.Path & "\Config.ini")
frmSetting.cmdIcon.Caption = GetINI("lng", "cmdIcon", App.Path & "\Config.ini")
frmSetting.cmdSave.Caption = GetINI("lng", "cmdSave", App.Path & "\Config.ini")
frmSetting.cmdapply.Caption = GetINI("lng", "cmdapply", App.Path & "\Config.ini")
End Sub

'���ļ�
Public Function Open3(Path As String, name As String) As Boolean
Dim one(1) As Byte '�б���
Dim i As String '�б���
Dim qwe As Variant 'Ԫ�ر���
Dim relative As String '������
Dim truePath As String '��ȷ��·��

ReDim element(0) As one_element
frmMain.ImageList1.ListImages.Clear
frmMain.TreeView1.Nodes.Clear

'���ͼƬĿ¼
If Len(Dir(Replace(Path & "\icons\", "\\", "\"))) = 0 Then MsgBox GetINI("lng", "Open_MG", App.Path & "\Config.ini"), 1, GetINI("lng", "warn_MG", App.Path & "\Config.ini"): Open3 = False: Exit Function
If name = "" Then name = "Start_Menu"
nownum = 1
element(0).image = Replace(IIf(SavePath = "", App.Path & "\Temp\", SavePath & "\icons\"), "\\", "\")

'INI�ļ�
linzi = False
truePath = Replace(Path & "\" & name & ".ini", "\\", "\")
If Get_TruePath(truePath, "*.ini") = False Then Open3 = False: Exit Function
Open truePath For Binary As #1
    Line Input #1, i
        frmMain.txtStart.Text = Mid(i, 10)
        If Not (EOF(1)) Then
            Line Input #1, i
            If i <> "" Then frmMain.txtOffset.Text = Mid(i, 8): frmMain.txtOffset.Visible = True: frmMain.LblGBKMAP.Visible = True: linzi = True
        End If
Close #1

'SKN�ļ�
truePath = Replace(Path & "\" & name & ".skn", "\\", "\")
If Get_TruePath(truePath, "*.skn") = False Then Open3 = False: Exit Function

frmMain.txtX.Text = GetINI("Start_menu_Skin", "Item_width", truePath)
frmMain.txtY.Text = GetINI("Start_menu_Skin", "Item_height", truePath)
frmMain.txtPict.Text = GetINI("Start_menu_Skin", "Bg_Picture_YPosition", truePath)
    If linzi = False And frmMain.txtPict.Text <> "" Then
        If MsgBox(GetINI("lng", "OpenSKN_MG", App.Path & "\Config.ini"), vbOKCancel, GetINI("lng", "warn_MG", App.Path & "\Config.ini")) = vbCancel Then Open3 = False: Exit Function
    End If
    linzi = IIf(frmMain.txtPict.Text = "", False, True) '��BG_P�������İ�
    frmMain.txtPict.Enabled = IIf(frmMain.txtPict.Text = "", False, True)
    frmMain.txtPict.Text = IIf(frmMain.txtPict.Text = "", "0", frmMain.txtPict.Text)
frmMain.txtNum.Text = GetINI("Start_menu_Skin", "items_num", truePath)
frmMain.cbbfontNO.ListIndex = GetINI("Start_menu_Skin", "Font", truePath)

Dim s(6) As String
s(0) = "BackgroundColor": s(1) = "TextColor": s(6) = "CusorColor"
s(2) = "Line1Color": s(3) = "Line2Color": s(4) = "Line3Color": s(5) = "Line4Color"

For one(0) = 0 To 6
    i = GetINI("Start_menu_Skin", s(one(0)), truePath)
    frmMain.lblColor(one(0)).BackColor = RBGT2TBGR(i, True)
    frmMain.lblColor(one(0)).Caption = Mid(i, 9, 2)
Next

'menu�ļ�
truePath = Replace(Path & "\" & name & ".menu", "\\", "\")
If Get_TruePath(truePath, "*.menu") = False Then Open3 = False: Exit Function
Open truePath For Binary As #1

    Do Until EOF(1)
    '��ȡһ�У���i
        i = ""
        Do
            Get #1, , one(0)
            If one(0) = 13 Then
                Get #1, , one(1)
            ElseIf one(0) > 128 Then
                Get #1, , one(1)
                i = i & Chr(CLng(one(0)) * 256 + one(1))
            ElseIf one(0) <> 10 And Not (EOF(1)) Then
                i = i & Chr(one(0))
            ElseIf one(0) = 0 And EOF(1) Then
                Exit Do
            End If
        Loop Until one(0) = 13 And one(1) = 10
    '�ж��Ƿ���Ԫ��
        'ȡ�ø����� relative
        If Mid(i, 1, 8) = "#/Start/" Then
            If Len(i) > 8 Then
                qwe = Split(i, "/")
                relative = qwe(UBound(qwe) - 1)
            End If
        '�ж��Ƿ�Ϊlinzi���İ�
        ElseIf Mid(i, 1, 8) = "#/End/" Then
            If linzi = False Then
                If MsgBox(GetINI("lng", "OpenMENU_MG", App.Path & "\Config.ini"), vbOKCancel, GetINI("lng", "warn_MG", App.Path & "\Config.ini")) = vbCancel Then Open3 = False: Exit Function
            End If
            linzi = True
        'Ԫ��
        ElseIf Mid(i, 1, 8) <> "#/End/" And Len(i) <> 0 Then
            ReDim Preserve element(UBound(element) + 1)
            'ȡ��Ԫ������·�����롢ͼƬ���������Ӽ�
            qwe = Split(i, ";")
            element(nownum).Path_code = qwe(1)
            element(nownum).image = qwe(2)
            qwe = Split(qwe(0), "=")
            element(nownum).TF = CBool(qwe(1))
            element(nownum).name = qwe(0)
           '��ӵ�ͼƬ���б���ӽڵ�ؼ�
            If relative = "" Then
                If ANewListImages(element(nownum).image) = False Then frmMain.ImageList1.ListImages.Add frmMain.ImageList1.ListImages.Count + 1, element(nownum).image, LoadPicture(element(0).image & element(nownum).image)
                frmMain.TreeView1.Nodes.Add , , element(nownum).name, element(nownum).name, element(nownum).image
            Else
                If ANewListImages(element(nownum).image) = False Then frmMain.ImageList1.ListImages.Add frmMain.ImageList1.ListImages.Count + 1, element(nownum).image, LoadPicture(element(0).image & element(nownum).image)
                frmMain.TreeView1.Nodes.Add relative, 4, element(nownum).name, element(nownum).name, element(nownum).image
            End If
        
            nownum = nownum + 1
        End If
    Loop

Close #1

'���Ԥ��ͼ��ؼ�
    For one(0) = 1 To frmMain.Iicon.UBound 'ɾ��ȫ��
        Unload frmMain.Iicon(one(0))
    Next
    For one(0) = 1 To CInt(frmMain.txtNum.Text)
        Load frmMain.Iicon(frmMain.Iicon.UBound + 1)
        frmMain.Iicon(one(0)).Top = 220 - CInt(frmMain.txtNum.Text) * CInt(frmMain.txtY.Text) - 20 - 4 - 1 + (one(0) - 1) * CInt(frmMain.txtY.Text) + 2
        frmMain.Iicon(one(0)).Left = 6
        frmMain.Iicon(one(0)).Visible = True
    Next
'��ʼ��
    lastnum = 1
    nownum = 1
    frmMain.LblEorC.Caption = GetINI("lng", "LblEorC-C", App.Path & "\Config.ini")
    If linzi = False Then frmMain.LblGBKMAP.Visible = False: frmMain.txtOffset.Visible = False: frmMain.LblBg_P.Visible = False: frmMain.txtPict.Visible = False: frmMain.LblEorC.Caption = GetINI("lng", "LblEorC-E", App.Path & "\Config.ini")
'ˢ��Ԥ��
    Call iPrint(frmMain.TreeView1.Nodes.Item(1), True, True)
'���
    Open3 = True
End Function
'ȡ��ͼ���б����޴�name
Public Property Get ANewListImages(name As String) As Boolean
Dim i As Byte
For i = 1 To frmMain.ImageList1.ListImages.Count
    If frmMain.ImageList1.ListImages(i).Key = name Then ANewListImages = True: Exit Property
Next
ANewListImages = False
End Property



Public Sub apply_Picture()
Dim nowPath$
Dim qwe As Variant, i As Byte
If GetINI("Setting", "PictureFT", App.Path & "\Config.ini") = 1 Then
    If GetINI("Setting", "PictureSuff_Other", App.Path & "\Config.ini") = 0 Then
        nowPath = App.Path & "\Default\icon\"
        frmMain.Wallpaper.Picture = LoadPicture(LoadP(App.Path & "\Default\icon\Wallpaper.jpg", 3))
    Else
        nowPath = GetINI("Setting", "IconPath", App.Path & "\Config.ini")
        frmMain.Wallpaper.Picture = LoadPicture(LoadP(GetINI("Setting", "WallpaperPath", App.Path & "\Config.ini"), 3))
    End If
    
    qwe = Split(GetINI("Setting", "IconID", App.Path & "\Config.ini"), "|")
    For i = 0 To UBound(qwe) - 2
        frmMain.imgIcon(i).Picture = LoadPicture(LoadP(nowPath & "\" & qwe(i) & ".gif", 2))
    Next
    frmMain.imgIcon(8).Picture = LoadPicture(LoadP(nowPath & "\" & qwe(7) & ".gif", 2))
    frmMain.imgIcon(9).Picture = LoadPicture(LoadP(nowPath & "\" & qwe(8) & ".gif", 2))
End If
End Sub
Public Function LoadP(Path As String, iT As Byte) As String
    If Len(Dir(Replace(Path, "\\", "\"))) = 0 Then
        If iT = 1 Then
            Path = Replace(App.Path & "\Default\icon\Start_Menu.gif", "\\", "\")
        ElseIf iT = 2 Then
            Path = Replace(App.Path & "\Default\icon\Error.gif", "\\", "\")
        ElseIf iT = 3 Then
            Path = Replace(App.Path & "\Default\icon\Wallpaper.jpg", "\\", "\")
        ElseIf iT = 4 Then
            Dim fs, f
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(Replace(App.Path & "\Default\SM_BG_PIC.GIF", "\\", "\"))
            f.Copy Replace(SavePath & "\SM_BG_PIC.GIF", "\\", "\")
        End If
    End If
    LoadP = Path
End Function



Public Function Save3(Path As String, name As String) As Boolean
    If name = "" Then name = "Start_Menu"
'INI�ļ�
    Open Path & "\" & name & ".ini" For Output As #2
        Print #2, "Register=" & frmMain.txtStart.Text
        If linzi = True Then Print #2, "GBKMAP=" & frmMain.txtOffset.Text
    Close #2
'SKN�ļ�
WriteINI "Start_menu_Skin", "Item_width", frmMain.txtX.Text, Replace(Path & "\" & name & ".skn", "\\", "\")
WriteINI "Start_menu_Skin", "Item_height", frmMain.txtY.Text, Replace(Path & "\" & name & ".skn", "\\", "\")
If linzi = True And frmMain.txtPict.Enabled = True Then WriteINI "Start_menu_Skin", "Bg_Picture_YPosition", frmMain.txtPict.Text, Replace(Path & "\" & name & ".skn", "\\", "\")
Dim s(6) As String
s(0) = "BackgroundColor": s(1) = "TextColor": s(6) = "CusorColor"
s(2) = "Line1Color": s(3) = "Line2Color": s(4) = "Line3Color": s(5) = "Line4Color"
Dim a As Byte
For a = 0 To 6
    element(0).name = Hex(frmMain.lblColor(a).BackColor)
    element(0).Path_code = IIf(Len(Hex(frmMain.lblColor(a).Caption)) < 2, "0", "") & Hex(frmMain.lblColor(a).Caption)
    element(0).name = RBGT2TBGR(element(0).name, False) & element(0).Path_code
    WriteINI "Start_menu_Skin", s(a), element(0).name, Replace(Path & "\" & name & ".skn", "\\", "\")
Next

WriteINI "Start_menu_Skin", "items_num", frmMain.txtNum.Text, Replace(Path & "\" & name & ".skn", "\\", "\")
WriteINI "Start_menu_Skin", "Font", frmMain.cbbfontNO.ListIndex, Replace(Path & "\" & name & ".skn", "\\", "\")
'menu�ļ�
    Open Path & "\" & name & ".menu" For Output As #2
    '��ߡ���ǰ
        Do Until frmMain.TreeView1.Nodes(nownum).Parent Is Nothing
            nownum = Get_Index(frmMain.TreeView1.Nodes(nownum).Parent)
        Loop
        Do Until frmMain.TreeView1.Nodes(nownum).Previous Is Nothing
            nownum = Get_Index(frmMain.TreeView1.Nodes(nownum).Previous)
        Loop
        
        Dim Allmenu() As String, i As Integer, j As Integer
        Dim isubname() As String
    'ȡ��һ���ж��ٸ���
        For i = 1 To frmMain.TreeView1.Nodes.Count
            If frmMain.TreeView1.Nodes(i).Children Then j = j + 1
            element(Get_Index(frmMain.TreeView1.Nodes(i))).TF = True
        Next

        ReDim Allmenu(j) As String
        ReDim isubname(j) As String
        j = 1
    'ȡ��ÿ���·��
        For i = 1 To frmMain.TreeView1.Nodes.Count
            If frmMain.TreeView1.Nodes(i).Children Then
                isubname(j) = frmMain.TreeView1.Nodes(i).FullPath & "\"
                j = j + 1
            End If
        Next
    '��������Ⱥ�˳��
        
        
    'ȡ��ÿ����ַ���
        For j = 0 To UBound(Allmenu)
            'ȡ����������κ�һ��Ԫ��
            For i = 1 To frmMain.TreeView1.Nodes.Count
                If isubname(j) = Left(frmMain.TreeView1.Nodes(i).FullPath, Len(frmMain.TreeView1.Nodes(i).FullPath) - Len(frmMain.TreeView1.Nodes(i))) Then
                    nownum = i: Exit For
                End If
            Next
            'ȡ�����������ǰԪ��
            Do Until frmMain.TreeView1.Nodes(nownum).Previous Is Nothing '�ҵ�ǰ
                nownum = Get_Index(frmMain.TreeView1.Nodes(nownum).Previous)
            Loop
            'ȡ��������������ַ���
            Do Until frmMain.TreeView1.Nodes(nownum).Next Is Nothing
                Allmenu(j) = Allmenu(j) & element(nownum).name & "=" & IIf(element(nownum).TF, 1, 0) & ";" & element(nownum).Path_code & ";" & element(nownum).image & Chr(13) & Chr(10)
                nownum = Get_Index(frmMain.TreeView1.Nodes(nownum).Next)
            Loop
            Allmenu(j) = "#/Start/" & Replace(isubname(j), "\", "/") & Chr(13) & Chr(10) & Allmenu(j) & IIf(linzi = True, "#/End/" & Chr(13) & Chr(10), "")
        Next
    'д���ļ�
        For i = 0 To UBound(Allmenu)
            Print #2, Allmenu(i)
        Next
    Close #2
Save3 = True
End Function

Public Function RBGT2TBGR(RGBT As String, TF As Boolean) As String
'ϵͳ&HBBGGRR&
Dim i As Byte
If Len(RGBT) < 6 Then
    For i = 1 To 6 - Len(RGBT)
        RGBT = "0" & RGBT
    Next
End If
RGBT = IIf(TF, "", "&H") & RGBT
RBGT2TBGR = IIf(TF, "&H", "0x") & Mid(RGBT, 7, 2) & Mid(RGBT, 5, 2) & Mid(RGBT, 3, 2)
End Function

Public Sub iPrint(name As String, Tline As Boolean, Ticon As Boolean)
'Line����0��ʼ�����ģ���-2
    Dim H As Integer
    H = CInt(frmMain.txtNum.Text) * CInt(frmMain.txtY.Text)
'����ͼƬ�������İ�&���겻��0��
If linzi = True And CInt(frmMain.txtPict.Text) <> 0 Then
    frmMain.Wallpaper.Picture = LoadPicture(LoadP(GetINI("Setting", "WallpaperPath", App.Path & "\Config.ini"), 3))
    Tline = False
    Ticon = True
    frmMain.Picture1.AutoSize = True
    frmMain.Picture1.Picture = LoadPicture(LoadP(IIf(SavePath = "", App.Path, SavePath) & "\SM_BG_PIC.GIF", 4))
    PPicture
ElseIf Tline = True Then  '�����߿�
    frmMain.Wallpaper.Picture = LoadPicture(LoadP(GetINI("Setting", "WallpaperPath", App.Path & "\Config.ini"), 3))
    iLine 3, 220 - H - 20 - 4 - 1 - 1, CInt(frmMain.txtX.Text) - 2 + 2, H - 2 + 2, frmMain.lblColor(2).BackColor, True
    iLine 2, 220 - H - 20 - 4 - 1 - 2, CInt(frmMain.txtX.Text) - 2 + 4, H - 2 + 4, frmMain.lblColor(3).BackColor, True
    iLine 1, 220 - H - 20 - 4 - 1 - 3, CInt(frmMain.txtX.Text) - 2 + 6, H - 2 + 6, frmMain.lblColor(4).BackColor, True
    iLine 0, 220 - H - 20 - 4 - 1 - 4, CInt(frmMain.txtX.Text) - 2 + 8, H - 2 + 8, frmMain.lblColor(5).BackColor, True
    'iLine 4, 220 - H - 20 - 4 - 1, CInt(frmMain.txtX.Text) - 2, H - 2, frmMain.lblColor(0).BackColor, False
End If
    'ȡ������
    Dim i As Byte, ALLname As Variant 'Ԫ�ر���
    ALLname = Split(Get_Path_ALL_name_AND_Index(name), "/")

'ͼƬ����
If Ticon = True Then
    If CInt(frmMain.txtPict.Text) = 0 Then iLine 4, 220 - H - 20 - 4 - 1, CInt(frmMain.txtX.Text) - 2, H - 2, frmMain.lblColor(0).BackColor, False
    For i = 1 To UBound(ALLname)
        frmMain.Iicon(i).Picture = LoadPicture(LoadP(element(0).image & element(Get_Index(CStr(ALLname(i)))).image, 1))
        frmMain.Iicon(i).Top = 220 - CInt(frmMain.txtNum.Text) * CInt(frmMain.txtY.Text) - 20 - 4 - 1 + (i - 1) * CInt(frmMain.txtY.Text) + 2
        frmMain.Iicon(i).Left = 6
        frmMain.Iicon(i).Visible = True
        frmMain.Wallpaper.ForeColor = frmMain.lblColor(1).BackColor
        frmMain.Wallpaper.CurrentX = 5 + 16 + 2
        frmMain.Wallpaper.CurrentY = 220 - H - 20 - 4 - 1 + (i - 1) * CInt(frmMain.txtY.Text) + 1 + 1 - 3
        frmMain.Wallpaper.Print ALLname(i)
    Next
    For i = UBound(ALLname) + 1 To frmMain.Iicon.UBound
        frmMain.Iicon(i).Visible = False
    Next
End If

'���߿�
If CInt(frmMain.txtPict.Text) = 0 Then
    For i = 1 To CInt(frmMain.txtNum.Text)
        iLine 5, 220 - H - 20 - 4 - 1 + (i - 1) * CInt(frmMain.txtY.Text) + 1, CInt(frmMain.txtX.Text) - 2 - 2, CInt(frmMain.txtY.Text) - 2 - 1, frmMain.lblColor(0).BackColor, True
    Next
End If
    iLine 5, 220 - H - 20 - 4 - 1 + (ALLname(0) - 1) * CInt(frmMain.txtY.Text) + 1, CInt(frmMain.txtX.Text) - 2 - 2, CInt(frmMain.txtY.Text) - 2 - 1, frmMain.lblColor(6).BackColor, True

End Sub
Public Sub iLine(X%, Y%, W%, H%, iColor As Long, iTF As Boolean)
Dim X2%, Y2%
    X2 = X + W: Y2 = Y + H
    frmMain.Wallpaper.CurrentX = 0
    frmMain.Wallpaper.CurrentY = 0
    If iTF = True Then
        frmMain.Wallpaper.Line Step(X, Y)-(X2, Y2), iColor, B
    Else
        frmMain.Wallpaper.Line Step(X, Y)-(X2, Y2), iColor, BF
    End If
End Sub
Public Sub PPicture()
    frmMain.Wallpaper.CurrentX = 0
    frmMain.Wallpaper.CurrentY = 0
    'frmMain.Picture1.Visible = True
    frmMain.Wallpaper.ZOrder 1
    frmMain.Picture1.AutoSize = False
    If frmMain.Picture1.Width > frmMain.Picture1.Height Then
        frmMain.Picture1.Height = frmMain.Picture1.Width
    Else
        frmMain.Picture1.Width = frmMain.Picture1.Height
    End If
    frmMain.Picture1.BackColor = &H10101
    'frmMain.Picture1.ZOrder 0
    fxRender frmMain.Wallpaper.hdc, frmMain.Picture1.Width \ 15 \ 2, CInt(frmMain.txtPict.Text) + frmMain.Picture1.Height \ 15 \ 2, frmMain.Picture1.hdc, _
    0, 0, frmMain.Picture1.Height \ 15, 0, _
    255, 0, 1, frmMain.Picture1.BackColor
    'frmMain.Picture1.Visible = False
    frmMain.Wallpaper.ZOrder 0
End Sub







'�������ֻ��element���
Public Function Get_Index(name As String) As Integer
Dim i As Integer
For i = 1 To UBound(element)
    If element(i).name = name Then Get_Index = i: Exit Function
Next
End Function
'�������ֻ�ýṹ����
Public Function Get_Path_ALL_name_AND_Index(name As String) As String
Dim id As Integer, iIndex As Integer
Dim nowPath As String, nextPath As String
nowPath = Mid(frmMain.TreeView1.Nodes(name).FullPath, 1, Len(frmMain.TreeView1.Nodes(name).FullPath) - Len(frmMain.TreeView1.Nodes(name).Text))

id = frmMain.TreeView1.Nodes(name).Index
Do Until frmMain.TreeView1.Nodes(id).Previous Is Nothing
    id = frmMain.TreeView1.Nodes(id).Previous.Index
    nextPath = Mid(frmMain.TreeView1.Nodes(id).FullPath, 1, Len(frmMain.TreeView1.Nodes(id).FullPath) - Len(frmMain.TreeView1.Nodes(id).Text))
    If nowPath <> nextPath Then Get_Path_ALL_name_AND_Index = "0": Exit Function
    Get_Path_ALL_name_AND_Index = frmMain.TreeView1.Nodes(id).Text & "/" & Get_Path_ALL_name_AND_Index
    iIndex = iIndex + 1
Loop
Get_Path_ALL_name_AND_Index = iIndex + 1 & "/" & Get_Path_ALL_name_AND_Index & name
iIndex = iIndex + 1

id = frmMain.TreeView1.Nodes(name).Index
Do Until frmMain.TreeView1.Nodes(id).Next Is Nothing
    id = frmMain.TreeView1.Nodes(id).Next.Index
    nextPath = Mid(frmMain.TreeView1.Nodes(id).FullPath, 1, Len(frmMain.TreeView1.Nodes(id).FullPath) - Len(frmMain.TreeView1.Nodes(id).Text))
    If nowPath <> nextPath Then Get_Path_ALL_name_AND_Index = "0": Exit Function
    Get_Path_ALL_name_AND_Index = Get_Path_ALL_name_AND_Index & "/" & frmMain.TreeView1.Nodes(id).Text
    iIndex = iIndex + 1
Loop
If iIndex > CInt(frmMain.txtNum.Text) Then Get_Path_ALL_name_AND_Index = "0"
End Function

Public Function Get_TruePath(ByRef Path As String, iT As String) As Boolean

    If Len(Dir(Replace(Path, "\\", "\"))) = 0 Then
        If MsgBox(GetINI("lng", "Get_TruePath_MG", App.Path & "\Config.ini") & Chr(13) & Chr(10) & "��" & iT, vbYesNo, GetINI("lng", "title_MG", App.Path & "\Config.ini")) = vbYes Then
            frmMain.CommonDialog1.FileName = ""
            frmMain.CommonDialog1.Filter = GetINI("lng", "Get_TruePath_CF", App.Path & "\Config.ini") & iT
            frmMain.CommonDialog1.CancelError = True
            On Error Resume Next
            frmMain.CommonDialog1.ShowOpen
            If frmMain.CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Function
            Path = frmMain.CommonDialog1.FileName
            Get_TruePath = True
            Exit Function
        End If
    Else
        Get_TruePath = True
    End If
End Function
