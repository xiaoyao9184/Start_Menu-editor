Attribute VB_Name = "ModIni"
Option Explicit

'INI��дģ��
'**************************************

'INI�ļ�API����
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
  (ByVal LpApplicationName As String, _
  ByVal LpKeyName As Any, _
  ByVal lpDefault As String, _
  ByVal lpReturnedString As String, _
  ByVal nSize As Long, _
  ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
  (ByVal LpApplicationName As String, _
  ByVal LpKeyName As Any, _
  ByVal LpString As Any, _
  ByVal lpFileName As String) As Long
'LpApplicationName: ����д�εĶ���?
'LpKeyName: ����д����Ŀ�Ĺؼ�����?
'LpString��д��INI�ļ���ֵ��������д����ʱ����
'lpDefault:�����е�Ĭ�Ϸ���ֵ�Լ������ߣ�������ʱ�ؼ���δ�ҵ�ʱ��Ĭ�Ϸ���ֵ��
'lpReturnedString����INI�ļ����������ݣ������ж�����ʱ����
'nSize: ���ص�����ַ���?����Ϊ lpReturnedString�ĳ���?
'lpFileName��INI�ļ����ļ���������������·������
'---------------------------------------------
Public Function GetINI(LpApplicationName As String, LpKeyName As String, Path As String) As String
    Dim Retval As Long
    Dim Value As String
    Value = Space(128)
    Retval = GetPrivateProfileString(LpApplicationName, LpKeyName, "", Value, Len(Value), Path)
    GetINI = Left(Trim(Value), Len(Trim(Value)) - 1)
End Function

Public Sub WriteINI(LpApplicationName As String, LpKeyName As String, Value As String, Path As String)
    Dim Retval As Long
    Retval = WritePrivateProfileString(LpApplicationName, LpKeyName, Value, Path)
End Sub

Public Function GetType(LpApplicationName As String, LpKeyName As String, Path As String) As String ', ByVal iType As String)
    Dim Retval As Long
    Dim Value As String
    Value = Space(128)
    Retval = GetPrivateProfileString(LpApplicationName, LpKeyName, "", Value, Len(Value), Path)
    GetType = Left(Trim(Value), Len(Trim(Value)) - 1)
End Function



