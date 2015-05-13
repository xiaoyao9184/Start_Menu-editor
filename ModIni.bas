Attribute VB_Name = "ModIni"
Option Explicit

'INI读写模块
'**************************************

'INI文件API函数
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
'LpApplicationName: 被读写段的段名?
'LpKeyName: 被读写的条目的关键字名?
'LpString：写入INI文件的值（当进行写操作时）。
'lpDefault:：段中的默认返回值以及（或者）读操作时关键字未找到时的默认返回值。
'lpReturnedString：从INI文件读到的数据（当进行读操作时）。
'nSize: 返回的最大字符数?设置为 lpReturnedString的长度?
'lpFileName：INI文件的文件名，包括完整的路径名。
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



