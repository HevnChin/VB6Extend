Attribute VB_Name = "ZLmdlINI"
Option Explicit

Public Declare Function ZLCE_Lib_GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function ZLCE_Lib_WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long


Public Function ZLCE_GetINI(ByVal filename As String, ByVal AppName As String, ByVal KeyName As String) As String
  Dim RetStr As String
  RetStr = String(255, Chr(0))
  ZLCE_Lib_GetPrivateProfileString AppName, ByVal KeyName, "", RetStr, Len(RetStr), filename
  ZLCE_GetINI = Left(RetStr, InStr(1, RetStr, Chr(0)) - 1)
End Function

Public Function ZLCE_SetINI(ByVal filename As String, ByVal AppName As String, ByVal KeyName As String, ByVal Entry As String) As Long
  ZLCE_SetINI = ZLCE_Lib_WritePrivateProfileString(AppName, ByVal KeyName, Entry, filename)
End Function


