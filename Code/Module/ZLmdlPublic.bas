Attribute VB_Name = "ZLMdlPublic"
Option Explicit

Public ZLCE_SysName As String
Public ZLCE_SQLString As String
Public ZLCE_Rscord As ADODB.Recordset

'取得计算机名字[MAC地址]
Public Declare Function ZLCE_Lib_GetPCName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'取得计算机IP
Public Declare Sub ZLCE_Lib_MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function ZLCE_Lib_GetIP Lib "IPHlpApi" Alias "GetIpAddrTable" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

'========================================================================================
'=取得计算机名
'========================================================================================
Public Function ZLCE_GetComputerName() As String
On Error GoTo ErrH
    Dim sBuffer As String * 255
    If ZLCE_Lib_GetPCName(sBuffer, 255&) <> 0 Then
        ZLCE_GetComputerName = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        ZLCE_GetComputerName = "(未知)"
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'========================================================================================
'取得计算机IP地址
'========================================================================================
Function ZLCE_GetComputerIP() As String
    Dim lngIP               As Long
    Dim lRet                As Long
    Dim Buffer()            As Byte
    Dim addrByte(3)         As Byte
    Dim Cnt                 As Long
    Dim strIP               As String
On Error GoTo ErrH:
    ZLCE_Lib_GetIP ByVal 0&, lRet, True
    If lRet <= 0 Then Exit Function
    ReDim Buffer(0 To lRet - 1) As Byte
    ' 取回 IP 地址的相关数据
    ZLCE_Lib_GetIP Buffer(0), lRet, False
    ZLCE_Lib_MoveMemory lngIP, Buffer(4 + (0 * Len(lngIP))), Len(lngIP)
    ZLCE_Lib_MoveMemory addrByte(0), lngIP, 4
    For Cnt = 0 To 3
        strIP = strIP + CStr(addrByte(Cnt)) + "."
    Next Cnt
    ZLCE_GetComputerIP = Left(strIP, Len(strIP) - 1)
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function


'Public ZLCE_Gobj_Database As Object: ZLCE_Gobj_Database = CreateObject("zl9Comlib.clsDatabase")
'Public ZLCE_Gobj_CommFun As Object: ZLCE_Gobj_Database = CreateObject("zl9Comlib.clsCommfun")
'Public ZLCE_Gobj_Control As Object: ZLCE_Gobj_Database = CreateObject("zl9Comlib.clsControl")
'Public ZLCE_Gobj_ComLib As Object: ZLCE_Gobj_Database = CreateObject("zl9Comlib.clsComlib")
'Public ZLCE_Gobj_PrintMode As Object: ZLCE_Gobj_Database = CreateObject("zl9PrintMode.zlPrintMethod")
'Public ZLCE_Gobj_Report As Object: ZLCE_Gobj_Database = CreateObject("zl9Report.clsReport")
 
'全局定义
Static Function ZLCE_G_Database() As Object
'获取 clsDatabase
On Error GoTo ErrH
Dim ZLCE_Database As Object
    If IsNull(ZLCE_Database) Or ZLCE_Database Is Nothing Then
        Set ZLCE_Database = CreateObject("zl9Comlib.clsDatabase")
    End If
    Set ZLCE_G_Database = ZLCE_Database
    Exit Function
ErrH:
End Function
 
'全局定义
Static Function ZLCE_G_CommFun() As Object
'获取 clsCommfun
On Error GoTo ErrH
Dim ZLCE_CommFun As Object
    If IsNull(ZLCE_CommFun) Or ZLCE_CommFun Is Nothing Then
        Set ZLCE_CommFun = CreateObject("zl9Comlib.clsCommfun")
    End If
    Set ZLCE_G_CommFun = ZLCE_CommFun
    Exit Function
ErrH:
End Function
 
Static Function ZLCE_G_Control() As Object
'获取 clsControl
On Error GoTo ErrH
Dim ZLCE_Control As Object
    If IsNull(ZLCE_Control) Or ZLCE_Control Is Nothing Then
        Set ZLCE_Control = CreateObject("zl9Comlib.clsControl")
    End If
    Set ZLCE_G_Control = ZLCE_Control
    Exit Function
ErrH:
End Function
 
Static Function ZLCE_G_ComLib() As Object
'获取 clsComlib
On Error GoTo ErrH
Dim ZLCE_ComLib As Object
    If IsNull(ZLCE_ComLib) Or ZLCE_ComLib Is Nothing Then
        Set ZLCE_ComLib = CreateObject("zl9Comlib.clsComlib")
    End If
    Set ZLCE_G_ComLib = ZLCE_ComLib
    Exit Function
ErrH:
End Function
 
Static Function ZLCE_G_PrintMode() As Object
'获取 zlPrintMethod
On Error GoTo ErrH
Dim ZLCE_PrintMode As Object
    If IsNull(ZLCE_PrintMode) Or ZLCE_PrintMode Is Nothing Then
        Set ZLCE_PrintMode = CreateObject("zl9PrintMode.zlPrintMethod")
    End If
    Set ZLCE_G_PrintMode = ZLCE_PrintMode
    Exit Function
ErrH:
End Function

Static Function ZLCE_G_Report() As Object
'获取 clsReport
On Error GoTo ErrH
Dim ZLCE_Report As Object
    If IsNull(ZLCE_Report) Or ZLCE_Report Is Nothing Then
        Set ZLCE_Report = CreateObject("zl9Report.clsReport")
    End If
    Set ZLCE_G_Report = ZLCE_Report
    Exit Function
ErrH:
End Function


