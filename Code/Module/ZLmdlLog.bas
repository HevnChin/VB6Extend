Attribute VB_Name = "ZLmdlLog"
Option Explicit

Public Function ZLCE_WriteLog(ByVal strLogPath As String, ByVal strFunc As String, ByVal time As String, Optional ByVal strInput As String = "", Optional ByVal strOutPut As String = "") As Boolean
    '功能：记录日志文件，主要用于接口调试
On Error GoTo ErrH
    '以下用于记录调用接口的入参
'    Const strFile As String = "C:\TstLog_"
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Dim str年 As String, str月 As String
    Dim str时间 As String: str时间 = time
 
    str年 = Format(str时间, "YYYY")
    str月 = Format(str时间, "MM")
  
    strFileName = ZLCE_GetFullPath(strLogPath) & Format(str时间, "YYYYMMDD") & ".Log"
    
    Call ZLCE_Set多级目录(strFileName)
    
    If Not Dir(strFileName) <> "" Then
        objFileSystem.CreateTextFile strFileName
    End If
    
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
 
    objStream.WriteLine (String(50, "-"))
    objStream.WriteLine ("部件版本号:" & App.Major & "." & App.Minor & "." & App.Revision)
    objStream.WriteLine ("执行时间:" & str时间)
    objStream.WriteLine ("函数名:" & strFunc)
    objStream.WriteLine ("  入参:" & strInput)
    objStream.WriteLine ("  出参:" & strOutPut)
    objStream.WriteLine (String(50, "-"))
    objStream.Close
    Set objStream = Nothing
    ZLCE_WriteLog = True
    Exit Function
ErrH:
    ZLCE_WriteLog = False
End Function

Public Function ZLCE_Set多级目录(ByVal str文件路径 As String) As Boolean
On Error GoTo ErrH
    Dim var目录 As Variant
    Dim str目录 As String
    Dim objFileSystem As New FileSystemObject
    Dim i As Integer
    Dim strMsg As String
 
    var目录 = Split(str文件路径, "\")
    For i = LBound(var目录) To UBound(var目录) - 1
        str目录 = str目录 & "\" & var目录(i)
    
        If Left(str目录, 1) = "\" Then str目录 = Mid(str目录, 2)
        
        If Dir(str目录, vbDirectory) = "" Then
            objFileSystem.CreateFolder (str目录)
        End If
    Next
    ZLCE_Set多级目录 = True
    Exit Function
ErrH:
    ZLCE_Set多级目录 = False
End Function


Public Function ZLCE_GetFullPath(ByVal strPath As String)
'适配path环境
On Error GoTo ErrH
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    ZLCE_GetFullPath = strPath
     Exit Function
ErrH:
    ZLCE_GetFullPath = strPath
End Function

Public Function ZLCE_AppendPath(ByVal strPath As String, ByVal strAppend As String)
'追加 path环境
On Error GoTo ErrH
    ZLCE_AppendPath = ZLCE_GetFullPath(strPath)
    ZLCE_AppendPath = ZLCE_AppendPath & strAppend & "\"
    Exit Function
ErrH:
    ZLCE_AppendPath = ""
End Function
