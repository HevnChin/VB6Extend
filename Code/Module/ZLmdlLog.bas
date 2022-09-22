Attribute VB_Name = "ZLmdlLog"
Option Explicit

Public Function ZLCE_WriteLog(ByVal strLogPath As String, ByVal strFunc As String, ByVal time As String, Optional ByVal strInput As String = "", Optional ByVal strOutPut As String = "") As Boolean
    '���ܣ���¼��־�ļ�����Ҫ���ڽӿڵ���
On Error GoTo ErrH
    '�������ڼ�¼���ýӿڵ����
'    Const strFile As String = "C:\TstLog_"
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Dim str�� As String, str�� As String
    Dim strʱ�� As String: strʱ�� = time
 
    str�� = Format(strʱ��, "YYYY")
    str�� = Format(strʱ��, "MM")
  
    strFileName = ZLCE_GetFullPath(strLogPath) & Format(strʱ��, "YYYYMMDD") & ".Log"
    
    Call ZLCE_Set�༶Ŀ¼(strFileName)
    
    If Not Dir(strFileName) <> "" Then
        objFileSystem.CreateTextFile strFileName
    End If
    
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
 
    objStream.WriteLine (String(50, "-"))
    objStream.WriteLine ("�����汾��:" & App.Major & "." & App.Minor & "." & App.Revision)
    objStream.WriteLine ("ִ��ʱ��:" & strʱ��)
    objStream.WriteLine ("������:" & strFunc)
    objStream.WriteLine ("  ���:" & strInput)
    objStream.WriteLine ("  ����:" & strOutPut)
    objStream.WriteLine (String(50, "-"))
    objStream.Close
    Set objStream = Nothing
    ZLCE_WriteLog = True
    Exit Function
ErrH:
    ZLCE_WriteLog = False
End Function

Public Function ZLCE_Set�༶Ŀ¼(ByVal str�ļ�·�� As String) As Boolean
On Error GoTo ErrH
    Dim varĿ¼ As Variant
    Dim strĿ¼ As String
    Dim objFileSystem As New FileSystemObject
    Dim i As Integer
    Dim strMsg As String
 
    varĿ¼ = Split(str�ļ�·��, "\")
    For i = LBound(varĿ¼) To UBound(varĿ¼) - 1
        strĿ¼ = strĿ¼ & "\" & varĿ¼(i)
    
        If Left(strĿ¼, 1) = "\" Then strĿ¼ = Mid(strĿ¼, 2)
        
        If Dir(strĿ¼, vbDirectory) = "" Then
            objFileSystem.CreateFolder (strĿ¼)
        End If
    Next
    ZLCE_Set�༶Ŀ¼ = True
    Exit Function
ErrH:
    ZLCE_Set�༶Ŀ¼ = False
End Function


Public Function ZLCE_GetFullPath(ByVal strPath As String)
'����path����
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
'׷�� path����
On Error GoTo ErrH
    ZLCE_AppendPath = ZLCE_GetFullPath(strPath)
    ZLCE_AppendPath = ZLCE_AppendPath & strAppend & "\"
    Exit Function
ErrH:
    ZLCE_AppendPath = ""
End Function
