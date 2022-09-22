Attribute VB_Name = "ZLmdlDBSetting"
Option Explicit

''----------------------------------------------------------------------------------------Custom----------------------------------------------------------------------------------------''
Public Function ZLCE_GetCustomSetting(ByVal tableName As String, ByVal whereColumnName As String, ByVal whereColumnValue As String, ByVal getColumn As String) As String
''获取【普通表】配置
'2022-08-31
''梁瀚文
''tableName 表名 |whereColumnName where字段名 |whereColumnValue where字段值 | getColumn 要获取的字段值


On Error GoTo ErrH
    Dim rs As New ADODB.Recordset
    
    ZLCE_SQLString = "Select Max(A." & getColumn & ") As " & getColumn & "  From " & tableName & " A Where A." & whereColumnName & " = [1]"
    Set rs = ZLCE_G_Database.OpenSQLRecord(ZLCE_SQLString, "查询表:" & tableName, whereColumnValue)
    If ZLCE_ChkRsState(rs) = False Then
        ZLCE_GetCustomSetting = ZLCE_RsValue(rs, getColumn)
        Else
        ZLCE_GetCustomSetting = ""
    End If
    
    '返回数据
    Exit Function
ErrH:
    MsgBox "ZLCE_GetCustomSetting" & Err.Description, vbCritical, ZLCE_Nvl(ZLCE_SysName, "VB6Extend")
    If 1 = 0 Then
        Debug.Print Err.Description
        Resume
    End If
    Err.Clear
End Function

''----------------------------------------------------------------------------------------DBSetting----------------------------------------------------------------------------------------''
Public Function ZLCE_GetPluginSetting(ByVal execKey As String, Optional fetchColumn As String = "内容", Optional fetchTable As String = "ZLPlugin配置表") As String
''获取ZLplugin配置信息
'2021-05-24
''王大聖
On Error GoTo ErrH
    Dim rs As New ADODB.Recordset
    
    ZLCE_SQLString = "Select Max(A." & fetchColumn & ") As " & fetchColumn & "  From " & fetchTable & " A Where A.标识 = [1]"
    Set rs = ZLCE_G_Database.OpenSQLRecord(ZLCE_SQLString, execKey, execKey)
    If ZLCE_ChkRsState(rs) = False Then
        ZLCE_GetPluginSetting = ZLCE_RsValue(rs, fetchColumn)
        Else
        ZLCE_GetPluginSetting = ""
    End If
    
    '返回数据
    Exit Function
ErrH:
    MsgBox "ZLCE_GetPluginSetting" & Err.Description, vbCritical, ZLCE_Nvl(ZLCE_SysName, "VB6Extend")
    If 1 = 0 Then
        Debug.Print Err.Description
        Resume
    End If
    Err.Clear
End Function

''--------------------------------------------------------------------------------获取执行SQL(工具)--------------------------------------------------------------------------------
Public Function ZLCE_GetExecSQL(execKey As String, Optional fetchColumn As String = "内容", Optional fetchTable As String = "ZLPLUGINSQL查询表") As String
''获取执行SQL
''王大圣
''2021-04-15
On Error GoTo ErrH
    Dim tmpStr As String
    Dim rs As New ADODB.Recordset
    tmpStr = ""
    
    ZLCE_SQLString = "Select A." & fetchColumn & " From " & fetchTable & " A Where A.标识 = [1] Order By A.序号 Asc"
    Set rs = ZLCE_G_Database.OpenSQLRecord(ZLCE_SQLString, execKey, execKey)
    With rs
        Do While Not .EOF
            tmpStr = tmpStr & vbCrLf & ZLCE_RsValue(rs, fetchColumn)
        .MoveNext
        Loop
    End With
    
    ZLCE_GetExecSQL = tmpStr
    '返回数据
    Exit Function
ErrH:
    ZLCE_GetExecSQL = ""
    MsgBox "ZLCE_GetExecSQL" & Err.Description, vbCritical, ZLCE_Nvl(ZLCE_SysName, "VB6Extend")
    If 1 = 0 Then
        Debug.Print Err.Description
        Resume
    End If
    Err.Clear
End Function


