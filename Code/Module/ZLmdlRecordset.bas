Attribute VB_Name = "ZLmdlRecordset"
Option Explicit

'检测数据集是否有记录
Function ZLCE_ChkRsState(rs As ADODB.Recordset) As Boolean
On Error GoTo ErrH:
    With rs
        If rs Is Nothing Then
            ZLCE_ChkRsState = True
            Exit Function
        Else
            ZLCE_ChkRsState = False
        End If
        If rs.State = 0 Then
            ZLCE_ChkRsState = True
            Exit Function
        Else
            ZLCE_ChkRsState = False
        End If
        If .RecordCount < 1 Then
            ZLCE_ChkRsState = True
        Else
            ZLCE_ChkRsState = False
        End If
        If .EOF Or .BOF Then
            ZLCE_ChkRsState = True
        Else
            ZLCE_ChkRsState = False
        End If
    End With
    Exit Function
ErrH:
    Err.Clear
    ZLCE_ChkRsState = True
End Function

Public Function ZLCE_RsValue(ByVal rs As ADODB.Recordset, field As String) As String
'==================================================================================================
'=功能：获取字段数值
'=调用：当程序获取记录集字段值时
'=作者：王大聖
'=时间：2021-04-16
'=入参：
'=  rs ADO结果集
'=  key 字段名称
'==================================================================================================
On Error GoTo ErrH
    ZLCE_RsValue = rs.fields.Item(field).Value & ""
    Exit Function
ErrH:
    Err.Clear
    ZLCE_RsValue = ""
End Function

