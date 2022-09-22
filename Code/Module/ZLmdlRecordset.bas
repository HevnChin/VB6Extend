Attribute VB_Name = "ZLmdlRecordset"
Option Explicit

'������ݼ��Ƿ��м�¼
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
'=���ܣ���ȡ�ֶ���ֵ
'=���ã��������ȡ��¼���ֶ�ֵʱ
'=���ߣ������}
'=ʱ�䣺2021-04-16
'=��Σ�
'=  rs ADO�����
'=  key �ֶ�����
'==================================================================================================
On Error GoTo ErrH
    ZLCE_RsValue = rs.fields.Item(field).Value & ""
    Exit Function
ErrH:
    Err.Clear
    ZLCE_RsValue = ""
End Function

