Attribute VB_Name = "ZLmdlZLHIS"
Option Explicit

Public Type ZLCE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ���� As String
    վ�� As String
    ����վ As String
    IP��ַ As String
End Type

Public ZLCE_UserInfo As ZLCE_USER_INFO


Public Function ZLCE_GetNo(ByVal int����1סԺ2 As Integer) As String
''����ҵ�񳡾�
''��ȡ���ʵ���
On Error GoTo ErrH
    Dim strType As Integer
    If int����1סԺ2 = 2 Then
        strType = 14
    Else
        strType = 13
    End If
    
    ZLCE_SQLString = "Select NextNO([1],'0','','1') As NO From Dual"
    Set ZLCE_Rscord = ZLCE_G_Database.OpenSQLRecord(ZLCE_SQLString, "��ȡ���ݺ�", strType)
    ZLCE_GetNo = ZLCE_Rscord!NO & ""                         '����OrסԺ ���˵��ݲ���
    Exit Function
ErrH:
    ZLCE_GetNo = ""
End Function

'========================================================================================
'���ܣ���ȡ��½�û���Ϣ
'========================================================================================
Public Function ZLCE_GetUserInfo() As ZLCE_USER_INFO
On Error GoTo ErrH
    Dim rsUser As ADODB.Recordset
    Dim strSQL As String: strSQL = "Select B.ID, B.���, B.����, B.����, C.���� As ���ű���, C.���� As ��������, D.����ID, A.�û���, B.վ��" & vbNewLine & _
                "From �ϻ���Ա�� A, ��Ա�� B, ���ű� C, ������Ա D" & vbNewLine & _
                "Where A.��Աid = B.id" & vbNewLine & _
                "And B.ID = D.��ԱID" & vbNewLine & _
                "And D.ȱʡ = 1" & vbNewLine & _
                "And D.����id = C.id" & vbNewLine & _
                "And A.�û��� = User"
    Set rsUser = ZLCE_G_Database.OpenSQLRecord(strSQL, "Get Userinfo")
    
    If Not ZLCE_ChkRsState(rsUser) Then
        ZLCE_UserInfo.ID = ZLCE_RsValue(rsUser, "ID")
        ZLCE_UserInfo.��� = ZLCE_RsValue(rsUser, "���")
        ZLCE_UserInfo.����ID = ZLCE_RsValue(rsUser, "����ID")
        ZLCE_UserInfo.���� = ZLCE_RsValue(rsUser, "����")
        ZLCE_UserInfo.���� = ZLCE_RsValue(rsUser, "����")
        ZLCE_UserInfo.���� = ZLCE_RsValue(rsUser, "��������")
        ZLCE_UserInfo.�û��� = ZLCE_RsValue(rsUser, "�û���")
        ZLCE_UserInfo.վ�� = ZLCE_RsValue(rsUser, "վ��")
        ZLCE_UserInfo.����վ = ZLCE_GetComputerName
        ZLCE_UserInfo.IP��ַ = ZLCE_GetComputerIP
    End If
    ZLCE_GetUserInfo = ZLCE_UserInfo
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

