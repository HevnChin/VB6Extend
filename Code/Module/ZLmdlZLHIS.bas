Attribute VB_Name = "ZLmdlZLHIS"
Option Explicit

Public Type ZLCE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    部门 As String
    站点 As String
    工作站 As String
    IP地址 As String
End Type

Public ZLCE_UserInfo As ZLCE_USER_INFO


Public Function ZLCE_GetNo(ByVal int门诊1住院2 As Integer) As String
''根据业务场景
''获取记帐单号
On Error GoTo ErrH
    Dim strType As Integer
    If int门诊1住院2 = 2 Then
        strType = 14
    Else
        strType = 13
    End If
    
    ZLCE_SQLString = "Select NextNO([1],'0','','1') As NO From Dual"
    Set ZLCE_Rscord = ZLCE_G_Database.OpenSQLRecord(ZLCE_SQLString, "获取单据号", strType)
    ZLCE_GetNo = ZLCE_Rscord!NO & ""                         '门诊Or住院 记账单据产生
    Exit Function
ErrH:
    ZLCE_GetNo = ""
End Function

'========================================================================================
'功能：获取登陆用户信息
'========================================================================================
Public Function ZLCE_GetUserInfo() As ZLCE_USER_INFO
On Error GoTo ErrH
    Dim rsUser As ADODB.Recordset
    Dim strSQL As String: strSQL = "Select B.ID, B.编号, B.简码, B.姓名, C.编码 As 部门编码, C.名称 As 部门名称, D.部门ID, A.用户名, B.站点" & vbNewLine & _
                "From 上机人员表 A, 人员表 B, 部门表 C, 部门人员 D" & vbNewLine & _
                "Where A.人员id = B.id" & vbNewLine & _
                "And B.ID = D.人员ID" & vbNewLine & _
                "And D.缺省 = 1" & vbNewLine & _
                "And D.部门id = C.id" & vbNewLine & _
                "And A.用户名 = User"
    Set rsUser = ZLCE_G_Database.OpenSQLRecord(strSQL, "Get Userinfo")
    
    If Not ZLCE_ChkRsState(rsUser) Then
        ZLCE_UserInfo.ID = ZLCE_RsValue(rsUser, "ID")
        ZLCE_UserInfo.编号 = ZLCE_RsValue(rsUser, "编号")
        ZLCE_UserInfo.部门ID = ZLCE_RsValue(rsUser, "部门ID")
        ZLCE_UserInfo.简码 = ZLCE_RsValue(rsUser, "简码")
        ZLCE_UserInfo.姓名 = ZLCE_RsValue(rsUser, "姓名")
        ZLCE_UserInfo.部门 = ZLCE_RsValue(rsUser, "部门名称")
        ZLCE_UserInfo.用户名 = ZLCE_RsValue(rsUser, "用户名")
        ZLCE_UserInfo.站点 = ZLCE_RsValue(rsUser, "站点")
        ZLCE_UserInfo.工作站 = ZLCE_GetComputerName
        ZLCE_UserInfo.IP地址 = ZLCE_GetComputerIP
    End If
    ZLCE_GetUserInfo = ZLCE_UserInfo
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

