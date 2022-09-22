Attribute VB_Name = "ZLmdlSQL"
Option Explicit
Private gddXml  As New DOMDocument                '请求数据包
 
 
Public Function ZLCE_DBAction( _
                ByVal gcnObj As ADODB.Connection, _
                ByVal execPro_Name As String, _
                ByVal execPro_InName As String, ByVal execPro_OutName As String, _
                             ByVal XmlStr_In As String, ByRef XmlStr_Out As String, _
                             Optional execPro_OutCode As String = "Res/Code", Optional execPro_OutMsg As String = "Res/Msg") As Boolean
    'gcnObj                          数据库对象
    'execProName             执行存储过程
    'execPro_InName        入参名称
    'execPro_OutName    出参名称
    'XmlStr_In                     入参数据
    'XmlStr_Out                 出参数据
    
    'execPro_OutCode   出参XML_ErrCode   默认:Res/Code
    'execPro_OutMsg    出参XML_ErrMsg     默认:Res/Msg
    
    'Code:0=OK  <Res><Code>0</Code><Msg>' || Err_Msg || '</Msg></Res>
    'Code:1=Error   <Res><Code>1</Code><Msg>' || Err_Msg || '</Msg></Res>
    
    '变量声名
    Dim cmdTmp As New ADODB.Command
    Dim cmdPara As New ADODB.Parameter
    Set cmdTmp = New ADODB.Command
    On Error GoTo ErrH
    '需要执行的存储过程
    
    With cmdTmp
        '类型为存储过程
        .CommandType = adCmdStoredProc
        
        '名称为执行过程名称
        .CommandText = execPro_Name
        
        '连接的数据库
        .ActiveConnection = gcnObj
                
        '添加 [入出] 参数
        .Parameters.Append .CreateParameter(execPro_InName, adVarChar, adParamInput, 4000, XmlStr_In)
        .Parameters.Append .CreateParameter(execPro_OutName, adVarChar, adParamOutput, 4000, XmlStr_Out)
      
        '执行过程
        .Execute
      
        XmlStr_Out = .Parameters(execPro_OutName)
        
        '释放参数
        .Parameters.Delete (execPro_InName)
        .Parameters.Delete (execPro_OutName)
   End With

    If Not gddXml.loadXML(XmlStr_Out) Then
        Err.Raise 9999, "", "数据库执行返回XML不能解析！"
        ZLCE_DBAction = False
        Exit Function
    End If
    
    If gddXml.selectSingleNode(execPro_OutCode).Text = "0" Then
        'gddXml.SelectSingleNode("project").Attributes.Item(0).Text  另外一种方式
        ZLCE_DBAction = True
        Exit Function
    Else
        Err.Raise 9999, "", gddXml.selectSingleNode(execPro_OutMsg).Text
        ZLCE_DBAction = False
        Exit Function
    End If
    Exit Function
ErrH:
    ZLCE_DBAction = False
    MsgBox Err.Description, vbCritical, ZLCE_Nvl(ZLCE_SysName, "VB6Extend")
    If 0 = 1 Then
        Resume
    End If
End Function
 
'1.0.8x
Public Function ZLCE_DBProExec(ByVal gcnObj As ADODB.Connection, ByVal proName As String, _
                        ByRef inputKeys() As String, ByRef inputValues() As String, _
                        ByVal outputKey As String, ByRef outputValue As String, _
                        Optional outCode As String = "Res/Code", Optional outMsg As String = "Res/Msg") As Boolean
    'gcnObj                          数据库对象
    'proName                      执行存储过程
    'inputKeys                     入参KeyArray
    'inputValues                  入参ValueArray
    'outputKey                     入参数据
    'outputValue                 出参数据
    'outCode                       传出code  默认:Res/Code
    'outMsg                          传出MSG  默认:Res/Msg
    
    'Code:0=OK  <Res><Code>0</Code><Msg>' || Err_Msg || '</Msg></Res>
    'Code:1=Error   <Res><Code>1</Code><Msg>' || Err_Msg || '</Msg></Res>
    
    '变量声名
    Dim cmdTmp As New ADODB.Command
    Dim cmdPara As New ADODB.Parameter
    Set cmdTmp = New ADODB.Command
    Dim itemKey As Variant, itemValue As String, index As Integer
    On Error GoTo ErrH
 
    '需要执行的存储过程
    
    With cmdTmp
        '类型为存储过程
        .CommandType = adCmdStoredProc
        
        '名称为执行过程名称
        .CommandText = proName
        
        '连接的数据库
        .ActiveConnection = gcnObj
                
        '添加 [入出] 参数
        .Parameters.Append .CreateParameter(outputKey, adVarChar, adParamOutput, 4000, outputValue)
        For Each itemKey In inputKeys
            itemValue = inputValues(index)
            .Parameters.Append .CreateParameter(itemKey, adVarChar, adParamInput, 4000, itemValue)
            index = index + 1
        Next
        '执行过程
        .Execute
      
        outputValue = .Parameters(outputKey)
        
        '删除参数
        For Each itemKey In inputKeys
            .Parameters.Delete (itemKey)
        Next
        .Parameters.Delete (outputKey)
   End With

    If Not gddXml.loadXML(outputValue) Then
        Err.Raise 9999, "", "数据库执行返回XML不能解析！"
        ZLCE_DBProExec = False
        Exit Function
    End If
    
    If gddXml.selectSingleNode(outCode).Text = "0" Then
        'gddXml.SelectSingleNode("project").Attributes.Item(0).Text  另外一种方式
        ZLCE_DBProExec = True
        Exit Function
    Else
        Err.Raise 9999, "", gddXml.selectSingleNode(outMsg).Text
        ZLCE_DBProExec = False
        Exit Function
    End If
    Exit Function
ErrH:
    ZLCE_DBProExec = False
    MsgBox Err.Description, vbCritical, ZLCE_Nvl(ZLCE_SysName, "VB6Extend")
    If 0 = 1 Then
        Resume
    End If
End Function


