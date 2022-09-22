Attribute VB_Name = "ZLmdlSQL"
Option Explicit
Private gddXml  As New DOMDocument                '�������ݰ�
 
 
Public Function ZLCE_DBAction( _
                ByVal gcnObj As ADODB.Connection, _
                ByVal execPro_Name As String, _
                ByVal execPro_InName As String, ByVal execPro_OutName As String, _
                             ByVal XmlStr_In As String, ByRef XmlStr_Out As String, _
                             Optional execPro_OutCode As String = "Res/Code", Optional execPro_OutMsg As String = "Res/Msg") As Boolean
    'gcnObj                          ���ݿ����
    'execProName             ִ�д洢����
    'execPro_InName        �������
    'execPro_OutName    ��������
    'XmlStr_In                     �������
    'XmlStr_Out                 ��������
    
    'execPro_OutCode   ����XML_ErrCode   Ĭ��:Res/Code
    'execPro_OutMsg    ����XML_ErrMsg     Ĭ��:Res/Msg
    
    'Code:0=OK  <Res><Code>0</Code><Msg>' || Err_Msg || '</Msg></Res>
    'Code:1=Error   <Res><Code>1</Code><Msg>' || Err_Msg || '</Msg></Res>
    
    '��������
    Dim cmdTmp As New ADODB.Command
    Dim cmdPara As New ADODB.Parameter
    Set cmdTmp = New ADODB.Command
    On Error GoTo ErrH
    '��Ҫִ�еĴ洢����
    
    With cmdTmp
        '����Ϊ�洢����
        .CommandType = adCmdStoredProc
        
        '����Ϊִ�й�������
        .CommandText = execPro_Name
        
        '���ӵ����ݿ�
        .ActiveConnection = gcnObj
                
        '��� [���] ����
        .Parameters.Append .CreateParameter(execPro_InName, adVarChar, adParamInput, 4000, XmlStr_In)
        .Parameters.Append .CreateParameter(execPro_OutName, adVarChar, adParamOutput, 4000, XmlStr_Out)
      
        'ִ�й���
        .Execute
      
        XmlStr_Out = .Parameters(execPro_OutName)
        
        '�ͷŲ���
        .Parameters.Delete (execPro_InName)
        .Parameters.Delete (execPro_OutName)
   End With

    If Not gddXml.loadXML(XmlStr_Out) Then
        Err.Raise 9999, "", "���ݿ�ִ�з���XML���ܽ�����"
        ZLCE_DBAction = False
        Exit Function
    End If
    
    If gddXml.selectSingleNode(execPro_OutCode).Text = "0" Then
        'gddXml.SelectSingleNode("project").Attributes.Item(0).Text  ����һ�ַ�ʽ
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
    'gcnObj                          ���ݿ����
    'proName                      ִ�д洢����
    'inputKeys                     ���KeyArray
    'inputValues                  ���ValueArray
    'outputKey                     �������
    'outputValue                 ��������
    'outCode                       ����code  Ĭ��:Res/Code
    'outMsg                          ����MSG  Ĭ��:Res/Msg
    
    'Code:0=OK  <Res><Code>0</Code><Msg>' || Err_Msg || '</Msg></Res>
    'Code:1=Error   <Res><Code>1</Code><Msg>' || Err_Msg || '</Msg></Res>
    
    '��������
    Dim cmdTmp As New ADODB.Command
    Dim cmdPara As New ADODB.Parameter
    Set cmdTmp = New ADODB.Command
    Dim itemKey As Variant, itemValue As String, index As Integer
    On Error GoTo ErrH
 
    '��Ҫִ�еĴ洢����
    
    With cmdTmp
        '����Ϊ�洢����
        .CommandType = adCmdStoredProc
        
        '����Ϊִ�й�������
        .CommandText = proName
        
        '���ӵ����ݿ�
        .ActiveConnection = gcnObj
                
        '��� [���] ����
        .Parameters.Append .CreateParameter(outputKey, adVarChar, adParamOutput, 4000, outputValue)
        For Each itemKey In inputKeys
            itemValue = inputValues(index)
            .Parameters.Append .CreateParameter(itemKey, adVarChar, adParamInput, 4000, itemValue)
            index = index + 1
        Next
        'ִ�й���
        .Execute
      
        outputValue = .Parameters(outputKey)
        
        'ɾ������
        For Each itemKey In inputKeys
            .Parameters.Delete (itemKey)
        Next
        .Parameters.Delete (outputKey)
   End With

    If Not gddXml.loadXML(outputValue) Then
        Err.Raise 9999, "", "���ݿ�ִ�з���XML���ܽ�����"
        ZLCE_DBProExec = False
        Exit Function
    End If
    
    If gddXml.selectSingleNode(outCode).Text = "0" Then
        'gddXml.SelectSingleNode("project").Attributes.Item(0).Text  ����һ�ַ�ʽ
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


