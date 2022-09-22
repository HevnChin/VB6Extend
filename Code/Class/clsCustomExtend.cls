VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZLCE"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'版本: 1.0.6
'时间:  2022.08.23
'作者:  wds.梁瀚文
'A.修改涉及:
'ZLCE_GetJsonByDictionary 增加判断: objType = "Variant()"
'ZLCE_GetXMLsByDictionary 增加判断: objType = "Variant()"
'ZLCE_StringAppend 增加可选参数:  IsSideAddSplitStr  , checkSame, 默认为Fasle
 
'B.新增涉及:
'TrimLeft  TrimRight TrimEdge XML_SetKeyValue

'版本: 1.0.7
'时间:  2022.08.31
'作者:  wds.梁瀚文
'新增:GetCustomSetting

'1.08
'时间:  2022.08.31
'作者:  wds.梁瀚文
'HttpRequest   DateGetUnixTimeStamp Exec_DBPro  ArrayGetStrCount


Private gSysName As String

'Http请求类型
Public Enum HttpRequestType
    HttpRequestType_XML = 0
    HttpRequestType_Text = 1
    HttpRequestType_Body = 2
    HttpRequestType_BodyText = 3
End Enum


'请求返回的数据样式
Public Enum HttpResponseType
    HttpResXML = 0
    HttpResText = 1
    HttpResBody = 2
End Enum


'属性设置
Public Property Let SysName(ByVal asysName As String)
    gSysName = asysName
    ZLCE_SysName = asysName
End Property
'
Public Property Get SysName() As String
    SysName = gSysName
End Property


'-------------------------------------ZLmdlDBSetting-------------------------------------
Public Function GetPluginSetting(ByVal execKey As String, Optional fetchColumn As String = "内容", Optional fetchTable As String = "ZLPlugin配置表") As String
      GetPluginSetting = ZLCE_GetPluginSetting(execKey, fetchColumn, fetchTable)
End Function
'
Public Function GetCustomSetting(ByVal tableName As String, ByVal whereColumnName As String, ByVal whereColumnValue As String, ByVal getColumn As String) As String
    GetCustomSetting = ZLCE_GetCustomSetting(tableName, whereColumnName, whereColumnValue, getColumn)
End Function
'
Public Function GetExecSQL(execKey As String, Optional fetchColumn As String = "内容", Optional fetchTable As String = "ZLPLUGINSQL查询表") As String
     GetExecSQL = ZLCE_GetExecSQL(execKey, fetchColumn, fetchTable)
End Function
'-------------------------------------ZLmdlDBSetting-------------------------------------


'-------------------------------------ZLmdlINI-------------------------------------
Public Function INI_GetValue(ByVal filename As String, ByVal AppName As String, ByVal KeyName As String) As String
    INI_GetValue = ZLCE_GetINI(filename, AppName, KeyName)
End Function
'
Public Function INI_SetValue(ByVal filename As String, ByVal AppName As String, ByVal KeyName As String, ByVal Entry As String) As Boolean
    INI_SetValue = ZLCE_SetINI(filename, AppName, KeyName, Entry) <> 0
End Function

'-------------------------------------ZLmdlINI-------------------------------------

'-------------------------------------ZLmdlLog-------------------------------------
'
Public Function Log_Write(ByVal strLogPath As String, ByVal strFunc As String, Optional ByVal strInput As String = "", Optional ByVal strOutPut As String = "") As Boolean
    '功能：记录日志文件，主要用于接口调试
    Log_Write = ZLCE_WriteLog(strLogPath, strFunc, CurrentDateTime, strInput, strOutPut)
End Function
'
Public Function Log_Set多级目录(ByVal str文件路径 As String) As Boolean
    Log_Set多级目录 = ZLCE_Set多级目录(str文件路径)
End Function

'
Public Function Log_GetFullPath(ByVal strPath As String) As String
    Log_GetFullPath = ZLCE_GetFullPath(strPath)
End Function
'
Public Function Log_AppendPath(ByVal strPath As String, ByVal strAppend As String) As String
    Log_AppendPath = ZLCE_AppendPath(strPath, strAppend)
End Function
'-------------------------------------ZLmdlLog-------------------------------------

'-------------------------------------ZLMdlPublic-------------------------------------
Public Function GetComputerName() As String
'=取得计算机名
    GetComputerName = ZLCE_GetComputerName
End Function
 
'
Function GetComputerIP() As String
'取得计算机IP地址
    GetComputerIP = ZLCE_GetComputerIP
End Function

'-------------------------------------GObj-------------------------------------
Public Function CurrentDateTime(Optional ByVal timeStyle As String = "yyyy-MM-dd HH:mm:ss") As String
    CurrentDateTime = Format(Now, "yyyy-MM-dd HH:mm:ss")
End Function
  
'-------------------------------------GObj-------------------------------------

'检测数据集是否有记录
Function ChkRsState(rs As ADODB.Recordset) As Boolean
    ChkRsState = ZLCE_ChkRsState(rs)
End Function
'
Public Function RsValue(ByVal rs As ADODB.Recordset, field As String) As String
'==================================================================================================
'=功能：获取字段数值
'=调用：当程序获取记录集字段值时
'=作者：王大聖
'=时间：2021-04-16
'=入参：
'=  rs ADO结果集
'=  key 字段名称
'==================================================================================================
 RsValue = ZLCE_RsValue(rs, field)
End Function

'-----------------------------------------------SQL---------------------------------------------------
 
'
Public Function Exec_DBAction( _
                ByVal gcnObj As ADODB.Connection, _
                ByVal execPro_Name As String, _
                ByVal execPro_InName As String, ByVal execPro_OutName As String, _
                             ByVal XmlStr_In As String, ByRef XmlStr_Out As String, _
                             Optional execPro_OutCode As String = "Res/Code", Optional execPro_OutMsg As String = "Res/Msg") As Boolean
        Exec_DBAction = ZLCE_DBAction(gcnObj, execPro_Name, execPro_InName, execPro_OutName, XmlStr_In, XmlStr_Out, execPro_OutCode, execPro_OutMsg)
End Function
'

'真正实现动执行数据库过程SQL
Public Function Exec_DBPro(ByVal gcnObj As ADODB.Connection, ByVal proName As String, _
                        ByRef inputKeys() As String, ByRef inputValues() As String, _
                        ByVal outputKey As String, ByRef outputValue As String, _
                        Optional outCode As String = "Res/Code", Optional outMsg As String = "Res/Msg") As Boolean
    Exec_DBPro = ZLCE_DBProExec(gcnObj, proName, inputKeys, inputValues, outputKey, outputValue, outCode, outMsg)
End Function
 
'--------------------------------------------String-------------------------------------------------------
'
Public Function Split(mStr As String, splitChar As String, index As Integer) As String
'以字符串切割,获取指定位置的数据,切割成数组再次索引
    Split = ZLCE_Split(mStr, splitChar, index)
End Function
'
Public Function SplitIndex(ByVal str As String, ByVal SplitStr As String, ByVal index As Integer) As String
'以字符串切割,获取指定位置的数据,直接截取更加高效
    SplitIndex = ZLCE_SplitIndex(str, SplitStr, index)
End Function

'
Public Function strAppend(ByRef mainStr As String, ByVal SplitStr As String, ByVal appendStr As String, Optional ByVal IsCheckNull As Boolean = False, _
                            Optional ByVal IsTrim As Boolean = False, Optional ByVal IsSideAddSplitStr As Boolean = False, Optional ByVal checkSame As Boolean = False) As String
    strAppend = ZLCE_StringAppend(mainStr, SplitStr, appendStr, IsCheckNull, IsTrim, IsSideAddSplitStr, checkSame)
End Function
'
Public Function ReplaceParamString(ByVal mainStr As String, ParamArray Items()) As String
''根据[1] [2] 参数替换信息
''王大聖
On Error GoTo ErrH
    Dim i
    Dim index As Integer, tmpStr As String
    
    tmpStr = mainStr
    '遍历解决(直接替换,不存在数据类型的问题)
    For Each i In Items
        index = index + 1
        '        Select Case TypeName()
        '            Case "String", "String()"
        '            Case Else
        '        End Select
        tmpStr = Replace(tmpStr, "[" & index & "]", i)
    Next
    
    ''获取替换字符串
    ReplaceParamString = tmpStr
    Exit Function
ErrH:
    ReplaceParamString = mainStr
End Function
 
'Public Function GetXMLByDictionary(ByVal dict As Dictionary, Optional ByVal show换行 As Boolean = True) As String
'废弃2022-05-06
''==================================================================================================
''=功能：将Dict转为String(只是一层的Dict)
''=调用：
''=入参：
''=  dict 字典
''= 出参：
''= XML换行字符串
'    GetXMLByDictionary = ZLCE_GetXMLByDictionary(dict, show换行)
'End Function
'
Public Function GetXMLsByDictionary(ByVal dict As Dictionary, Optional ByVal show换行 As Boolean = True) As String
'==================================================================================================
'=功能：将Dict转为String(复杂Dict深度切换)
'=调用：
'=入参：
'=  dict 字典
'= 出参：
'= XML换行字符串
    GetXMLsByDictionary = ZLCE_GetXMLsByDictionary(dict, show换行)
End Function
'
''获取json数据
Public Function GetJSONValue(ByVal JSONString As String, ByVal JSONPath As String) As Variant
    GetJSONValue = ZLCE_JSONParse(JSONString, JSONPath)
End Function
'
Public Function GetJsonByDictionary(ByVal dict As Dictionary) As String
'==================================================================================================
'=功能：将Dict转为String
'=入参：
'=  dict 字典
'= 出参：Json换行字符串
    GetJsonByDictionary = ZLCE_GetJsonByDictionary(dict)
End Function

'
Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
 '-----------------------------------------------------------------------------------------------------------
'--功  能:按指定长度填制空格
'--入参数:
'--出参数:
'--返  回:返回字串
    Lpad = ZLCE_Lpad(strCode, lngLen, strChar)
End Function
'
Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
'-----------------------------------------------------------------------------------------------------------
'--功  能:按指定长度填制空格
'--入参数:
'--出参数:
'--返  回:返回字串
    Rpad = ZLCE_Rpad(strCode, lngLen, strChar)
End Function
'
Public Function SubString(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'-----------------------------------------------------------------------------------------------------------
'--功  能:读取指定字串的值,字串中可以包含汉字
'--入参数:strInfor-原串
'         lngStart-直始位置
'         lngLen-长度
'--出参数:
'--返  回:子串
    SubString = ZLCE_SubString(strInfor, lngStart, lngLen)
End Function

'
Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = ZLCE_Nvl(varValue, DefaultValue)
End Function
'
Public Function ToLower(ByRef Expression) As String
'==================================================================================================
'=功能：String转小写
'=调用：String函数
'=入参：
'=  Expression 表达式
'= 出参：
'= 小写字符串
'-------------------------
    ToLower = ZLCE_ToLower(Expression)
End Function

'
Public Function ToUpper(ByRef Expression) As String
'==================================================================================================
'=功能：String转小写
'=调用：String函数
'=入参：
'=  Expression 表达式
'= 出参：
'= 小写字符串
'-------------------------
    ToUpper = ZLCE_ToUpper(Expression)
End Function
'
Public Function ToStr(ByRef Expression) As String
'==================================
'=功能：转String
'=调用：String函数
'=入参：
'=  Expression 表达式
'= 出参：
'= 字符串
    ToStr = ZLCE_ToStr(Expression)
End Function
'
Public Function ToNum(ByRef Expression) As Double
'==================================
'=功能：转Number
'=调用 Number函数
'=入参：
'=  Expression 表达式
'= 出参：
'=  Number
    ToNum = ZLCE_ToNum(Expression)
End Function
'
Public Function StrFitChar(ByVal str As String, ByVal IsLeft As Boolean, fixLen As Long, fixChar As String) As String
'功能: 适配String
'LoR : 左侧或者右侧
'fixLen :需要适配的长度
'fixChar: 适配的字符是什么
     StrFitChar = ZLCE_StrFitChar(str, IsLeft, fixLen, fixChar)
End Function
'
Public Function Decode(ParamArray arrPar() As Variant) As Variant
Rem ZLCE_Decode暂不使用
    Dim varValue As Variant, i As Integer
    '功能：模拟Oracle的Decode函数
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

'
Public Function Trim(ByVal str As String) As String
'功能：去掉字符串中\0以后的字符，并且去掉两端的空格
    Trim = ZLCE_Trim(str)
End Function

'
Public Function TrimLeft(ByVal mStr As String, Optional leftStr As String = " ") As String
    TrimLeft = ZLCE_TrimLeft(mStr, leftStr)
End Function
'
Public Function TrimRight(ByVal mStr As String, Optional rightStr As String = " ") As String
    TrimRight = ZLCE_TrimRight(mStr, rightStr)
End Function
'
Public Function TrimEdge(ByVal mStr As String, Optional sideStr As String = " ") As String
    TrimEdge = ZLCE_TrimEdge(mStr, sideStr)
End Function
 
 Public Function StrContains(ByVal mainStr As String, ByVal subStr As String, Optional ByVal SplitStr As String = ",") As Boolean
 '功能：是否包含字符串
     StrContains = ZLCE_ContainSubStr(mainStr, subStr, SplitStr)
 End Function
'-----------------------------------------------------------XML-----------------------------------------------------------
'
Public Function XML_GetLoadXMLObj(ByVal strXML As String) As Object
    Set XML_GetLoadXMLObj = ZLCE_GetLoadXMLObj(strXML)
End Function
'
Public Function XML_GetElemnetValue(ByVal doc As Object, ByVal name As String, Optional ByVal itemIndex As Integer = 0, Optional ByVal IsOrion As Boolean = False) As String
'功能：得到指定元素的值
  XML_GetElemnetValue = ZLCE_GetElemnetValue(doc, name, itemIndex, IsOrion)
End Function

'简化XML数值接口
Public Function XML_GetXMLNode(ByVal doc As Object, ByVal key As String) As String
    XML_GetXMLNode = ZLCE_GetElemnetValue(doc, key, 0, True)
End Function

Public Function XML_GetXMLStrNode(ByVal strXML As String, ByVal key As String) As String
'简化XML数值接口
    XML_GetXMLStrNode = ZLCE_GetXMLStrNode(strXML, key)
End Function

'
'简化XML数值接口
Public Function XML_GetXMLSingleNode(ByVal doc As Object, ByVal keyPath As String) As String
    XML_GetXMLSingleNode = ZLCE_GetXMLSingleNode(doc, keyPath)
End Function
'
Public Function XML_GetXMLStrSingleNode(ByVal strXML As String, ByVal keyPath As String) As String
'简化XML数值接口
    XML_GetXMLStrSingleNode = ZLCE_GetXMLStrSingleNode(strXML, keyPath)
End Function

'
Public Function XML_SetKeyValue(ByVal xmlStr As String, ByVal key As String, ByVal value As String, Optional index As Integer = 0) As String
    XML_SetKeyValue = ZLCE_XMLSetKeyValue(xmlStr, key, value, index)
End Function

'1.0.8New 获取节点LIST数据或者 子节点数组
Public Function XML_GetNodeByListIndex(ByVal doc As Object, ByVal ListKey As String, ByRef valueArray() As String, Optional subListKey As String = "", Optional idx As Integer = -1) As String
    XML_GetNodeByListIndex = ZLCE_XMLGetNodeByListIndex(doc, ListKey, valueArray, subListKey, idx)
End Function
  
 '---------------------------------------------------------Array 1.0.4------------------------------------------------
 '
 Public Function GetArrayByIndex(ByRef arr() As String, ByVal index As Integer) As String
    GetArrayByIndex = ZLCE_Str_GetArrayByIndex(arr, index)
End Function
'
 Public Function ArrayGetEntryByIndex(ByRef arr() As Variant, ByVal index As Integer) As Variant
    ArrayGetEntryByIndex = ZLCE_GetArrayByIndex(arr, index)
End Function
'
Rem 获取最小index
Public Function ArrayGetMinIndex(ByRef arr() As Variant) As Long
    ArrayGetMinIndex = ZLCE_GetArrayMinIndex(arr)
End Function
'
Rem 获取最大index
Public Function ArrayGetMaxIndex(ByRef arr() As Variant) As Long
    ArrayGetMaxIndex = ZLCE_GetArrayMaxIndex(arr)
End Function
'
Rem 获取Count
Public Function ArrayGetCount(ByRef arr() As Variant) As Long
    ArrayGetCount = ZLCE_GetArrayCount(arr)
End Function

Rem 1.0.8
Public Function ArrayGetStrCount(ByRef arr() As String) As Long
   ArrayGetStrCount = ZLCE_GetStrArrayCount(arr)
End Function

'1.0.8New
Public Function ArrayInsertStrIndex(ByRef arr() As String, ByVal str As String, Optional index As Integer = -1) As Boolean
    ArrayInsertStrIndex = ZLCE_Str_ArrayInsertIndex(arr, str, index)
End Function

'1.0.8New
Public Function ArrayInsertIndex(ByRef arr() As Variant, ByVal var As Variant, Optional index As Integer = -1) As Boolean
    ArrayInsertIndex = ZLCE_ArrayInsertIndex(arr, var, index)
End Function
 
'---------------------------------------------------------Date 1.0.8----------------------------------------------------
Rem Date
Public Function DateGetUnixTimeStamp(ByVal strDate As String, Optional isMillis As Boolean = False) As String
    DateGetUnixTimeStamp = ZLCE_DateGetUnixTimeStamp(strDate, isMillis)
End Function
'---------------------------------------------------------Http 1.0.8----------------------------------------------------
Public Function HttpRequest(ByVal reqURL As String, ByVal reqContent As String, ByVal httpReqType As HttpRequestType, _
                                                                                    Optional Method As String = "POST", Optional reqHeaderKeyValues As Dictionary = Nothing) As String
    HttpRequest = ZLCE_XMLHTTPRequest(reqURL, reqContent, httpReqType, Method, reqHeaderKeyValues)
End Function
