Attribute VB_Name = "ZLmdlString"
Option Explicit

Public Function ZLCE_Split(mStr As String, splitChar As String, index As Integer) As String
'以字符串切割,获取指定位置的数据
'王大聖 2021-06-18
On Error GoTo ErrH
    If InStr(1, mStr, splitChar) <= 0 Then
        ZLCE_Split = mStr
        Exit Function
    End If
    
    Dim arr: arr = Split(mStr, splitChar)
    If IsArray(arr) Then ZLCE_Split = arr(index)
    Exit Function
ErrH:
ZLCE_Split = ""
End Function


'1.0.6版本增加 2022.08.23 于青海妇儿
Public Function ZLCE_TrimLeft(ByVal mStr As String, Optional sideStr As String = " ") As String
'以字符Side格式化
'王大聖 2022-08-23
On Error GoTo ErrH
    Dim str As String: str = mStr
    Dim strLen As Long: strLen = Len(str)
    Dim strSideLen As Long: strSideLen = Len(sideStr)
    
    If sideStr = " " Or strLen <= 0 Then
        ZLCE_TrimLeft = Trim(str)
        Exit Function
    Else
        If Left(str, strSideLen) = sideStr Then str = ZLCE_SubString(str, strSideLen, strLen - strSideLen)
        ZLCE_TrimLeft = str
    End If
    Exit Function
ErrH:
    ZLCE_TrimLeft = ""
End Function
'1.0.6版本增加 2022.08.23 于青海妇儿
Public Function ZLCE_TrimRight(ByVal mStr As String, Optional sideStr As String = " ") As String
'以字符Side格式化
'王大聖 2022-08-23
On Error GoTo ErrH
    Dim str As String: str = mStr
    Dim strLen As Long: strLen = Len(str)
    Dim strSideLen As Long: strSideLen = Len(sideStr)
    
    If sideStr = " " Or strLen <= 0 Then
        ZLCE_TrimRight = Trim(str)
        Exit Function
    Else
        If Right(str, strSideLen) = sideStr Then str = ZLCE_SubString(str, 0, strLen - strSideLen)
        ZLCE_TrimRight = str
    End If
    Exit Function
ErrH:
    ZLCE_TrimRight = ""
End Function

Public Function ZLCE_TrimEdge(ByVal mStr As String, Optional sideStr As String = " ") As String
'以字符Side格式化
'王大聖 2022-08-23
On Error GoTo ErrH
    Dim str As String: str = mStr
    Dim strLen As Long: strLen = Len(mStr)
    Dim strSideLen As Long: strSideLen = Len(sideStr)
        
    If sideStr = " " Or strLen <= 0 Then
        ZLCE_TrimEdge = Trim(str)
        Exit Function
    Else
        If Left(str, strSideLen) = sideStr Then str = ZLCE_SubString(str, strSideLen, strLen - strSideLen)
        If Right(str, strSideLen) = sideStr Then str = ZLCE_SubString(str, 0, Len(str) - strSideLen)
        ZLCE_TrimEdge = str
    End If
    Exit Function
ErrH:
    ZLCE_TrimEdge = ""
End Function



Public Function ZLCE_Trim(ByVal str As String) As String
'功能：去掉字符串中\0以后的字符，并且去掉两端的空格
    Dim Nstr As String
    If InStr(str, Chr(0)) > 0 Then
        Nstr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        Nstr = Trim(str)
    End If
    ZLCE_Trim = Replace(Replace(Replace(Nstr, Chr(13), vbCr), vbLf, ""), vbTab, "")
End Function


Public Function ZLCE_ReplaceParamString(ByVal mainStr As String, ParamArray Items()) As String
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
    ZLCE_ReplaceParamString = tmpStr
    Exit Function
ErrH:
    ZLCE_ReplaceParamString = mainStr
End Function

Public Function ZLCE_GetXMLByDictionary(ByVal dict As Dictionary, Optional ByVal show换行 As Boolean = True) As String
'==================================================================================================
'=功能：将Dict转为String
'=调用：
'=入参：
'=  dict 字典
'= 出参：
'= XML换行字符串
'-------------------------
'作者: 王大聖
'时间: 2021-06-06
'==================================================================================================
'demo
'            Dim dict As New Dictionary
'            dict.Add "1", "1Values"
'            Dim vKey As Variant
'            For Each vKey In dict
'                Debug.Print vKey
'            Next
On Error GoTo ErrH
    Dim tmpStr As String: tmpStr = ""
    Dim vKey As Variant, objType As String
    
    For Each vKey In dict
        objType = TypeName(dict.Item(vKey))
        If objType = "Dictionary" Then
            tmpStr = tmpStr & "<" & vKey & ">" & ZLCE_GetXMLByDictionary(dict.Item(vKey), show换行) & "</" & vKey & ">"
        Else
            If IsNull(dict.Item(vKey)) Or Len(dict.Item(vKey)) <= 0 Then
                tmpStr = tmpStr & "<" & vKey & "></" & vKey & ">" & IIf(show换行, vbNewLine, "")
            Else
                tmpStr = tmpStr & "<" & vKey & ">" & dict.Item(vKey) & "</" & vKey & ">" & IIf(show换行, vbNewLine, "")
            End If
        End If
        
    Next
    
    ZLCE_GetXMLByDictionary = tmpStr
    Exit Function
ErrH:
    Err.Clear
    ZLCE_GetXMLByDictionary = ""
End Function

Public Function ZLCE_GetXMLsByDictionary(ByVal dict As Dictionary, Optional ByVal show换行 As Boolean = True) As String
'==================================================================================================
'=功能：将Dict转为String
'=调用：
'=入参：
'=  dict 字典
'= 出参：
'= XML换行字符串
'-------------------------
'作者: 研发部
'时间: 2021-06-06
'==================================================================================================
'demo
'Dim dictArray1(0 To 1) As Dictionary
'Dim dict_Data As New Dictionary
'Dim dictArray(0 To 1) As Dictionary
'Dim dict_Details As New Dictionary
'Dim dict_Details_sub As New Dictionary
'Dim dict_Item1 As New Dictionary
'Dim dict_Item1_sub As New Dictionary
'Dim dict_Item2 As New Dictionary
'Dim dict_Item2_sub As New Dictionary

'dict_Item1_sub.Add "ItemCode", "9527"
'dict_Item1_sub.Add "ItemName", "东坡肉"
'dict_Item1.Add "Item", dict_Item1_sub
'dict_Item2_sub.Add "ItemCode", "9528"
'dict_Item2_sub.Add "ItemName", "西门吹雪"
'dict_Item2.Add "Item", dict_Item2_sub
'
'Set dictArray(0) = dict_Item1
'Set dictArray(1) = dict_Item2
'
'dict_Details.Add "Details", dictArray
'Set dictArray1(0) = dict_Details
'Set dictArray1(1) = dict_Details
'dict_Data.Add "data", dictArray1
'
'Dim strA As String
'strA = GetXMLsByDictionary(dict_Data)

On Error GoTo ErrH
    Dim str换行 As String: str换行 = IIf(show换行, vbNewLine, "")
    Dim tmpStr As String: tmpStr = ""
    Dim vKey As Variant, tmpArrayStr As String, i As Integer, objType As String
    For Each vKey In dict
        objType = TypeName(dict.Item(vKey))
        If objType = "Dictionary" Then
            tmpStr = tmpStr & "<" & vKey & ">" & str换行 & ZLCE_GetXMLsByDictionary(dict.Item(vKey), show换行) & "</" & vKey & ">" & str换行
        ElseIf objType = "Object()" Or objType = "Variant()" Then
            tmpArrayStr = "<" & vKey & ">" & str换行
            For i = LBound(dict.Item(vKey)) To UBound(dict.Item(vKey)) Step 1 '数组里面放字典
                objType = TypeName(dict.Item(vKey))
                If objType = "Dictionary" Or objType = "Object()" Or objType = "Variant()" Then
                     tmpArrayStr = tmpArrayStr & ZLCE_GetXMLsByDictionary(dict.Item(vKey)(i), show换行) '& IIf(show换行, vbNewLine, "")
                Else
                    tmpArrayStr = tmpArrayStr & "<" & vKey & ">" & ZLCE_Nvl(dict.Item(vKey)(i), "") & "</" & vKey & ">" & str换行
                End If
            Next i
            
            tmpArrayStr = tmpArrayStr & "</" & vKey & ">" & IIf(show换行, vbNewLine, "")
            tmpStr = tmpStr & tmpArrayStr
        Else
            If IsNull(dict.Item(vKey)) Or Len(dict.Item(vKey)) <= 0 Then
                tmpStr = tmpStr & "<" & vKey & "></" & vKey & ">" & str换行
            Else
                tmpStr = tmpStr & "<" & vKey & ">" & dict.Item(vKey) & "</" & vKey & ">" & str换行
            End If
        End If
        
        Rem 其他写法 tmpStr = tmpStr & Chr(60) & vKey & Char(62) & value & Chr(60) & Chr(47) & vKey & Char(62) & vbNewLine
    Next
    
    ZLCE_GetXMLsByDictionary = tmpStr
    'Debug.Print tmpStr
    Exit Function
ErrH:
    Err.Clear
    ZLCE_GetXMLsByDictionary = ""
End Function

Public Function ZLCE_GetJsonByDictionary(ByVal dict As Dictionary) As String
'1.0.6版本优化:
'增加判断: objType = "Variant()"

'==================================================================================================
'=功能：将Dict转为String
'=入参：
'=  dict 字典
'= 出参：Json换行字符串
'-------------------------
'作者: 王大聖
'时间: 2021-06-20
'==================================================================================================
'demo
'    Dim arr(1 To 2) As Dictionary, arr1(1 To 1) As Dictionary, jsonStr As String
'
'    Dim dict As New Dictionary, dict1 As New Dictionary, dict2 As New Dictionary, dict3 As New Dictionary
'    dict3.Add "uniqueCode", "E3"
'    dict3.Add "jfCode", "j3"
'    dict3.Add "patientId", 13
'    Set arr1(1) = dict3
'
'    dict1.Add "uniqueCode", "E211610003001001"
'    dict1.Add "jfCode", "0103027601"
'    dict1.Add "patientId", 1021
'
'    dict2.Add "uniqueCode", "E211610003001002"
'    dict2.Add "jfCode", "0103027602"
'    dict2.Add "patientId", arr1
'
'    Set arr(1) = dict1
'    Set arr(2) = dict2
'
'    dict.Add "name", "测试"
'    dict.Add "author", "WSF"
'    dict.Add "array", arr
'    jsonStr = GetJsonByDictionary(dict)
On Error GoTo ErrH
    Dim tmpStr As String: tmpStr = "{"
    Dim vKey As Variant
    Dim YH As String, MH As String, DH As String
    Dim tmpArrayStr  As String
    Dim i As Integer, objType As String
    
    MH = Chr(58)
    YH = Chr(34)
    DH = Chr(44)

    For Each vKey In dict
        objType = TypeName(dict.Item(vKey))
        If objType = "Dictionary" Then
            tmpStr = tmpStr & YH & vKey & YH & MH & ZLCE_GetJsonByDictionary(dict.Item(vKey)) & DH
            
        ElseIf IsArray(dict.Item(vKey)) Or objType = "Object()" Or objType = "Variant()" Then
        '字典里面放数组
            tmpArrayStr = "["
             
            For i = LBound(dict.Item(vKey)) To UBound(dict.Item(vKey)) Step 1 '数组里面放字典
                objType = TypeName(dict.Item(vKey))
                If objType = "Dictionary" Or objType = "Object()" Or objType = "Variant()" Then
                     tmpArrayStr = tmpArrayStr & ZLCE_GetJsonByDictionary(dict.Item(vKey)(i)) & DH
                Else
                    tmpArrayStr = tmpArrayStr & YH & ZLCE_Nvl(dict.Item(vKey)(i), "") & YH & DH
                End If
                
            Next i
            
            '取消最后的,号
            If Right(tmpArrayStr, 1) = DH Then
                tmpArrayStr = Left(tmpArrayStr, Len(tmpArrayStr) - 1)
            End If
            tmpArrayStr = tmpArrayStr & "]"
            
            '数组数据串串
            tmpStr = tmpStr & YH & vKey & YH & MH & tmpArrayStr & DH
        Else
            tmpStr = tmpStr & YH & vKey & YH & MH & YH & dict.Item(vKey) & YH & DH
        End If
         
    Next
    
    '取消最后一个逗号
    If Right(tmpStr, 1) = DH Then
        tmpStr = Left(tmpStr, Len(tmpStr) - 1)
    End If
    
    ZLCE_GetJsonByDictionary = tmpStr & "}"
    'Debug.Print GetJsonByDictionary
    Exit Function
ErrH:
    Err.Clear
    ZLCE_GetJsonByDictionary = ""
End Function


Public Function ZLCE_Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
On Error GoTo ErrH
 '-----------------------------------------------------------------------------------------------------------
'--功  能:按指定长度填制空格
'--入参数:
'--出参数:
'--返  回:返回字串
'-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '大于长度时,自动载断
        strTmp = ZLCE_SubString(strCode, 0, lngLen)
    End If
    ZLCE_Lpad = Replace(strTmp, Chr(0), strChar)
    Exit Function
ErrH:
    ZLCE_Lpad = strCode
End Function

Public Function ZLCE_Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
On Error GoTo ErrH
'-----------------------------------------------------------------------------------------------------------
'--功  能:按指定长度填制空格
'--入参数:
'--出参数:
'--返  回:返回字串
'-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '主要有空格引起的
        strTmp = ZLCE_SubString(strCode, 0, lngLen)
    End If
    '取掉最后半个字符
    ZLCE_Rpad = Replace(strTmp, Chr(0), strChar)
    Exit Function
ErrH:
    ZLCE_Rpad = strCode
End Function

Public Function ZLCE_SubString(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'-----------------------------------------------------------------------------------------------------------
'--功  能:读取指定字串的值,字串中可以包含汉字
'--入参数:strInfor-原串
'         lngStart-直始位置
'         lngLen-长度
'--出参数:
'--返  回:子串
'-----------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
On Error GoTo ErrH

    'ZLCE_SubString = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    'ZLCE_SubString = Replace(ZLCE_SubString, Chr(0), " ")
    ZLCE_SubString = Mid(strInfor, lngStart + 1, lngLen)
    Exit Function
ErrH:
    ZLCE_SubString = ""
End Function


Public Function ZLCE_Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Dim varReturn As Variant
    ZLCE_Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ZLCE_ToLower(ByRef Expression) As String
'==================================================================================================
'=功能：String转小写
'=调用：String函数
'=入参：
'=  Expression 表达式
'= 出参：
'= 小写字符串
'-------------------------
'作者: 王大聖
'时间: 2021-04-15
'==================================================================================================
On Error GoTo ErrH
    ZLCE_ToLower = StrConv(Expression, vbLowerCase)
    Exit Function
ErrH:
    ZLCE_ToLower = Expression
End Function


Public Function ZLCE_ToUpper(ByRef Expression) As String
'==================================================================================================
'=功能：String转小写
'=调用：String函数
'=入参：
'=  Expression 表达式
'= 出参：
'= 小写字符串
'-------------------------
'作者: 王大聖
'时间: 2021-04-15
'==================================================================================================
On Error GoTo ErrH
    ZLCE_ToUpper = StrConv(Expression, vbUpperCase)
    Exit Function
ErrH:
    ZLCE_ToUpper = Expression
End Function

Public Function ZLCE_ToStr(ByRef Expression) As String
'==================================
'=功能：转String
'=调用：String函数
'=入参：
'=  Expression 表达式
'= 出参：
'= 字符串
'-------------------------
'作者: 王大聖
'时间: 2021-07-23
'==================================
On Error GoTo ErrH
    ZLCE_ToStr = CStr(Expression & "")
    Exit Function
ErrH:
    ZLCE_ToStr = ""
End Function

Public Function ZLCE_ToNum(ByRef Expression) As Double
'==================================
'=功能：转Number
'=调用 Number函数
'=入参：
'=  Expression 表达式
'= 出参：
'=  Number
'-------------------------
'作者: 王大聖
'时间: 2021-07-23
'==================================
On Error GoTo ErrH
    ZLCE_ToNum = Val(Expression & "")
    Exit Function
ErrH:
    ZLCE_ToNum = 0
End Function

Public Function ZLCE_StrFitChar(ByVal str As String, ByVal IsLeft As Boolean, fixLen As Long, fixChar As String) As String
'功能: 适配String
'LoR : 左侧或者右侧
'fixLen :需要适配的长度
'fixChar: 适配的字符是什么
 
    Dim rStr As String, fixStr As String
    
    '长度处理
    If Len(str) > fixLen Then
        ZLCE_StrFitChar = str
        Exit Function
    End If
    
    '适配String
    fixStr = String(fixLen - LenB(StrConv(str, vbFromUnicode)), fixChar)
    If IsLeft Then
          rStr = fixStr & CStr(str)
    Else
        rStr = CStr(str) & fixStr
    End If
    
    ZLCE_StrFitChar = rStr
End Function


Public Function ZLCE_Decode(ParamArray arrPar() As Variant) As Variant
    Dim varValue As Variant, i As Integer
    '功能：模拟Oracle的Decode函数
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            ZLCE_Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            ZLCE_Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function


'1.0.6 2022.08.23 增加校验相同值参数,默认false
Public Function ZLCE_StringAppend(ByRef mainStr As String, ByVal SplitStr As String, ByVal appendStr As String, Optional ByVal IsCheckNull As Boolean = False, _
                            Optional ByVal IsTrim As Boolean = False, Optional ByVal IsSideAddSplitStr As Boolean = False, Optional ByVal checkSame As Boolean = False) As String
'功能：追加字符串并且 是否trim
    Dim tmpStr As String, splitLen As Integer: splitLen = Len(SplitStr)
    
    '1.校验边际位置添加splitStr
    If IsSideAddSplitStr Then
         If Left(mainStr, splitLen) <> SplitStr Then
            mainStr = SplitStr & mainStr
         End If
    End If
    
    '1.1.checkSame
    If checkSame Then
        tmpStr = mainStr
        '检查是否重复.若side存在直接比对
        If IsSideAddSplitStr Then
            If InStr(1, tmpStr, SplitStr & appendStr & SplitStr) >= 1 Then
                ZLCE_StringAppend = mainStr
                Exit Function
                '退出函数
            End If
        Else
            If InStr(1, SplitStr & tmpStr & SplitStr, SplitStr & appendStr & SplitStr) >= 1 Then
                ZLCE_StringAppend = mainStr
                Exit Function
                '退出函数
            End If
        End If
    End If
    
    '2.处理字符串
    If IsCheckNull Then
        mainStr = mainStr & IIf((Len(mainStr) >= 1 And Right(mainStr, splitLen) <> SplitStr), SplitStr, "") & appendStr
    Else
        mainStr = mainStr & SplitStr & appendStr
    End If

    
    '3.Trim
    If IsTrim Then
        mainStr = ZLCE_Trim(mainStr)
    End If
    
    '4.边际增加splitStr
    If IsSideAddSplitStr Then
        mainStr = mainStr & SplitStr
    End If
    
    ZLCE_StringAppend = mainStr
End Function

Public Function ZLCE_SplitIndex(ByVal str As String, ByVal SplitStr As String, ByVal index As Integer) As String
On Error GoTo ErrH
'功能：去掉字符串中\0以后的字符，并且去掉两端的空格
  ZLCE_SplitIndex = Split(str, SplitStr)(index)
  Exit Function
ErrH:
  ZLCE_SplitIndex = ""
End Function

Public Function ZLCE_ContainSubStr(ByVal mainStr As String, ByVal subStr As String, Optional ByVal SplitStr As String = ",") As Boolean
On Error GoTo ErrH
'功能：查询字符串是不是在主字符串中出现
    Dim tmpMainStr As String: tmpMainStr = mainStr
    Dim tmpSubStr As String: tmpSubStr = subStr
    
    'MainStr
    If ZLCE_SubString(tmpMainStr, 0, Len(SplitStr)) <> SplitStr Then
        tmpMainStr = SplitStr & tmpMainStr
    End If
    
    If ZLCE_SubString(tmpMainStr, Len(tmpMainStr) - Len(SplitStr), Len(SplitStr)) <> SplitStr Then
        tmpMainStr = tmpMainStr & SplitStr
    End If
    
    'SubStr
    If ZLCE_SubString(tmpSubStr, 0, Len(SplitStr)) <> SplitStr Then
        tmpSubStr = SplitStr & tmpSubStr
    End If
    
    If ZLCE_SubString(tmpSubStr, Len(tmpSubStr) - Len(SplitStr), Len(SplitStr)) <> SplitStr Then
        tmpSubStr = tmpSubStr & SplitStr
    End If

    ZLCE_ContainSubStr = InStr(1, tmpMainStr, tmpSubStr)
Exit Function
ErrH:
    ZLCE_ContainSubStr = False
End Function

