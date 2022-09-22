Attribute VB_Name = "ZLmdlString"
Option Explicit

Public Function ZLCE_Split(mStr As String, splitChar As String, index As Integer) As String
'���ַ����и�,��ȡָ��λ�õ�����
'�����} 2021-06-18
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


'1.0.6�汾���� 2022.08.23 ���ຣ����
Public Function ZLCE_TrimLeft(ByVal mStr As String, Optional sideStr As String = " ") As String
'���ַ�Side��ʽ��
'�����} 2022-08-23
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
'1.0.6�汾���� 2022.08.23 ���ຣ����
Public Function ZLCE_TrimRight(ByVal mStr As String, Optional sideStr As String = " ") As String
'���ַ�Side��ʽ��
'�����} 2022-08-23
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
'���ַ�Side��ʽ��
'�����} 2022-08-23
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
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�������ȥ�����˵Ŀո�
    Dim Nstr As String
    If InStr(str, Chr(0)) > 0 Then
        Nstr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        Nstr = Trim(str)
    End If
    ZLCE_Trim = Replace(Replace(Replace(Nstr, Chr(13), vbCr), vbLf, ""), vbTab, "")
End Function


Public Function ZLCE_ReplaceParamString(ByVal mainStr As String, ParamArray Items()) As String
''����[1] [2] �����滻��Ϣ
''�����}
On Error GoTo ErrH
    Dim i
    Dim index As Integer, tmpStr As String
    
    tmpStr = mainStr
    '�������(ֱ���滻,�������������͵�����)
    For Each i In Items
        index = index + 1
        '        Select Case TypeName()
        '            Case "String", "String()"
        '            Case Else
        '        End Select
        tmpStr = Replace(tmpStr, "[" & index & "]", i)
    Next
    
    ''��ȡ�滻�ַ���
    ZLCE_ReplaceParamString = tmpStr
    Exit Function
ErrH:
    ZLCE_ReplaceParamString = mainStr
End Function

Public Function ZLCE_GetXMLByDictionary(ByVal dict As Dictionary, Optional ByVal show���� As Boolean = True) As String
'==================================================================================================
'=���ܣ���DictתΪString
'=���ã�
'=��Σ�
'=  dict �ֵ�
'= ���Σ�
'= XML�����ַ���
'-------------------------
'����: �����}
'ʱ��: 2021-06-06
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
            tmpStr = tmpStr & "<" & vKey & ">" & ZLCE_GetXMLByDictionary(dict.Item(vKey), show����) & "</" & vKey & ">"
        Else
            If IsNull(dict.Item(vKey)) Or Len(dict.Item(vKey)) <= 0 Then
                tmpStr = tmpStr & "<" & vKey & "></" & vKey & ">" & IIf(show����, vbNewLine, "")
            Else
                tmpStr = tmpStr & "<" & vKey & ">" & dict.Item(vKey) & "</" & vKey & ">" & IIf(show����, vbNewLine, "")
            End If
        End If
        
    Next
    
    ZLCE_GetXMLByDictionary = tmpStr
    Exit Function
ErrH:
    Err.Clear
    ZLCE_GetXMLByDictionary = ""
End Function

Public Function ZLCE_GetXMLsByDictionary(ByVal dict As Dictionary, Optional ByVal show���� As Boolean = True) As String
'==================================================================================================
'=���ܣ���DictתΪString
'=���ã�
'=��Σ�
'=  dict �ֵ�
'= ���Σ�
'= XML�����ַ���
'-------------------------
'����: �з���
'ʱ��: 2021-06-06
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
'dict_Item1_sub.Add "ItemName", "������"
'dict_Item1.Add "Item", dict_Item1_sub
'dict_Item2_sub.Add "ItemCode", "9528"
'dict_Item2_sub.Add "ItemName", "���Ŵ�ѩ"
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
    Dim str���� As String: str���� = IIf(show����, vbNewLine, "")
    Dim tmpStr As String: tmpStr = ""
    Dim vKey As Variant, tmpArrayStr As String, i As Integer, objType As String
    For Each vKey In dict
        objType = TypeName(dict.Item(vKey))
        If objType = "Dictionary" Then
            tmpStr = tmpStr & "<" & vKey & ">" & str���� & ZLCE_GetXMLsByDictionary(dict.Item(vKey), show����) & "</" & vKey & ">" & str����
        ElseIf objType = "Object()" Or objType = "Variant()" Then
            tmpArrayStr = "<" & vKey & ">" & str����
            For i = LBound(dict.Item(vKey)) To UBound(dict.Item(vKey)) Step 1 '����������ֵ�
                objType = TypeName(dict.Item(vKey))
                If objType = "Dictionary" Or objType = "Object()" Or objType = "Variant()" Then
                     tmpArrayStr = tmpArrayStr & ZLCE_GetXMLsByDictionary(dict.Item(vKey)(i), show����) '& IIf(show����, vbNewLine, "")
                Else
                    tmpArrayStr = tmpArrayStr & "<" & vKey & ">" & ZLCE_Nvl(dict.Item(vKey)(i), "") & "</" & vKey & ">" & str����
                End If
            Next i
            
            tmpArrayStr = tmpArrayStr & "</" & vKey & ">" & IIf(show����, vbNewLine, "")
            tmpStr = tmpStr & tmpArrayStr
        Else
            If IsNull(dict.Item(vKey)) Or Len(dict.Item(vKey)) <= 0 Then
                tmpStr = tmpStr & "<" & vKey & "></" & vKey & ">" & str����
            Else
                tmpStr = tmpStr & "<" & vKey & ">" & dict.Item(vKey) & "</" & vKey & ">" & str����
            End If
        End If
        
        Rem ����д�� tmpStr = tmpStr & Chr(60) & vKey & Char(62) & value & Chr(60) & Chr(47) & vKey & Char(62) & vbNewLine
    Next
    
    ZLCE_GetXMLsByDictionary = tmpStr
    'Debug.Print tmpStr
    Exit Function
ErrH:
    Err.Clear
    ZLCE_GetXMLsByDictionary = ""
End Function

Public Function ZLCE_GetJsonByDictionary(ByVal dict As Dictionary) As String
'1.0.6�汾�Ż�:
'�����ж�: objType = "Variant()"

'==================================================================================================
'=���ܣ���DictתΪString
'=��Σ�
'=  dict �ֵ�
'= ���Σ�Json�����ַ���
'-------------------------
'����: �����}
'ʱ��: 2021-06-20
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
'    dict.Add "name", "����"
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
        '�ֵ����������
            tmpArrayStr = "["
             
            For i = LBound(dict.Item(vKey)) To UBound(dict.Item(vKey)) Step 1 '����������ֵ�
                objType = TypeName(dict.Item(vKey))
                If objType = "Dictionary" Or objType = "Object()" Or objType = "Variant()" Then
                     tmpArrayStr = tmpArrayStr & ZLCE_GetJsonByDictionary(dict.Item(vKey)(i)) & DH
                Else
                    tmpArrayStr = tmpArrayStr & YH & ZLCE_Nvl(dict.Item(vKey)(i), "") & YH & DH
                End If
                
            Next i
            
            'ȡ������,��
            If Right(tmpArrayStr, 1) = DH Then
                tmpArrayStr = Left(tmpArrayStr, Len(tmpArrayStr) - 1)
            End If
            tmpArrayStr = tmpArrayStr & "]"
            
            '�������ݴ���
            tmpStr = tmpStr & YH & vKey & YH & MH & tmpArrayStr & DH
        Else
            tmpStr = tmpStr & YH & vKey & YH & MH & YH & dict.Item(vKey) & YH & DH
        End If
         
    Next
    
    'ȡ�����һ������
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
'--��  ��:��ָ���������ƿո�
'--�����:
'--������:
'--��  ��:�����ִ�
'-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '���ڳ���ʱ,�Զ��ض�
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
'--��  ��:��ָ���������ƿո�
'--�����:
'--������:
'--��  ��:�����ִ�
'-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '��Ҫ�пո������
        strTmp = ZLCE_SubString(strCode, 0, lngLen)
    End If
    'ȡ��������ַ�
    ZLCE_Rpad = Replace(strTmp, Chr(0), strChar)
    Exit Function
ErrH:
    ZLCE_Rpad = strCode
End Function

Public Function ZLCE_SubString(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'-----------------------------------------------------------------------------------------------------------
'--��  ��:��ȡָ���ִ���ֵ,�ִ��п��԰�������
'--�����:strInfor-ԭ��
'         lngStart-ֱʼλ��
'         lngLen-����
'--������:
'--��  ��:�Ӵ�
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
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Dim varReturn As Variant
    ZLCE_Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ZLCE_ToLower(ByRef Expression) As String
'==================================================================================================
'=���ܣ�StringתСд
'=���ã�String����
'=��Σ�
'=  Expression ���ʽ
'= ���Σ�
'= Сд�ַ���
'-------------------------
'����: �����}
'ʱ��: 2021-04-15
'==================================================================================================
On Error GoTo ErrH
    ZLCE_ToLower = StrConv(Expression, vbLowerCase)
    Exit Function
ErrH:
    ZLCE_ToLower = Expression
End Function


Public Function ZLCE_ToUpper(ByRef Expression) As String
'==================================================================================================
'=���ܣ�StringתСд
'=���ã�String����
'=��Σ�
'=  Expression ���ʽ
'= ���Σ�
'= Сд�ַ���
'-------------------------
'����: �����}
'ʱ��: 2021-04-15
'==================================================================================================
On Error GoTo ErrH
    ZLCE_ToUpper = StrConv(Expression, vbUpperCase)
    Exit Function
ErrH:
    ZLCE_ToUpper = Expression
End Function

Public Function ZLCE_ToStr(ByRef Expression) As String
'==================================
'=���ܣ�תString
'=���ã�String����
'=��Σ�
'=  Expression ���ʽ
'= ���Σ�
'= �ַ���
'-------------------------
'����: �����}
'ʱ��: 2021-07-23
'==================================
On Error GoTo ErrH
    ZLCE_ToStr = CStr(Expression & "")
    Exit Function
ErrH:
    ZLCE_ToStr = ""
End Function

Public Function ZLCE_ToNum(ByRef Expression) As Double
'==================================
'=���ܣ�תNumber
'=���� Number����
'=��Σ�
'=  Expression ���ʽ
'= ���Σ�
'=  Number
'-------------------------
'����: �����}
'ʱ��: 2021-07-23
'==================================
On Error GoTo ErrH
    ZLCE_ToNum = Val(Expression & "")
    Exit Function
ErrH:
    ZLCE_ToNum = 0
End Function

Public Function ZLCE_StrFitChar(ByVal str As String, ByVal IsLeft As Boolean, fixLen As Long, fixChar As String) As String
'����: ����String
'LoR : �������Ҳ�
'fixLen :��Ҫ����ĳ���
'fixChar: ������ַ���ʲô
 
    Dim rStr As String, fixStr As String
    
    '���ȴ���
    If Len(str) > fixLen Then
        ZLCE_StrFitChar = str
        Exit Function
    End If
    
    '����String
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
    '���ܣ�ģ��Oracle��Decode����
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


'1.0.6 2022.08.23 ����У����ֵͬ����,Ĭ��false
Public Function ZLCE_StringAppend(ByRef mainStr As String, ByVal SplitStr As String, ByVal appendStr As String, Optional ByVal IsCheckNull As Boolean = False, _
                            Optional ByVal IsTrim As Boolean = False, Optional ByVal IsSideAddSplitStr As Boolean = False, Optional ByVal checkSame As Boolean = False) As String
'���ܣ�׷���ַ������� �Ƿ�trim
    Dim tmpStr As String, splitLen As Integer: splitLen = Len(SplitStr)
    
    '1.У��߼�λ�����splitStr
    If IsSideAddSplitStr Then
         If Left(mainStr, splitLen) <> SplitStr Then
            mainStr = SplitStr & mainStr
         End If
    End If
    
    '1.1.checkSame
    If checkSame Then
        tmpStr = mainStr
        '����Ƿ��ظ�.��side����ֱ�ӱȶ�
        If IsSideAddSplitStr Then
            If InStr(1, tmpStr, SplitStr & appendStr & SplitStr) >= 1 Then
                ZLCE_StringAppend = mainStr
                Exit Function
                '�˳�����
            End If
        Else
            If InStr(1, SplitStr & tmpStr & SplitStr, SplitStr & appendStr & SplitStr) >= 1 Then
                ZLCE_StringAppend = mainStr
                Exit Function
                '�˳�����
            End If
        End If
    End If
    
    '2.�����ַ���
    If IsCheckNull Then
        mainStr = mainStr & IIf((Len(mainStr) >= 1 And Right(mainStr, splitLen) <> SplitStr), SplitStr, "") & appendStr
    Else
        mainStr = mainStr & SplitStr & appendStr
    End If

    
    '3.Trim
    If IsTrim Then
        mainStr = ZLCE_Trim(mainStr)
    End If
    
    '4.�߼�����splitStr
    If IsSideAddSplitStr Then
        mainStr = mainStr & SplitStr
    End If
    
    ZLCE_StringAppend = mainStr
End Function

Public Function ZLCE_SplitIndex(ByVal str As String, ByVal SplitStr As String, ByVal index As Integer) As String
On Error GoTo ErrH
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�������ȥ�����˵Ŀո�
  ZLCE_SplitIndex = Split(str, SplitStr)(index)
  Exit Function
ErrH:
  ZLCE_SplitIndex = ""
End Function

Public Function ZLCE_ContainSubStr(ByVal mainStr As String, ByVal subStr As String, Optional ByVal SplitStr As String = ",") As Boolean
On Error GoTo ErrH
'���ܣ���ѯ�ַ����ǲ��������ַ����г���
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

