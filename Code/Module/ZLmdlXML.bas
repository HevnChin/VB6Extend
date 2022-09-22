Attribute VB_Name = "ZLmdlXML"
Option Explicit
'ȫ��XML����
'Public gstrOut As String * 4000, strOutXml As String,
Public ZLCE_RepDoc As Object
'Public gOutErrMsg As String
'������Դ
Public Const XMLErrorNum = vbObjectError + 3
 
Public Function ZLCE_GetElemnetValue(ByVal doc As Object, ByVal name As String, Optional ByVal itemIndex As Integer = 0, Optional ByVal IsOrion As Boolean = False) As String
'���ܣ��õ�ָ��Ԫ�ص�ֵ
    Dim xmlElement As Object 'MSXML2.IXMLDOMNodeList
    If IsOrion = False Then
        name = LCase(name)
    End If
    
    Set xmlElement = doc.documentElement.getElementsByTagName(name)
    If (Not xmlElement Is Nothing) And xmlElement.Length >= 1 Then
        '�ҵ�ָ����Ԫ��
        ZLCE_GetElemnetValue = xmlElement.Item(itemIndex).Text
        Exit Function
    End If
End Function


 Public Function ZLCE_GetLoadXMLObj(ByVal strXML As String) As Object
'��ȡ���Ҽ���XML����
On Error GoTo ErrH
    If IsNull(ZLCE_RepDoc) Or ZLCE_RepDoc Is Nothing Then
        Set ZLCE_RepDoc = CreateObject("MSXML2.DOMDocument")
        'Else
        'Set ZLCE_RepDoc = GetObject("", "MSXML2.DOMDocument")
    End If
    
    If ZLCE_RepDoc.loadXML(strXML) = False Then
        Err.Raise XMLErrorNum, "", "У��XML������ʽ����ȷ�����飡" & vbNewLine & strXML
    End If
     Set ZLCE_GetLoadXMLObj = ZLCE_RepDoc
    Exit Function
ErrH:
    Set ZLCE_GetLoadXMLObj = Null
End Function

'��XML��ֵ�ӿ�
Public Function ZLCE_GetXMLNode(ByVal doc As Object, ByVal key As String) As String
On Error GoTo ErrH
    ZLCE_GetXMLNode = ZLCE_GetElemnetValue(doc, key, 0, True)
    Exit Function
ErrH:
    ZLCE_GetXMLNode = ""
End Function

Public Function ZLCE_GetXMLSingleNode(ByVal xmlDoc As Object, ByVal keyPath As String) As String
'get Single Node
On Error GoTo ErrH
    ZLCE_GetXMLSingleNode = xmlDoc.selectSingleNode(keyPath).Text
    Exit Function
ErrH:
    ZLCE_GetXMLSingleNode = ""
End Function

'XMLString
Public Function ZLCE_GetXMLStrNode(ByVal strXML As String, ByVal key As String) As String
'��XML��ֵ�ӿ�
On Error GoTo ErrH
     Call ZLCE_GetLoadXMLObj(strXML)
    ZLCE_GetXMLStrNode = ZLCE_GetElemnetValue(ZLCE_RepDoc, key, 0, True)
    Exit Function
ErrH:
    ZLCE_GetXMLStrNode = ""
End Function

Public Function ZLCE_GetXMLStrSingleNode(ByVal strXML As String, ByVal keyPath As String) As String
'get Single Node
On Error GoTo ErrH
    Call ZLCE_GetLoadXMLObj(strXML)
    ZLCE_GetXMLStrSingleNode = ZLCE_RepDoc.selectSingleNode(keyPath).Text
    Exit Function
ErrH:
    ZLCE_GetXMLStrSingleNode = ""
End Function

'
Public Function ZLCE_XMLSetKeyValue(ByVal xmlPath As String, ByVal key As String, ByVal value As String, Optional index As Integer = 0) As String
'�޸�XML����ֵ
On Error GoTo ErrH
    Dim doc As Object, ret As String, xmlElement As Object
    
    Set doc = CreateObject("MSXML2.DOMDocument")
    doc.async = False
    doc.loadXML xmlPath
    
    Set xmlElement = doc.documentElement.getElementsByTagName(key)
    If Not xmlElement Is Nothing Then
        '�ҵ�ָ����Ԫ��
        xmlElement.Item(index).Text = value
    End If
    
    ZLCE_XMLSetKeyValue = doc.xml
    Set doc = Nothing
     Exit Function
ErrH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
    Set doc = Nothing
    ZLCE_XMLSetKeyValue = xmlPath
End Function


'����Index��ȡ�ڵ�����XML
Public Function ZLCE_XMLGetNodeByListIndex(ByVal doc As Object, ByVal ListKey As String, ByRef valueArray() As String, Optional subListKey As String = "", Optional idx As Integer = -1) As String
 On Error GoTo ErrH
    'Dim ItemNodes As MSXML2.IXMLDOMNodeList, subItem As IXMLDOMElement, Items As MSXML2.IXMLDOMNode, ItemNode As MSXML2.IXMLDOMNode
    '=================
    Dim ItemNodes As Object, subItem As Object, Items As Object, ItemNode As Object
    Dim itemKey As Variant, itemValue As Variant
     
    Set Items = doc.selectSingleNode(ListKey)
    '-------------------------------------------------------
    Dim index As Integer, subIndex As Integer
    Set ItemNodes = Items.childNodes
    Erase valueArray
    index = 0
    For Each ItemNode In ItemNodes
        If Len(CStr(subListKey)) >= 1 Then
        subIndex = 0
        '>>>--------------------�������Ҫ���ҵ���Key-------------------->>>
        For subIndex = 0 To ItemNode.childNodes.Length - 1
            Set subItem = ItemNode.childNodes(subIndex)
        
            ReDim Preserve valueArray(index)
            itemKey = subItem.nodeName
            itemValue = subItem.Text
            If itemKey = subListKey Then
                If idx = index Then
                    ZLCE_XMLGetNodeByListIndex = itemValue
                End If
                '----------------------------ֱ�Ӵ洢SubItem�ڵ�------------------
                ReDim Preserve valueArray(index): valueArray(index) = itemValue
                index = index + 1
            End If
            Next
        Else
            '----------------------------ֱ�Ӵ洢XML�ڵ�------------------
            itemValue = ItemNode.xml
            ReDim Preserve valueArray(index)
            valueArray(index) = itemValue
            If index = idx Then
                ZLCE_XMLGetNodeByListIndex = itemValue
            End If
             index = index + 1
        End If
        '<<<--------------------�������Ҫ���ҵ���Key--------------------<<<
    Next
     Exit Function
ErrH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
    ZLCE_XMLGetNodeByListIndex = ""
End Function
