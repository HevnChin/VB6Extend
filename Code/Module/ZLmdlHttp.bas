Attribute VB_Name = "ZLmdlHttp"
Option Explicit

Public Function ZLCE_WinHttpRequest5_1() As Object
On Error GoTo ErrH
    Set ZLCE_WinHttpRequest5_1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Exit Function
ErrH:
    MsgBox Err.Description, vbCritical, ZLCE_Nvl(ZLCE_SysName, "VB6Extend")
    Set ZLCE_WinHttpRequest5_1 = Null
    Err.Clear
End Function

'1.0.8   HttpRequestType
Public Function ZLCE_XMLHTTPRequest(ByVal reqURL As String, ByVal reqContent As String, ByVal httpReqType As HttpRequestType, _
                                                                                    Optional Method As String = "POST", Optional reqKeyValues As Dictionary = Nothing) As String
On Error GoTo ErrH
    Dim oXMLHTTP  As MSXML2.XMLHTTP  'As Object
    Dim HttpRequest As String, vKey As Variant
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    
        '������ڼ�ֵ��
    If reqKeyValues Is Nothing Then
    Else
          For Each vKey In reqKeyValues
              oXMLHTTP.setRequestHeader vKey, reqKeyValues(vKey)
          Next
    End If
  
    '----�����ڼ䲻���н���-------
    oXMLHTTP.Open Method, reqURL, False
'    oXMLHTTP.setRequestHeader "aAccept-Encoding", "gzip,deflate"
'    oXMLHTTP.setRequestHeader "Content-Type", "text/XML;charset=UTF-8"
'    oXMLHTTP.setRequestHeader "SOAPAction", "http://soap.jkgs.gov.cn/PayRefundNew"
    oXMLHTTP.setRequestHeader "Content-Length", Len(reqContent)
'    oXMLHTTP.setRequestHeader "Connection", "Keep-Alive"
'    oXMLHTTP.setRequestHeader "Host", "10.85.40.76:8083"
'    oXMLHTTP.setRequestHeader "User-Agent", "Apache-HttpClient/4.1.1 (java 1.5)"
    oXMLHTTP.send reqContent
    
    Do Until oXMLHTTP.readyState = 4
        DoEvents
    Loop
  
    '--------------------------------��������
    Select Case httpReqType
      Case HttpRequestType_XML
        '--------------------------------ֱ�ӷ���XML
          HttpRequest = oXMLHTTP.responseXML
      Case HttpRequestType_Text
        '--------------------------------ֱ�ӷ����ַ���
        HttpRequest = oXMLHTTP.responseText
      Case HttpRequestType_Body
        '--------------------------------ֱ�ӷ��ض�����
        HttpRequest = oXMLHTTP.responseBody
      Case HttpRequestType_BodyText
        '--------------------------------������ת�ַ���[ֱ�ӷ����ִ���������ʱ����]
        HttpRequest = ZLCE_BytesToStr(oXMLHTTP.responseBody)
      Case Else
        '--------------------------------��Ч�ķ���
        HttpRequest = ""
    End Select
    
    ZLCE_XMLHTTPRequest = HttpRequest
    '--------------------------------�ͷſռ�
    Set oXMLHTTP = Nothing
    Exit Function
ErrH:
    Set oXMLHTTP = Nothing
    Err.Clear
End Function


Private Function ZLCE_BytesToStr(ByVal vIn) As String
On Error GoTo ErrH
    Dim strReturn As String: strReturn = ""
    
    Dim i As Integer, ThisCharCode As String, NextCharCode As String
    
    For i = 1 To LenB(vIn)
        ThisCharCode = AscB(MidB(vIn, i, 1))
        If ThisCharCode < &H80 Then
            strReturn = strReturn & Chr(ThisCharCode)
        Else
            NextCharCode = AscB(MidB(vIn, i + 1, 1))
            strReturn = strReturn & Chr(CLng(ThisCharCode) * &H100 + CInt(NextCharCode))
            i = i + 1
        End If
    Next
    
    ZLCE_BytesToStr = strReturn
    Exit Function
ErrH:
    ZLCE_BytesToStr = ""
    Err.Clear
End Function
