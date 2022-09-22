Attribute VB_Name = "ZLmdlDictionary"
Option Explicit

Public Function ZLCE_JSONParse(ByVal JSONString As String, ByVal JSONPath As String) As Variant
On Error GoTo ErrH
    Dim JSON As Object
    Set JSON = CreateObject("MSScriptControl.ScriptControl")
    JSON.Language = "JScript"
    ZLCE_JSONParse = JSON.eval("JSON=" & JSONString & ";JSON." & JSONPath & ";")
    Set JSON = Nothing
    Exit Function
ErrH:
    Err.Clear
    ZLCE_JSONParse = ""
    Set JSON = Nothing
End Function
