Attribute VB_Name = "ZLMdlDate"
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Function ZLCE_DateGetUnixTimeStamp(Optional strDate As String = "", Optional isMillis As Boolean = False) As String
On Error GoTo ErrH
    If Len(CStr(strDate)) <= 0 Then
        strDate = CStr(Now)
    End If
    ZLCE_DateGetUnixTimeStamp = DateDiff("S", "1970-01-01 00:00:00", DateAdd("h", -8, CDate(strDate))) & IIf(isMillis, Right(timeGetTime, 3), "")
    Exit Function
ErrH:
    ZLCE_DateGetUnixTimeStamp = ""
End Function
