Attribute VB_Name = "ZLmdlArray"
Option Explicit

Public Function ZLCE_Str_GetArrayByIndex(ByRef arr() As String, ByVal index As Integer) As String
On Error GoTo ErrH
    ZLCE_Str_GetArrayByIndex = arr(index)
    Exit Function
ErrH:
    ZLCE_Str_GetArrayByIndex = ""
End Function

Public Function ZLCE_GetArrayByIndex(ByRef arr() As Variant, ByVal index As Integer) As Variant
On Error GoTo ErrH
    ZLCE_GetArrayByIndex = arr(index)
    Exit Function
ErrH:
    ZLCE_GetArrayByIndex = ""
End Function

Rem ��ȡ��Сindex
Public Function ZLCE_GetArrayMinIndex(ByRef arr() As Variant) As Long
On Error GoTo ErrH
    ZLCE_GetArrayMinIndex = LBound(arr)
    Exit Function
ErrH:
    ZLCE_GetArrayMinIndex = 0
End Function

Rem ��ȡ���index
Public Function ZLCE_GetArrayMaxIndex(ByRef arr() As Variant) As Long
On Error GoTo ErrH
    ZLCE_GetArrayMaxIndex = UBound(arr)
    Exit Function
ErrH:
    ZLCE_GetArrayMaxIndex = 0
End Function

Rem ��ȡCount
Public Function ZLCE_GetArrayCount(ByRef arr() As Variant) As Long
'ReDim Preserve arr(n - 1)
On Error GoTo ErrH
    ZLCE_GetArrayCount = UBound(arr) - LBound(arr) + 1
    Exit Function
ErrH:
    ZLCE_GetArrayCount = 0
End Function

Rem 1.0.8
Public Function ZLCE_GetStrArrayCount(ByRef arr() As String) As Long
'ReDim Preserve arr(n - 1)
On Error GoTo ErrH
    ZLCE_GetStrArrayCount = UBound(arr) - LBound(arr) + 1
    Exit Function
ErrH:
    ZLCE_GetStrArrayCount = 0
End Function

'1.0.8
''��������: Ĭ�ϲ鵽ĩβ
Public Function ZLCE_Str_ArrayInsertIndex(ByRef arr() As String, ByVal str As String, Optional index As Integer = -1) As Boolean
On Error GoTo ErrH
    Dim minIndex As Integer, maxIndex As Integer, i As Integer
    Dim tmpArray() As String
    minIndex = LBound(arr): maxIndex = UBound(arr)
    
    '1.׷��-LastAdd
    If index = -1 Then
         ReDim Preserve arr(minIndex To maxIndex + 1)  'Preserve�����ı�ԭ��������ĩά�Ĵ�Сʱ��ʹ�ô˹ؼ��ֿ��Ա���������ԭ�������ݡ�
         arr(maxIndex + 1) = str
         ZLCE_Str_ArrayInsertIndex = True
         Exit Function
    End If
    
    
    '2.����ǲ���Ļ�, ���ȸ���һ���������
    ReDim tmpArray(minIndex To maxIndex) As String
    For i = minIndex To maxIndex Step 1  '��������
            tmpArray(i) = arr(i)
    Next i
     
     maxIndex = maxIndex + 1
    '2.1ǰ�� InsertHead
    ReDim Preserve arr(minIndex To maxIndex)
    
    If minIndex = index Then
        arr(minIndex) = str
        For i = minIndex + 1 To maxIndex Step 1  '��������
            arr(i) = tmpArray(i - 1)
        Next i
    Else
        '2.2 �м��
        For i = minIndex To maxIndex Step 1   '��������
            If i < index Then
                arr(i) = tmpArray(i)
            ElseIf i = index Then
                arr(i) = str
             Else
                arr(i) = tmpArray(i - 1)
            End If
        Next i
    End If
    
    ZLCE_Str_ArrayInsertIndex = True
    Exit Function
ErrH:
    ZLCE_Str_ArrayInsertIndex = False
End Function

Public Function ZLCE_ArrayInsertIndex(ByRef arr() As Variant, ByVal var As Variant, Optional index As Integer = -1) As Boolean
On Error GoTo ErrH
    Dim minIndex As Integer, maxIndex As Integer, i As Integer
    Dim tmpArray() As Variant
    minIndex = LBound(arr): maxIndex = UBound(arr)
    
    '1.׷��-LastAdd
    If index = -1 Then
         ReDim Preserve arr(minIndex To maxIndex + 1)  'Preserve�����ı�ԭ��������ĩά�Ĵ�Сʱ��ʹ�ô˹ؼ��ֿ��Ա���������ԭ�������ݡ�
         arr(maxIndex + 1) = var
         ZLCE_ArrayInsertIndex = True
         Exit Function
    End If
    
    
    '2.����ǲ���Ļ�, ���ȸ���һ���������
    ReDim tmpArray(minIndex To maxIndex) As Variant
    For i = minIndex To maxIndex Step 1  '��������
            tmpArray(i) = arr(i)
    Next i
     
     maxIndex = maxIndex + 1
    '2.1ǰ�� InsertHead
    ReDim Preserve arr(minIndex To maxIndex)
    
    If minIndex = index Then
        arr(minIndex) = var
        For i = minIndex + 1 To maxIndex Step 1  '��������
            arr(i) = tmpArray(i - 1)
        Next i
    Else
        '2.2 �м��
        For i = minIndex To maxIndex Step 1   '��������
            If i < index Then
                arr(i) = tmpArray(i)
            ElseIf i = index Then
                arr(i) = var
             Else
                arr(i) = tmpArray(i - 1)
            End If
        Next i
    End If
    
    ZLCE_ArrayInsertIndex = True
    Exit Function
ErrH:
    ZLCE_ArrayInsertIndex = False
End Function

