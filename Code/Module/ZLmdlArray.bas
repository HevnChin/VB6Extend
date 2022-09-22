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

Rem 获取最小index
Public Function ZLCE_GetArrayMinIndex(ByRef arr() As Variant) As Long
On Error GoTo ErrH
    ZLCE_GetArrayMinIndex = LBound(arr)
    Exit Function
ErrH:
    ZLCE_GetArrayMinIndex = 0
End Function

Rem 获取最大index
Public Function ZLCE_GetArrayMaxIndex(ByRef arr() As Variant) As Long
On Error GoTo ErrH
    ZLCE_GetArrayMaxIndex = UBound(arr)
    Exit Function
ErrH:
    ZLCE_GetArrayMaxIndex = 0
End Function

Rem 获取Count
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
''插入数据: 默认查到末尾
Public Function ZLCE_Str_ArrayInsertIndex(ByRef arr() As String, ByVal str As String, Optional index As Integer = -1) As Boolean
On Error GoTo ErrH
    Dim minIndex As Integer, maxIndex As Integer, i As Integer
    Dim tmpArray() As String
    minIndex = LBound(arr): maxIndex = UBound(arr)
    
    '1.追加-LastAdd
    If index = -1 Then
         ReDim Preserve arr(minIndex To maxIndex + 1)  'Preserve：当改变原有数组最末维的大小时，使用此关键字可以保持数组中原来的数据。
         arr(maxIndex + 1) = str
         ZLCE_Str_ArrayInsertIndex = True
         Exit Function
    End If
    
    
    '2.如果是插入的话, 首先复制一个数组出来
    ReDim tmpArray(minIndex To maxIndex) As String
    For i = minIndex To maxIndex Step 1  '数组里面
            tmpArray(i) = arr(i)
    Next i
     
     maxIndex = maxIndex + 1
    '2.1前加 InsertHead
    ReDim Preserve arr(minIndex To maxIndex)
    
    If minIndex = index Then
        arr(minIndex) = str
        For i = minIndex + 1 To maxIndex Step 1  '数组里面
            arr(i) = tmpArray(i - 1)
        Next i
    Else
        '2.2 中间加
        For i = minIndex To maxIndex Step 1   '数组里面
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
    
    '1.追加-LastAdd
    If index = -1 Then
         ReDim Preserve arr(minIndex To maxIndex + 1)  'Preserve：当改变原有数组最末维的大小时，使用此关键字可以保持数组中原来的数据。
         arr(maxIndex + 1) = var
         ZLCE_ArrayInsertIndex = True
         Exit Function
    End If
    
    
    '2.如果是插入的话, 首先复制一个数组出来
    ReDim tmpArray(minIndex To maxIndex) As Variant
    For i = minIndex To maxIndex Step 1  '数组里面
            tmpArray(i) = arr(i)
    Next i
     
     maxIndex = maxIndex + 1
    '2.1前加 InsertHead
    ReDim Preserve arr(minIndex To maxIndex)
    
    If minIndex = index Then
        arr(minIndex) = var
        For i = minIndex + 1 To maxIndex Step 1  '数组里面
            arr(i) = tmpArray(i - 1)
        Next i
    Else
        '2.2 中间加
        For i = minIndex To maxIndex Step 1   '数组里面
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

