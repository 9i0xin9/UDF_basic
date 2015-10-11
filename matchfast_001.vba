'Sub matchfast001() 'Funziona velocemente
'Dim vArr As Variant
'Dim j As Long
'Dim N As Long
'Dim dTime As Double
'vArr = rngM(0).Value2
'For j = LBound(vArr) To UBound(vArr)
'If vArr(j, 1) = "X" Then
'If vArr(j, 1) = "Y" Then
'N = N + 1
'End If
'End If
'Next j
'[C1].Value = (MicroTimer - dTime) * 1000 & " number:" & N
'End Sub

'***********************************************************
'Questa la base ricavata da https://fastexcel.wordpress.com/2011/10/26/match-vs-find-vs-variant-array-vba-performance-shootout/
'L'idea adesso e' quella di ricreare un CountIfs personalizzato

Public Function countIfs_speed(ParamArray rngs() As Variant)

'0      Empty (unitialized)
'1      Null (no valid data)
'2      Integer
'3      Long Integer
'4      Single
'5      Double
'6      Currency
'7      Date
'8      String
'9      Object
'10     Error Value
'11     Boolean
'12     Variant (only used with arrays of variants)
'13     Data access object
'14     Decimal value
'17     Byte
'36     User Defined Type
'8204   Range
'8192   Array

Dim arr1(), arr2()
Dim k As Long, lcount1 As Long, lcount2 As Long

For k = LBound(rngs) To UBound(rngs)

 'MsgBox IsArray(rngs(k)) & VarType(rngs(k))
 'On Error Resume Next
 'MsgBox rngs(k).Address
 'If Not rngs(k) Is Nothing Then
 
 If VarType(rngs(k)) = 8204 Then
'  If UBound(rngs(k)) > 0 Then
   ReDim Preserve arr1(lcount1)
   arr1(lcount1) = rngs(k)    'Compilo l'array
   lcount1 = lcount1 + 1
 Else
 If VarType(rngs(k)) < 2 Then Exit Function
   ReDim Preserve arr2(lcount2)
   arr2(lcount2) = rngs(k)    'Compilo l'array
   lcount2 = lcount2 + 1
 End If
' End If
 
Next k

MsgBox "Final " & UBound(arr1) & " " & UBound(arr2)

End Function
