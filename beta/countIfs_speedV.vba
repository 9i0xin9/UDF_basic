'If f_con(arr1(j, 1), con1) *!check if we can obmit the "= True" Then* this operation have to be instored into a private function of comparation to understand with kind of match is searching the function will return a Boolean result.

Private Function f_con(arrV as Variant, con as Variant) as Boolean

'VarType(var) results
'vbEmpty            0 Vuoto (nessuna inizializzazione)
'vbNull             1 Null (nessun dato valido)
'vbInteger          2 Numero intero
'vbLong             3 Intero lungo
'vbSingle           4 Numero a virgola mobile a precisione singola
'vbDouble           5 Numero a virgola mobile a precisione doppia
'vbCurrency         6 Valore di valuta
'vbDate             7 Valore di data
'vbString           8 Valore stringa
'vbObject           9 Oggetto
'vbError           10 Valore di errore
'vbBoolean         11 Valore booleano
'vbVariant         12 Valore Variant (utilizzato solo con matrice di varianti)
'vbDataObject      13 Oggetto DAO (Data Access Object)
'vbDecimal         14 Valore decimale
'vbByte            17 Valore byte
'vbUserDefinedType 36 Varianti che contengono tipi definiti dall'utente
'Omessi gli array

If varType(arrV) = 8 then compareSign = 0'check just for equal comparatorSign 0 or absent

Dim compareSign as Long

'Classic method
If Left(con,2) = ">=" then 
    compareSign = 3
    GoTo endifs
ElseIf Left(con,2) = "<=" then 
    compareSign = 4
    GoTo endifs
ElseIf Left(con,2) = "<>" then 
    compareSign = 5
    GoTo endifs  
ElseIf Find(con,"><") > 0 then 
    compareSign = 6
    GoTo endifs
ElseIf Left(con,1) = ">"  then
    compareSign = 1
    GoTo endifs
ElseIf Left(con,1) = "<"  then
    compareSign = 2
    GoTo endifs
ElseIf Left(con,1) = "="  then 
    compareSign = 0
    GoTo endifs
Else
    compareSign = 0
End If
Endifs:

'Faster method
compareSign = asc(Left(con,1))  &  asc(Mid(con,2,1))  '074 < '075 = '076 >
'>= 76075
'<= 74075
'<> 74076
'>< 76074
'=  75 (000-177)
'>  76 (000-177)
'<  74 (000-177)
' controllare come selezionare il caso specifico rispetto al range, se scriverlo prima o dopo

Select Case compareSign

Case 76075
    If arrV >= Mid(con,3) Then
        f_con = TRUE
    Else
        f_con = FALSE
    End If

Case 74075
    If arrV <= Mid(con,3) Then
        f_con = TRUE
    Else
        f_con = FALSE
    End If

Case 76076
    If arrV <> Mid(con,3) Then
        f_con = TRUE
    Else
        f_con = FALSE
    End If

Case 76074
    Dim correct As Long
    correctV = InStr(con,"|")

    If correctV = 0 Then
        MsgBox "invalid format for between condition Chr(10) ex. "><1.20|1.50"
    End If
    If arrV > Mid(con,3,correctV-3) And arrV < Mid(con,correctV+1) Then
        f_con = TRUE
    Else
        f_con = FALSE
    End If

Case 76000 to 76999
    If arrV > Mid(con,2) Then
        f_con = TRUE
    Else
        f_con = FALSE
    End If

Case 74000 to 74999
    If arrV < Mid(con,2) Then
        f_con = TRUE
    Else
        f_con = FALSE
    End If

Case 75000 to 75999
    If arrV = Mid(con,2) Then
        f_con = TRUE
    Else
        f_con = FALSE
    End If

Else Case
    If arrV = Mid(con,2) Then
        f_con = TRUE
    Else
        f_con = FALSE
    End If

End Select

End Function


Public Function countIfs_speed(rng1 As Range, con1 As Variant, _
Optional rng2 As Range, Optional con2 As Variant, _
Optional rng3 As Range, Optional con3 As Variant, _
Optional rng4 As Range, Optional con4 As Variant, _
Optional rng5 As Range, Optional con5 As Variant, _
Optional rng6 As Range, Optional con6 As Variant, _
Optional rng7 As Range, Optional con7 As Variant, _
Optional rng8 As Range, Optional con8 As Variant, _
Optional rng9 As Range, Optional con9 As Variant, _
Optional rng10 As Range, Optional con10 As Variant, _
Optional rng11 As Range, Optional con11 As Variant, _
Optional rng12 As Range, Optional con12 As Variant, _
Optional rng13 As Range, Optional con13 As Variant)

Dim arr1 As Variant, arr2 As Variant, arr3 As Variant, _
arr4 As Variant, arr5 As Variant, arr6 As Variant, _
arr7 As Variant, arr8 As Variant, arr9 As Variant, _
arr10 As Variant, arr11 As Variant, arr12 As Variant, arr13 As Variant
Dim N As Long

arr1 = rng1.Value2
con1 = f_con(con1) 'must return an array(1) two value stored Byte number (es. 0 for =, 1 for >, 2 for <, 3 for >=, 4 for <=, 5 for <>, 6 for ><) and value to compare (include also a validating process ex. for strings admit only = match

 If rng2 Is Nothing Then
 N = 1
  GoTo outArr
  Else
 End If
arr2 = rng2.Value2

 If rng3 Is Nothing Then
 N = 2
  GoTo outArr
  Else
 End If
arr3 = rng3.Value2

 If rng4 Is Nothing Then
 N = 3
  GoTo outArr
  Else
 End If
arr4 = rng4.Value2

 If rng5 Is Nothing Then
 N = 4
  GoTo outArr
  Else
 End If
arr5 = rng5.Value2

 If rng6 Is Nothing Then
 N = 5
  GoTo outArr
  Else
 End If
arr6 = rng6.Value2

 If rng7 Is Nothing Then
 N = 6
  GoTo outArr
  Else
 End If
arr7 = rng7.Value2

 If rng8 Is Nothing Then
 N = 7
  GoTo outArr
  Else
 End If
arr8 = rng8.Value2

 If rng9 Is Nothing Then
 N = 8
  GoTo outArr
  Else
 End If
arr9 = rng9.Value2

 If rng10 Is Nothing Then
 N = 9
  GoTo outArr
  Else
 End If
arr10 = rng10.Value2

 If rng11 Is Nothing Then
 N = 10
  GoTo outArr
  Else
 End If
arr11 = rng11.Value2

 If rng12 Is Nothing Then
 N = 11
  GoTo outArr
  Else
 End If
arr12 = rng12.Value2

 If rng13 Is Nothing Then
 N = 12
  GoTo outArr
  Else
 End If
arr13 = rng13.Value2
N = 13
outArr:



vArr1 = rng1.Value
'MsgBox N & " " & LBound(vArr1) & " " & UBound(vArr1)
Select Case N

Case 1
For j = LBound(vArr1) To UBound(vArr1)
 If arr1(j, 1) = con1 Then
  N = N + 1
 End If
Next j

Case 2
For j = LBound(vArr1) To UBound(vArr1)
 If arr1(j, 1) = con1 Then
  If arr2(j, 1) = con2 Then
   N = N + 1
  End If
 End If
Next j

Case 3
For j = LBound(vArr1) To UBound(vArr1)
 If arr1(j, 1) = con1 Then  'If f_con(arr1(j, 1), con1) *!check if we can obmit the "= True" Then* this operation have to be instored into a private function of comparation to understand with kind of match is searching the function will return a Boolean result.
  If arr2(j, 1) = con2 Then
   If arr3(j, 1) = con3 Then
    N = N + 1
   End If
  End If
 End If
Next j

Case 4
For j = LBound(vArr1) To UBound(vArr1)
 If arr1(j, 1) = con1 Then
  If arr2(j, 1) = con2 Then
   If arr3(j, 1) = con3 Then
    If arr4(j, 1) = con4 Then
     N = N + 1
    End If
   End If
  End If
 End If
Next j

Case 5
For j = LBound(vArr1) To UBound(vArr1)
 If arr1(j, 1) = con1 Then
  If arr2(j, 1) = con2 Then
   If arr3(j, 1) = con3 Then
    If arr4(j, 1) = con4 Then
     If arr5(j, 1) = con5 Then
      N = N + 1
     End If
    End If
   End If
  End If
 End If
Next j

Case 6
For j = LBound(vArr1) To UBound(vArr1)
 If arr1(j, 1) = con1 Then
  If arr2(j, 1) = con2 Then
   If arr3(j, 1) = con3 Then
    If arr4(j, 1) = con4 Then
     If arr5(j, 1) = con5 Then
      If arr6(j, 1) = con6 Then
       N = N + 1
      End If
     End If
    End If
   End If
  End If
 End If
Next j

Case 7
For j = LBound(vArr1) To UBound(vArr1)
 If arr1(j, 1) = con1 Then
  If arr2(j, 1) = con2 Then
   If arr3(j, 1) = con3 Then
    If arr4(j, 1) = con4 Then
     If arr5(j, 1) = con5 Then
      If arr6(j, 1) = con6 Then
       If arr7(j, 1) = con7 Then
        N = N + 1
       End If
      End If
     End If
    End If
   End If
  End If
 End If
Next j

Case 8
For j = LBound(vArr1) To UBound(vArr1)
 If arr1(j, 1) = con1 Then
  If arr2(j, 1) = con2 Then
   If arr3(j, 1) = con3 Then
    If arr4(j, 1) = con4 Then
     If arr5(j, 1) = con5 Then
      If arr6(j, 1) = con6 Then
       If arr7(j, 1) = con7 Then
        If arr8(j, 1) = con8 Then
         N = N + 1
        End If
       End If
      End If
     End If
    End If
   End If
  End If
 End If
Next j

Case 9
For j = LBound(vArr1) To UBound(vArr1)
 If arr1(j, 1) = con1 Then
  If arr2(j, 1) = con2 Then
   If arr3(j, 1) = con3 Then
    If arr4(j, 1) = con4 Then
     If arr5(j, 1) = con5 Then
      If arr6(j, 1) = con6 Then
       If arr7(j, 1) = con7 Then
        If arr8(j, 1) = con8 Then
         If arr9(j, 1) = con9 Then
          N = N + 1
         End If
        End If
       End If
      End If
     End If
    End If
   End If
  End If
 End If
Next j

Case 10
For j = LBound(vArr1) To UBound(vArr1)
 If arr1(j, 1) = con1 Then
  If arr2(j, 1) = con2 Then
   If arr3(j, 1) = con3 Then
    If arr4(j, 1) = con4 Then
     If arr5(j, 1) = con5 Then
      If arr6(j, 1) = con6 Then
       If arr7(j, 1) = con7 Then
        If arr8(j, 1) = con8 Then
         If arr9(j, 1) = con9 Then
          If arr10(j, 1) = con10 Then
           N = N + 1
          End If
         End If
        End If
       End If
      End If
     End If
    End If
   End If
  End If
 End If
Next j

Case 11
For j = LBound(vArr1) To UBound(vArr1)
 If arr1(j, 1) = con1 Then
  If arr2(j, 1) = con2 Then
   If arr3(j, 1) = con3 Then
    If arr4(j, 1) = con4 Then
     If arr5(j, 1) = con5 Then
      If arr6(j, 1) = con6 Then
       If arr7(j, 1) = con7 Then
        If arr8(j, 1) = con8 Then
         If arr9(j, 1) = con9 Then
          If arr10(j, 1) = con10 Then
           If arr11(j, 1) = con11 Then
            N = N + 1
           End If
          End If
         End If
        End If
       End If
      End If
     End If
    End If
   End If
  End If
 End If
Next j

Case 12
For j = LBound(vArr1) To UBound(vArr1)
 If arr1(j, 1) = con1 Then
  If arr2(j, 1) = con2 Then
   If arr3(j, 1) = con3 Then
    If arr4(j, 1) = con4 Then
     If arr5(j, 1) = con5 Then
      If arr6(j, 1) = con6 Then
       If arr7(j, 1) = con7 Then
        If arr8(j, 1) = con8 Then
         If arr9(j, 1) = con9 Then
          If arr10(j, 1) = con10 Then
           If arr11(j, 1) = con11 Then
            If arr12(j, 1) = con12 Then
             N = N + 1
            End If
           End If
          End If
         End If
        End If
       End If
      End If
     End If
    End If
   End If
  End If
 End If
Next j

Case 13
For j = LBound(vArr1) To UBound(vArr1)
 If arr1(j, 1) = con1 Then
  If arr2(j, 1) = con2 Then
   If arr3(j, 1) = con3 Then
    If arr4(j, 1) = con4 Then
     If arr5(j, 1) = con5 Then
      If arr6(j, 1) = con6 Then
       If arr7(j, 1) = con7 Then
        If arr8(j, 1) = con8 Then
         If arr9(j, 1) = con9 Then
          If arr10(j, 1) = con10 Then
           If arr11(j, 1) = con11 Then
            If arr12(j, 1) = con12 Then
             If arr13(j, 1) = con13 Then
              N = N + 1
             End If
            End If
           End If
          End If
         End If
        End If
       End If
      End If
     End If
    End If
   End If
  End If
 End If
Next j
End Select

countIfs_speed = N

End Function
