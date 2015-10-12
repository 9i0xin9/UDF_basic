
'Sub matchfast001() 'Funziona velocemente
'Dim vArr As Variant
'Dim j As Long
'Dim N As Long
'Dim dTime As Double
'vArr = range("A1:B50000").Value2
'For j = LBound(vArr) To UBound(vArr)
'If vArr(j, 1) = "X" Then
'If vArr(j, 2) = "Y" Then
'N = N + 1
'End If
'End If
'Next j
'[C1].Value = (MicroTimer - dTime) * 1000 & " number:" & N
'End Sub

'***********************************************************
'Questa la base ricavata da https://fastexcel.wordpress.com/2011/10/26/match-vs-find-vs-variant-array-vba-performance-shootout/
'L'idea adesso e' quella di ricreare un CountIfs personalizzato

Public Function countIfs_speed(rng0() as Variant, con0 as Variant, ParamArray rngs() As Variant)

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

ReDim Preserve arr1(lcount1)
arr1(lcount1) = rng0
lcount1 = 1

ReDim Preserve arr2(lcount2)
arr2(lcount2) = con0
lcount2 = 1

For k = LBound(rngs) To UBound(rngs) 'usare Step 2?
 
 If VarType(rngs(k)) = 8204 Then

    ReDim Preserve arr1(lcount1)
    arr1(lcount1) = rngs(k)   'Compilo l'array dei ranges
    ' come compilare l arrV con i valori (ricostruire un redim preserve 2d) se trova altri ranges redim e aggiunge i valori nel nuovo spazio
    ' provare se rngs(k).value per la condizione va bene sicuramente andrà elaborato un parsing per capire se maggiore minore, maggiore uguale minore uguale, compreso, diverso o solo numero
    ' settare un array temporale che viaggia tra i ranges es. se incontro nella prima colonna il match alla riga 11 proseguo 
    ' la scansione scaricando l array della prima colonna ed caricando quello della seconda?
    ' rngs(k) = cons(k) cnt=(ubound(rngs)+2)/2 l'idea
    lcount1 = lcount1 + 1

  Else

  If VarType(rngs(k)) < 2 and VarType(rngs(k)) <> 8192 Then errorMsg-comp 'se vuoti o array vai al messaggio di errore

    ReDim Preserve arr2(lcount2)
    arr2(lcount2) = rngs(k)    'Compilo l'array delle condizioni
    lcount2 = lcount2 + 1

 End If
 
Next k

If Ubound(arr1) <> UBound(arr2) then goto errorMsg-comp 'se arr1 e arr2 lunghi uguali procedi con la funzione altrimenti vai al messaggio di errrore

'************************************************************************************
'Adesso dobbiamo settare l'array 2d con i valori dei ranges o forse l abbiamo settato alla compilazione
'es. Dim arrV(1 to maxCellaRanges, 1 to Ubound(arr1)+1))
'
'
'for v = 0 to maxCellaRanges-1
' for h = Lbound(arr1) to Ubound(arr1)
'  arrV(v,h+1) = arr1(h).value2 'qua dovrebbero venir abbinate le voci ricordarsi di scartare se stringa e se numerico per le condizioni diverse
'es. se stringa con > o < e numero allora nel range filtriamo dei numeri altrimenti delle stringhe
'!!! la condizione di verifica deve avvenire all'interno di questo ciclo con la compilazione
' obbligatoriamente quindi dobbiamo inserire il ciclo dei range all interno del ciclo celle (prima v e dopo h)
' next h
'next v
'************************************************************************************



Exit Function
errorMsg-comp:
MsgBox "i parametri inseriti non sono corretti"
MsgBox join(arr1, vbTab) & chr(10) _
& join(arr2, vbTab)
