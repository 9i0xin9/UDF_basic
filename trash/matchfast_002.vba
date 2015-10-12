'***********************************************************
'Questa la base ricavata da https://fastexcel.wordpress.com/2011/10/26/match-vs-find-vs-variant-array-vba-performance-shootout/
'L'idea adesso e' quella di ricreare un CountIfs personalizzato

Public Function countIfs_speed2(ParamArray rngs() As Variant)


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



Dim arr1(), arr2(), arrV
Dim k As Long, lcount1 As Long, lcount2 As Long, cntk As Long, cntM As Long, CountO As Long

cntk = 0
'ReDim Preserve arr1(lcount1)
'arr1(lcount1) = rng0
'lcount1 = 1

'ReDim Preserve arr2(lcount2)
'arr2(lcount2) = con0
'lcount2 = 1

For k = LBound(rngs) To UBound(rngs) 'usare Step 2?
 
 If VarType(rngs(k)) = 8204 Then

    ReDim Preserve arr1(lcount1)
    arr1(lcount1) = rngs(k)   'Compilo l'array dei ranges
    ' come compilare l arrV con i valori (ricostruire un redim preserve 2d) se trova altri ranges redim e aggiunge i valori nel nuovo spazio
    ' provare se rngs(k).value per la condizione va bene sicuramente andrà elaborato un parsing per capire se maggiore minore, maggiore uguale minore uguale, compreso, diverso o solo numero
    ' settare un array temporale che viaggia tra i ranges es. se incontro nella prima colonna il match alla riga 11 proseguo
    ' la scansione scaricando l array della prima colonna ed caricando quello della seconda?
    ' rngs(k) = cons(k) cnt=(ubound(rngs)+2)/2 l'idea
    ' creare un ciclo che confronta la prima colonna, se trova un match (for each elements) passa all'array successivo '.
    ' cercando dalla stessa riga una corrispondenza ?capire se possiamo ricavare la posizione della riga nel ciclo for each
    '
     For i = LBound(rngs(k)) To UBound(rngs(k)) 'possibile che non funzioni nel caso passare l array ad una variabile
    
     cntM = 0
recount:

'MsgBox cntk
     ReDim arrV(UBound(rngs(cntk)))
     arrV = rngs(cntk)
    'MsgBox arrV(1, 1)
'MsgBox "arrV:" & arrV(i, 1) & " rngs:" & UBound(arrV) & "  cntk:" & rngs(cntk + 1) & "  i:" & i
    
     If arrV(i, 1) = rngs(cntk + 1) Then 'controlla se esegue il primo match
'MsgBox arrV(i, 1) & " " & rngs(cntk + 1)
     Else
     GoTo fine
     End If
     
     If cntk = LBound(rngs) Then
     cntk = cntk + 2 'in caso di un solo if non funzionerebbe
     GoTo end_cntk
     Else
     End If
'MsgBox UBound(rngs) - 1

     If cntk = UBound(rngs) - 1 Then cntk = cntk - 2
end_cntk:
     cntM = cntM + 1

     If cntM = (UBound(rngs) + 1) / 2 Then
     CountO = CountO + 1
          GoTo fine
     Else
          GoTo recount
     End If
    ' End If
    ' End If
    ' End If

fine:
     Next i

  lcount1 = lcount1 + 1

  Else

  If VarType(rngs(k)) < 2 And VarType(rngs(k)) <> 8192 Then GoTo errorMsg_comp 'se vuoti o array vai al messaggio di errore

    ReDim Preserve arr2(lcount2)
    arr2(lcount2) = rngs(k)    'Compilo l'array delle condizioni
    lcount2 = lcount2 + 1

 End If
 
Next k

'If UBound(arr1) <> UBound(arr2) Then GoTo errorMsg_comp 'se arr1 e arr2 lunghi uguali procedi con la funzione altrimenti vai al messaggio di errrore

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


MsgBox CountO / 2

errorMsg_comp:
'MsgBox "i parametri inseriti non sono corretti"
'MsgBox Join(arr1, vbTab) & Chr(10)
'& Join(arr2, vbTab)
End Function
