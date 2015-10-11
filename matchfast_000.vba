'https://fastexcel.wordpress.com/2011/10/26/match-vs-find-vs-variant-array-vba-performance-shootout/
'Reply
'Jim Cone says:
'October 27, 2011 at 4:34 am
'Charles,
'I have some rather long Find code to do a multi sheet, multi word search and it completes generally before the start button 'rebounds.
'So your test results showing that Match beats Find surprised me.
'I tried your code on 50,000 rows of data and got approximately the same relative results that you did. I am going to 'continue reading your posts.

'Another surprise was the speed increase when I replaced your range.resize code line with a different range callout.
'This came up about 18% faster…

'!!!!!!ATTENZIONE stiamo cercando il match tra colonna A e B quando A contiene "X" e nella stessa riga B contiene "Y" il risultato n è il numero di presenze
'!!!!!!Provare a capire l'uso delle specialCells per velocizzare i conti?

'— code starts
Sub FindXY333()
Dim oRng As Range
Dim oLastRng As Range
Dim j As Long
Dim n As Long
Dim Rw As Long
Dim dTime As Long  'calcolo prestazione

dTime = timeGetTime  'calcolo prestazione
Set oRng = Range("a1:A50000")  
Set oLastRng = oRng(oRng.Rows.Count) 'Questo conta il numero di celle del nuovo range (diminuisce ad ogni match trovato)
'se A2:B3 lui restituisce 2 come su A4:A5 = oRng(50000) dove il valore della colonna omesso corrisponde ad 1
Rw = oLastRng.Row
On Error GoTo Finish
With Application.WorksheetFunction
Do
'originale riga era  Set oRng = Range(oRng(j+1),oLastRng) '<<<= Rw provare con un address a vedere cosa risulta
'prima prova fattaSet oRng = Range(Cells(oRng(j + 1), 1),Cells(Rw,2)) '<<<= Rw impostiamo il nuovo range che alla partenza 
'sarà (0+1,50000)
Set oRng = Range(oRng(j+1),oLastRng(Rw)) 'altra prova il numero delle colonne resta sottointeso 1
j = .Match("X", oRng, 0)  'Match(valueToMatch,arrayToCompare,matchType) risultato è un numero Double della posizione del 
'valueToMatch all'interno dell'array, la posizione del primo elemento è 1, se nessun match è trovato #N/A
jRow = jRow + j 'se ho ben capito dovrebbe restituire la somma righe dei match trovati che resettandosi il range dovrebbe 
'corrispondere alla riga vera del primo range A1:A50000
If oRng(j, 2).Value2 = "Y" Then n = n + 1  'verifica seconda se il valore trovato nel range oRng ha la seconda corrispondenza 
'nella colonna B se si aumenta il contatore n di 1
Loop Until jRow + 1 > Rw 'Esegue il loop fino a quando l ultima corrispondenza trovata +1 è maggiore di Rw (50000?) dovrebbe 
'essere statico all'interno del loop non è presente, controllare! come fa a risultare maggiore se non viene trovata un ultima 
'corrispondenza sull'ultima riga?
End With

Finish:
MsgBox TellMeHowLong(dTime) & vbCr & n & " found"
End Sub
'— code ends

'There were about 25,000 XY occurances.
'I use the Resize property frequently because of its ease of use.
'Wondering if I still should.
'—
'Jim Cone
