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
Rw = oLastRng.Row
On Error GoTo Finish
With Application.WorksheetFunction
Do
'originale riga era  Set oRng = Range(oRng(j+1),oLastRng) '<<<= Rw provare con un address a vedere cosa risulta
Set oRng = Range(Cells(oRng(j + 1), 1),Cells(Rw,2)) '<<<= Rw impostiamo il nuovo range che alla partenza sarà (0+1,50000)
j = .Match("X", oRng, 0)  'Match(valueToMatch,arrayToCompare,matchType) risultato è un numero Double della posizione del valueToMatch all'interno dell'array, la posizione del primo elemento è 1, se nessun match è trovato #N/A
jRow = jRow + j
If oRng(j, 2).Value2 = "Y" Then n = n + 1  'compila il valore se trovato nel range oRng
Loop Until jRow + 1 > Rw
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
