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

'— code starts
Sub FindXY333()
Dim oRng As Range
Dim oLastRng As Range
Dim j As Long
Dim n As Long
Dim Rw As Long
Dim dTime As Long

dTime = timeGetTime
Set oRng = Range("a1:A50000")
Set oLastRng = oRng(oRng.Rows.Count)
Rw = oLastRng.Row
On Error GoTo Finish
With Application.WorksheetFunction
Do
j = .Match("X", oRng, 0)
If oRng(j, 2).Value2 = "Y" Then n = n + 1
Set oRng = Range(oRng(j + 1), oLastRng) '<<<= Rw
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
