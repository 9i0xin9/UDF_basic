#If VBA7 Then
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias _
"QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function GetTickCount Lib "kernel32" Alias _
"QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
Private Declare Function getFrequency Lib "kernel32" Alias _
"QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare Function GetTickCount Lib "kernel32" Alias _
"QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If
Public Function MicroTimer() As Double
'
' returns seconds
' uses Windows API calls to the high resolution timer
'
Dim cyTicks1 As Currency
Dim cyTicks2 As Currency
Static cyFrequency As Currency
'
MicroTimer = 0
'
' get frequency
'
If cyFrequency = 0 Then getFrequency cyFrequency
'
' get ticks
'
GetTickCount cyTicks1
GetTickCount cyTicks2
If cyTicks2 < cyTicks1 Then cyTicks2 = cyTicks1
'
' calc seconds
'
If cyFrequency Then MicroTimer = cyTicks2 / cyFrequency
End Function
