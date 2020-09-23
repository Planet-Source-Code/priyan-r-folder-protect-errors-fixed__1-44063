Attribute VB_Name = "passconv"
Option Explicit

'---------------------------------------------------
'Priyan's password encryption functions
'---------------------------------------
'This is very simpe encryption and it is not so good
'It only converts the charectors to ASCII
'
'---------------------------------------------------
Public Function cvtpasstostring(ByVal str As String) As String
Dim a As String
Dim b As String
Dim i As Integer
Dim j As Integer
j = Len(str)
If str <> 0 Then
For i = 1 To j
        b = Chr(Mid$(str, i, 2))
        a = a & b
        i = i + 1
Next
End If
cvtpasstostring = UCase(a)
End Function
Public Function cvtstringtopass(ByVal str As String) As Variant
str = Trim(str)
str = UCase(str)
Dim a As Variant
Dim b As Integer
Dim i As Integer
If Len(str) <> 0 Then
For i = 1 To Len(str)
        b = Asc(Mid$(str, i, 1))
        a = a & b
Next
End If
cvtstringtopass = a
End Function

