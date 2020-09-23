Attribute VB_Name = "ConvertBINARY"
'***********************************
' Convert Long integers to binary
'***********************************
'  usage:   F$ = Bin(Number)
'
' for MP3, MPG, AVI  encoding etc.
'
' Written by Rizwan Tahir
' http://www.rizwantahir.8m.com
'
'----------------------------------------------
Option Explicit

Public Function BIN(ByVal x As Long) As String

Dim temp As String

temp = ""
'start translation to binary
Do


' Check whether it is 1 bit or 0 bit
If x Mod 2 Then
      temp = "1" + temp
Else
      temp = "0" + temp
End If

x = x \ 2
'  Normal division     7/2 = 3.5
' Integer division     7\2 = 3
'

If x < 1 Then Exit Do

Loop '
BIN = temp

End Function

