Attribute VB_Name = "ConvertSnmpVal"
'Author: Petr Domorazek
'E-mail: dsoft@fcanet.cz
'http://dsoft.php5.cz

Option Explicit

Function ConvertSnmpValue(SnmpValueType As Byte, SnmpValueString As String) As String
Dim i As Integer
Select Case SnmpValueType
    Case 2 'Integer
        ConvertSnmpValue = MyConvertInteger(SnmpValueString)
    Case 4 'Octet String
        ConvertSnmpValue = ConvertOctetString(SnmpValueString)
    Case 5 'Null
        ConvertSnmpValue = "Null"
    Case 6 'OID
        ConvertSnmpValue = DeOID(SnmpValueString)
    Case &H40 'IP Address
        ConvertSnmpValue = MyConvertIPAddress(SnmpValueString)
    Case &H41 'Counter
        ConvertSnmpValue = MyConvertNumber(SnmpValueString)
    Case &H42 'Gauge
        ConvertSnmpValue = MyConvertNumber(SnmpValueString)
    Case &H43 'TimeTicks
        ConvertSnmpValue = ConvertTimeTicks(SnmpValueString)
    Case Else 'Other values
        ConvertSnmpValue = ConvertOctetString(SnmpValueString)
End Select
End Function

Private Function ConvertTimeTicks(OctString As String) As String
Dim PStr As String
Dim Vtime As Double
Dim i As Integer
        ConvertTimeTicks = "&H"
        For i = 1 To Len(OctString)
            PStr = Hex(Asc(Mid(OctString, i, 1)))
            If Len(PStr) = 1 Then PStr = "0" & PStr
            ConvertTimeTicks = ConvertTimeTicks & PStr
            Next
        Vtime = Val(ConvertTimeTicks)
        ConvertTimeTicks = "(" & CStr(Vtime) & ") " & CStr(Vtime \ 8640000) & " day(s), "
        Vtime = Vtime Mod 8640000
        ConvertTimeTicks = ConvertTimeTicks & CStr(Vtime \ 360000) & ":"
        Vtime = Vtime Mod 360000
        ConvertTimeTicks = ConvertTimeTicks & CStr(Vtime \ 6000) & ":"
        Vtime = Vtime Mod 6000
        ConvertTimeTicks = ConvertTimeTicks & CStr(Vtime \ 100) & "."
        Vtime = Vtime Mod 100
        ConvertTimeTicks = ConvertTimeTicks & CStr(Vtime)
End Function

Private Function ConvertOctetString(OctString As String) As String
Dim PStr As String
Dim i As Integer
Dim H As Boolean

H = False
For i = 1 To Len(OctString)
    If Asc(Mid(OctString, i, 1)) < 32 Then
        H = True
        Exit For
        End If
    Next

If H = True Then
    ConvertOctetString = StringToHex(OctString)
Else
    ConvertOctetString = OctString
    End If
End Function

Private Function MyConvertIPAddress(OctString As String) As String
Dim PStr As String
Dim i As Integer
    For i = 1 To Len(OctString)
        MyConvertIPAddress = MyConvertIPAddress & Asc(Mid(OctString, i, 1)) & "."
        Next
        MyConvertIPAddress = Left(MyConvertIPAddress, Len(MyConvertIPAddress) - 1)
End Function


Private Function MyConvertInteger(OctString As String) As String
Dim TmpStr As String
Dim F As Integer

If Asc(OctString) > 127 Then
        For F = 1 To Len(OctString)
            TmpStr = TmpStr & "FF"
            Next
        TmpStr = "&H" & TmpStr
        MyConvertInteger = MyConvertNumber(OctString) - CDbl(TmpStr) - 1
    Else
        MyConvertInteger = MyConvertNumber(OctString)
        End If
End Function

Private Function MyConvertNumber(OctString As String) As String
Dim PStr As String
Dim i As Integer
        'MyConvertNumber = "&H"
        For i = 1 To Len(OctString)
            PStr = Hex(Asc(Mid(OctString, i, 1)))
            If Len(PStr) = 1 Then PStr = "0" & PStr
            MyConvertNumber = MyConvertNumber & PStr
            Next
        'MyConvertNumber = CDbl(MyConvertNumber)
        MyConvertNumber = HexUnsigned2Dbl(MyConvertNumber)
End Function
