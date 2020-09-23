Attribute VB_Name = "EncodeSnmpVal"
'Author: Petr Domorazek
'E-mail: dsoft@fcanet.cz
'http://dsoft.php5.cz

Option Explicit

Function EncodeSnmpValue(SnmpValueType As Byte, SnmpValueEn As Variant) As String
Dim i As Integer
Select Case SnmpValueType
    Case 2 'Integer
        EncodeSnmpValue = EnInteger(SnmpValueEn)
    Case 4 'Octet String
        EncodeSnmpValue = CStr(SnmpValueEn)
    Case 5 'Null
        EncodeSnmpValue = ""
    Case 6 'OID
        EncodeSnmpValue = EnOID(CStr(SnmpValueEn))
    Case &H40 'IP Address
        EncodeSnmpValue = EnIPAddress(SnmpValueEn)
    Case &H41 'Counter
        EncodeSnmpValue = EnNumber(SnmpValueEn)
    Case &H42 'Gauge
        EncodeSnmpValue = EnNumber(SnmpValueEn)
    'Case &H43 'TimeTicks
        'EncodingSnmpValue = EnTimeTicks(SnmpValueEn)
    Case Else 'Other values (Null)
        EncodeSnmpValue = ""
End Select
End Function

Private Function EnInteger(ValEn As Variant) As String
Dim TmpStr, TmpStr2  As String
Dim TmpDbl, TmpDbl2, F As Long

If ValEn < 0 Then
    TmpStr = EnNumber(Abs(ValEn))
    TmpStr2 = "&H"
    For F = 1 To Len(TmpStr)
        TmpStr2 = TmpStr2 & "FF"
        Next
    TmpDbl = (CLng(TmpStr2) + 1) \ 2
    If Abs(ValEn) > TmpDbl Then
        TmpStr2 = TmpStr2 & "FF"
        TmpDbl = (CLng(TmpStr2) + 1)
        TmpDbl2 = TmpDbl + ValEn
    Else
        TmpDbl = (CLng(TmpStr2) + 1)
        TmpDbl2 = TmpDbl + ValEn
        End If
    EnInteger = EnNumber(TmpDbl2)
Else
    TmpStr = EnNumber(ValEn)
    If Asc(TmpStr) > 127 Then TmpStr = Chr(0) & TmpStr
    EnInteger = TmpStr
    End If
End Function

Private Function EnNumber(ValEn As Variant) As String
Dim TmpStr As String
Dim TmpNum, F As Long

TmpNum = CLng(ValEn)
TmpStr = Hex(TmpNum)
If Len(TmpStr) Mod 2 = 1 Then TmpStr = "0" & TmpStr
For F = 1 To Len(TmpStr) Step 2
    EnNumber = EnNumber & Chr(Val("&H" & Mid(TmpStr, F, 2)))
    Next
End Function

Private Function EnIPAddress(ValEn As Variant) As String
Dim TmpStr As String
Dim tmpArr() As String
Dim F As Integer

TmpStr = CStr(ValEn)
tmpArr = Split(TmpStr, ".")
TmpStr = ""
If UBound(tmpArr) <> 3 Then
    EnIPAddress = Chr(0) & Chr(0) & Chr(0) & Chr(0)
    Exit Function
    End If
For F = LBound(tmpArr) To UBound(tmpArr)
    If Val(tmpArr(F)) > 255 Then
        TmpStr = TmpStr & Chr(255)
    Else
        TmpStr = TmpStr & Chr(Val(tmpArr(F)))
        End If
    Next
EnIPAddress = TmpStr
End Function


