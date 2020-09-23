Attribute VB_Name = "Snmp_Function"
'Author: Petr Domorazek
'E-mail: dsoft@fcanet.cz
'http://dsoft.php5.cz

Option Explicit

Function DecodePacket(AData As String) As String()
Dim TempNum As Long
Dim TempNum2 As Long
Dim temp As String
Dim SnmpVType() As Byte
Dim SnmpVal() As String
Dim SnmpPacketID As Long
Dim SnmpOID() As String
Dim SnmpArr() As String
Dim ErrSnmpByte As Byte
Dim ErrSnmpDesc As String
Dim PacketLen As Long
Dim OrgPacket As String
Dim SnmpStatusError As Byte
Dim SnmpErrorIndex As Byte
Dim N As Integer
Dim M As Integer

OrgPacket = AData
PacketLen = Len(AData)
If Asc(AData) <> 48 Then
                        ErrSnmpByte = 1
                        ErrSnmpDesc = "Start byte must be H30!"
                        Debug.Print ErrSnmpDesc
                        GoTo SnmpError
                        End If
AData = Mid(AData, 2)
TempNum = DeCRCNum(AData)
TempNum2 = DeCRCNumLEN(AData)
AData = Mid(AData, 1 + TempNum2)
If Len(AData) <> TempNum Then
                        ErrSnmpByte = PacketLen - Len(AData) + 1
                        ErrSnmpDesc = "Packet length error! Code: L1"
                        Debug.Print ErrSnmpDesc
                        GoTo SnmpError
                        End If
temp = Chr(2) & Chr(1) & Chr(0)   ' Snmp V1
If Mid(AData, 1, Len(temp)) <> temp Then
                        ErrSnmpByte = PacketLen - Len(AData)
                        ErrSnmpDesc = "Error in packet version!"
                        Debug.Print ErrSnmpDesc
                        GoTo SnmpError
                        End If
AData = Mid(AData, 1 + Len(temp))
If Asc(AData) <> 4 Then
                        ErrSnmpByte = PacketLen - Len(AData)
                        ErrSnmpDesc = "Error in packet community!"
                        Debug.Print ErrSnmpDesc
                        GoTo SnmpError
                        End If
AData = Mid(AData, 2)
TempNum = DeCRCNum(AData)
TempNum2 = DeCRCNumLEN(AData)
AData = Mid(AData, 1 + TempNum2)
AData = Mid(AData, 1 + TempNum)
If Asc(AData) <> 162 Then
                        Debug.Print "ConvertAPacket: Error packet A2 - Response"
                        ErrSnmpByte = PacketLen - Len(AData)
                        ErrSnmpDesc = "Error in packet: HA2 - Response not found!"
                        Debug.Print ErrSnmpDesc
                        GoTo SnmpError
                        End If
AData = Mid(AData, 2)
TempNum = DeCRCNum(AData)
TempNum2 = DeCRCNumLEN(AData)
AData = Mid(AData, 1 + TempNum2)
If Len(AData) <> TempNum Then
                        ErrSnmpByte = PacketLen - Len(AData)
                        ErrSnmpDesc = "Packet length error! PDU Code: L2"
                        Debug.Print ErrSnmpDesc
                        GoTo SnmpError
                        End If
temp = Chr(2) & Chr(1) & Chr(1)  ' Packet ID
If Asc(AData) <> 2 Then
                        ErrSnmpByte = PacketLen - Len(AData)
                        ErrSnmpDesc = "Packet ID Error!"
                        Debug.Print ErrSnmpDesc
                        GoTo SnmpError
                        End If
AData = Mid(AData, 2)
TempNum = DeCRCNum(AData)
TempNum2 = DeCRCNumLEN(AData)
AData = Mid(AData, 1 + TempNum2)
SnmpPacketID = DeBNum(Left(AData, TempNum))
AData = Mid(AData, 1 + TempNum)

temp = Chr(2) & Chr(1)    ' Error
If Mid(AData, 1, Len(temp)) <> temp Then
                        ErrSnmpByte = PacketLen - Len(AData)
                        ErrSnmpDesc = "SNMP Error Code."
                        Debug.Print ErrSnmpDesc
                        GoTo SnmpError
                        End If
AData = Mid(AData, 1 + Len(temp))
SnmpStatusError = Asc(AData)
AData = Mid(AData, 2)
temp = Chr(2) & Chr(1)     ' Error Index
If Mid(AData, 1, Len(temp)) <> temp Then
                        ErrSnmpByte = PacketLen - Len(AData)
                        ErrSnmpDesc = "SNMP Error Index Code."
                        Debug.Print ErrSnmpDesc
                        GoTo SnmpError
                        End If
AData = Mid(AData, 1 + Len(temp))
SnmpErrorIndex = Asc(AData)
AData = Mid(AData, 2)

If Asc(AData) <> 48 Then
                        ErrSnmpByte = PacketLen - Len(AData)
                        ErrSnmpDesc = "Start byte VarBindList must be H30! Code: V1"
                        Debug.Print ErrSnmpDesc
                        GoTo SnmpError
                        End If
AData = Mid(AData, 2)
TempNum = DeCRCNum(AData)
TempNum2 = DeCRCNumLEN(AData)
AData = Mid(AData, 1 + TempNum2)
If Len(AData) <> TempNum Then
                        ErrSnmpByte = PacketLen - Len(AData)
                        ErrSnmpDesc = "VarBindList length error! Code: L3"
                        Debug.Print ErrSnmpDesc
                        GoTo SnmpError
                        End If

ReDim SnmpOID(0 To 25)
ReDim SnmpVType(0 To 25)
ReDim SnmpVal(0 To 25)
For N = 0 To 25
    If Asc(AData) <> 48 Then
                            ErrSnmpByte = PacketLen - Len(AData)
                            ErrSnmpDesc = "Start byte iVarBindList be H30! Code: V2"
                            Debug.Print ErrSnmpDesc
                            GoTo SnmpError
                            End If
    AData = Mid(AData, 2)
    TempNum = DeCRCNum(AData)
    TempNum2 = DeCRCNumLEN(AData)
    AData = Mid(AData, 1 + TempNum2)
    If Len(AData) <> TempNum Then
            If Asc(Mid(AData, TempNum + 1)) <> 48 Then
                            ErrSnmpByte = PacketLen - Len(AData)
                            ErrSnmpDesc = "VarBindList length error! Code: L4" & N
                            Debug.Print ErrSnmpDesc
                            GoTo SnmpError
                            End If
                            End If
    If Asc(AData) <> 6 Then
                            ErrSnmpByte = PacketLen - Len(AData)
                            ErrSnmpDesc = "OID error! CODE OID" & N
                            Debug.Print ErrSnmpDesc
                            GoTo SnmpError
                            End If

    AData = Mid(AData, 2)
    TempNum = DeCRCNum(AData)
    TempNum2 = DeCRCNumLEN(AData)
    AData = Mid(AData, 1 + TempNum2)
    SnmpOID(N) = DeOID(Left(AData, TempNum)) 'OID
    AData = Mid(AData, 1 + TempNum)
    SnmpVType(N) = Asc(AData) ' Data type
    AData = Mid(AData, 2)
    TempNum = DeCRCNum(AData)
    TempNum2 = DeCRCNumLEN(AData)
    AData = Mid(AData, 1 + TempNum2)
    If Len(AData) <> TempNum Then
            If Asc(Mid(AData, TempNum + 1)) <> 48 Then
                            ErrSnmpByte = PacketLen - Len(AData)
                            ErrSnmpDesc = "Value length error! CODE: VAL" & N
                            Debug.Print ErrSnmpDesc
                            GoTo SnmpError
                            End If
                            End If
    SnmpVal(N) = Mid(AData, 1, TempNum)
    AData = Mid(AData, TempNum + 1)
    If AData = "" Then Exit For
Next



ReDim SnmpArr(0 To N, 0 To 8)
SnmpArr(0, 6) = OrgPacket
For M = 0 To N
    SnmpArr(M, 0) = "255" 'No error
    SnmpArr(M, 1) = SnmpPacketID
    SnmpArr(M, 2) = SnmpOID(M)
    SnmpArr(M, 3) = SnmpVType(M)
    SnmpArr(M, 4) = SnmpVal(M)
    SnmpArr(M, 5) = 0
    SnmpArr(M, 6) = ""
    SnmpArr(M, 7) = SnmpStatusError
    SnmpArr(M, 8) = SnmpErrorIndex
    Next
    
DecodePacket = SnmpArr


Exit Function
SnmpError:
ReDim SnmpArr(0 To 0, 0 To 8)

SnmpArr(0, 0) = "0" 'Error
SnmpArr(0, 1) = 0
'SnmpArr(0, 2) = SnmpOID(0)
'SnmpArr(0, 2) = ""
If SnmpArr(0, 2) = "" Then SnmpArr(0, 2) = "1.3"
SnmpArr(0, 3) = 0
SnmpArr(0, 4) = ErrSnmpDesc
SnmpArr(0, 5) = ErrSnmpByte
SnmpArr(0, 6) = OrgPacket
SnmpArr(0, 7) = SnmpStatusError
SnmpArr(0, 8) = SnmpErrorIndex

DecodePacket = SnmpArr

Exit Function
FunctionError:

End Function

Function GenMultiPacket(SnmpCommunity As String, SnmpPacketID As Long, SnmpReqType As Byte, SnmpOID() As String) As String
'SnmpOID As String,
Dim temp As String
Dim TempNum As Long
Dim SnmpO As String
Dim N As Long

Rem Gen
SnmpO = ""

For N = UBound(SnmpOID) To LBound(SnmpOID) Step -1
    SnmpO = Chr(5) & Chr(0) & SnmpO ' Value Null
    temp = EnOID(CStr(SnmpOID(N))) ' OID
    SnmpO = Chr(6) & Chr(Len(temp)) & temp & SnmpO  ' OID
    TempNum = Len(Chr(6) & Chr(Len(temp)) & temp & Chr(5) & Chr(0))
    SnmpO = Chr(48) & EnCRCNum(TempNum) & SnmpO   ' PDU1
    Next
    
TempNum = Len(SnmpO)
SnmpO = Chr(48) & EnCRCNum(TempNum) & SnmpO   ' PDU
SnmpO = Chr(2) & Chr(1) & Chr(0) & SnmpO    ' Error Index
SnmpO = Chr(2) & Chr(1) & Chr(0) & SnmpO    ' Error
TempNum = SnmpPacketID 'Packet ID
temp = EnBNum(TempNum)
TempNum = Len(temp)
SnmpO = Chr(2) & EnCRCNum(TempNum) & temp & SnmpO    ' Packet ID
TempNum = Len(SnmpO)
SnmpO = Chr(SnmpReqType) & EnCRCNum(TempNum) & SnmpO   ' GET=160 GETNEXT=161
temp = SnmpCommunity 'Community
SnmpO = Chr(4) & Chr(Len(temp)) & temp & SnmpO  'Community
SnmpO = Chr(2) & Chr(1) & Chr(0) & SnmpO    ' Snmp V1
TempNum = Len(SnmpO)
SnmpO = Chr(48) & EnCRCNum(TempNum) & SnmpO   ' PDU
GenMultiPacket = SnmpO
End Function

Function GenSimplePacket(SnmpCommunity As String, SnmpOID As String, SnmpPacketID As Long, SnmpReqType As Byte, SnmpVType As Byte, SnmpValS As Variant) As String
Dim temp As String
Dim TempNum As Long
Dim SnmpO As String

Rem Gen
SnmpO = ""

temp = EncodeSnmpValue(SnmpVType, SnmpValS) 'Value
SnmpO = Chr(Len(temp)) & temp 'Value Len
SnmpO = Chr(SnmpVType) & SnmpO 'Value Type
temp = EnOID(SnmpOID) ' OID
SnmpO = Chr(6) & Chr(Len(temp)) & temp & SnmpO  ' OID
TempNum = Len(SnmpO)
SnmpO = Chr(48) & EnCRCNum(TempNum) & SnmpO   ' PDU1
TempNum = Len(SnmpO)
SnmpO = Chr(48) & EnCRCNum(TempNum) & SnmpO   ' PDU
SnmpO = Chr(2) & Chr(1) & Chr(0) & SnmpO    ' Error Index
SnmpO = Chr(2) & Chr(1) & Chr(0) & SnmpO    ' Error
TempNum = SnmpPacketID 'Packet ID
temp = EnBNum(TempNum)
TempNum = Len(temp)
SnmpO = Chr(2) & EnCRCNum(TempNum) & temp & SnmpO    ' Packet ID
TempNum = Len(SnmpO)
SnmpO = Chr(SnmpReqType) & EnCRCNum(TempNum) & SnmpO   ' GET=160 GETNEXT=161
temp = SnmpCommunity 'Community
SnmpO = Chr(4) & Chr(Len(temp)) & temp & SnmpO  'Community
SnmpO = Chr(2) & Chr(1) & Chr(0) & SnmpO    ' Snmp V1
TempNum = Len(SnmpO)
SnmpO = Chr(48) & EnCRCNum(TempNum) & SnmpO   ' PDU
GenSimplePacket = SnmpO
End Function


Function EnOID(OIDIn As String) As String
Dim OIDArray As Variant
Dim OIDSize As Integer
Dim F As Integer

OIDIn = Trim(OIDIn)

If Left(OIDIn, 1) = "." Then OIDIn = Mid(OIDIn, 2)
If Right(OIDIn, 1) = "." Then OIDIn = Left(OIDIn, Len(OIDIn) - 1)
If OIDIn = "1.3" Then
            EnOID = Chr(&H2B) '2B = 1.3
            Exit Function
            End If
                        
If Left(OIDIn, 4) <> "1.3." Then
            EnOID = ""
            Debug.Print "EnOID: OID do must start 1.3 or .1.3" & vbCrLf
            Exit Function
            End If

OIDIn = Mid(OIDIn, 5)
EnOID = Chr(&H2B) '2B = 1.3

OIDArray = Split(OIDIn, ".", , vbBinaryCompare)
OIDSize = UBound(OIDArray)
If OIDArray(UBound(OIDArray)) = "" Then OIDSize = OIDSize - 1

For F = 0 To OIDSize
    EnOID = EnOID & EnOIDNum(Val(OIDArray(F)))
    Next

End Function


Function DeOID(OIDIn As String) As String
Dim NumTemp As Long
Dim NumTemp2 As Long
Dim F As Integer

                 
If Left(OIDIn, 1) <> Chr(&H2B) Then
            DeOID = ""
            Debug.Print "DeOID: OID do must start 2B" & vbCrLf
            Exit Function
            End If

OIDIn = Mid(OIDIn, 2)
DeOID = ".1.3" '2B = 1.3

For F = 1 To Len(OIDIn)
    NumTemp = DeOIDNumLEN(Mid(OIDIn, F))
    NumTemp2 = DeOIDNum(Mid(OIDIn, F, NumTemp))
    DeOID = DeOID & "." & NumTemp2
    If NumTemp > 1 Then F = F + NumTemp - 1
    Next

End Function


Function EnOIDNum(NumTemp As Long) As String
Dim N As Integer
Dim M As Integer
Dim TrVal As Long
Dim TrBin As String
Dim TrNBin As String
Dim TrCurBin As String
Dim temp As Double


TrVal = Round(NumTemp)
If TrVal < 128 Then
        EnOIDNum = Chr(TrVal)
Else
        TrBin = StrReverse(BIN(TrVal))
        For N = 1 To Len(TrBin) Step 7
            TrCurBin = Mid(TrBin, N, 7)
            If N = 1 Then
                TrNBin = TrCurBin & "0"
            Else
                TrNBin = TrNBin & TrCurBin
                If Len(TrCurBin) < 7 Then TrNBin = TrNBin & Space(7 - Len(TrCurBin))
                TrNBin = Replace(TrNBin, " ", "0")
                TrNBin = TrNBin & "1"
                End If
            Next
        TrNBin = StrReverse(TrNBin)
        N = 1
        For N = 1 To Len(TrNBin) Step 8
                temp = BinaryToDouble(Mid(TrNBin, N, 8))
                EnOIDNum = EnOIDNum & Chr(temp)
                Next
        End If

End Function

Function EnBNum(NumTemp As Long) As String
Dim N As Integer
Dim temp As String

temp = (Hex(NumTemp))
If (Len(temp) Mod 2) = 1 Then temp = "0" & temp
For N = 1 To Len(temp) Step 2
    EnBNum = EnBNum & Chr(Val("&H" & Mid(temp, N, 2)))
    Next
End Function

Function DeBNum(temp As String) As Long
temp = StringToHex(temp)
DeBNum = Val("&H" & temp)
End Function


Function DeOIDNum(temp As String) As Long
Dim N As Long
Dim M As Integer
Dim TempCplBin As String
Dim TempBin As String

For N = 1 To Len(temp)
    TempBin = Right(BIN(Asc(Mid(temp, N, 1))), 7)
    If Len(TempBin) < 7 Then
        For M = 1 To 7 - Len(TempBin)
            TempBin = "0" & TempBin
            Next
        End If
    TempCplBin = TempCplBin & TempBin
    If Asc(Mid(temp, N, 1)) < 128 Then Exit For
    Next

DeOIDNum = CLng(BinaryToDouble(TempCplBin))
End Function

Function DeOIDNumLEN(temp As String) As Long
Dim N As Long
Dim M As Integer

For N = 1 To Len(temp)
    If Asc(Mid(temp, N, 1)) < 128 Then Exit For
    Next
DeOIDNumLEN = N
End Function


Function DeCRCNum(temp As String) As Long
Dim N As Integer
Dim LenTemp As Integer
Dim Temp0 As String
Dim Temp2 As String
Dim Temp3 As String

If Asc(temp) < 128 Then
    DeCRCNum = Asc(temp)
    Exit Function
    End If

LenTemp = Asc(temp) And 127
Temp0 = Mid(temp, 2, LenTemp)
For N = 1 To Len(Temp0)
    Temp3 = Hex(Asc(Mid(Temp0, N, 1)))
    If Len(Temp3) = 1 Then Temp3 = "0" & Temp3
    Temp2 = Temp2 & Temp3
    Next
DeCRCNum = Val("&H" & Temp2)
End Function


Function EnCRCNum(NumTemp As Long) As String
Dim N As Integer
Dim temp As String
Dim Temp2 As String

If NumTemp < 128 Then
    EnCRCNum = Chr(NumTemp)
    Exit Function
    End If

temp = Hex(NumTemp)
If (Len(temp) Mod 2) = 1 Then temp = "0" & temp

Temp2 = Chr(128 + (Len(temp) / 2))
For N = 1 To Len(temp) Step 2
    Temp2 = Temp2 & Chr(Val("&H" & Mid(temp, N, 2)))
    Next
EnCRCNum = Temp2
End Function


Function DeCRCNumLEN(sTemp As String) As Long
If Asc(sTemp) < 128 Then
    DeCRCNumLEN = 1
    Exit Function
    End If

DeCRCNumLEN = (Asc(sTemp) And 127) + 1
End Function


Function SnmpErrStatusDesc(TempNum As Byte) As String
Dim TempErr As String
Select Case TempNum
    Case 0
     TempErr = "noError (0)"
    Case 1
     TempErr = "tooBig (1)"
    Case 2
     TempErr = "noSuchName (2)"
    Case 3
     TempErr = "badValue (3)"
    Case 4
     TempErr = "ReadOnly (4)"
    Case 5
     TempErr = "genErr (5)"
    Case Else
     TempErr = "Other (" & TempNum & ")"
    End Select
SnmpErrStatusDesc = TempErr
End Function
