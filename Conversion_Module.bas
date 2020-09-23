Attribute VB_Name = "Conversion_Module"
'Author: Anne-Lise Pasch
'E-mail: annelise@slayers.co.uk
'String Conversion Module
'http://www.freevbcode.com/ShowCode.Asp?ID=2921

Option Explicit

Public Function BinaryToDouble(ByVal strData As String) As Double
    Dim dblOutput As Double
    Dim lngIterator As Long
    Do Until Len(strData) = 0
        dblOutput = dblOutput + IIf(Right$(strData, 1) = "1", (2 ^ lngIterator), 0)
        strData = Left$(strData, Len(strData) - 1)
        lngIterator = lngIterator + 1
    Loop
    BinaryToDouble = dblOutput
End Function
Public Function ByteToString(ByRef bytData() As Byte, ByVal lngDataLength As Long) As String
    Dim lngIterator As Long
    For lngIterator = LBound(bytData) To (LBound(bytData) + lngDataLength)
        ByteToString = ByteToString & Chr$(bytData(lngIterator))
    Next lngIterator
End Function
Public Function Capitalise(ByVal strData As String) As String
    Capitalise = UCase(Left$(strData, 1)) + LCase(Right$(strData, Len(strData) - 1))
End Function
Public Function DoubleToBinary(ByVal dblData As Double) As String
    Dim strOutput As String
    Dim lngIterator As Long
    Do Until (2 ^ lngIterator) > dblData
        strOutput = IIf(((2 ^ lngIterator) And dblData) > 0, "1", "0") + strOutput
        lngIterator = lngIterator + 1
    Loop
    DoubleToBinary = strOutput
End Function
Public Function HexToString(ByVal strData As String) As String
    Dim strOutput As String
    Do Until Len(strData) < 2
        strOutput = strOutput + Chr$(CLng("&H" + Left$(strData, 2)))
        strData = Right$(strData, Len(strData) - 2)
    Loop
    HexToString = strOutput
End Function
Public Function NoNulls(ByVal varData As Variant, Optional ByVal varDefault As Variant = "") As Variant
    NoNulls = IIf(TypeName(varData) <> "Null", varData, varDefault)
End Function
Public Function StringToHex(ByVal strData As String) As String
    Dim strOutput As String
    Do Until Len(strData) = 0
        strOutput = strOutput + Right$(String$(2, "0") + Hex$(Asc(Left$(strData, 1))), 2)
        strData = Right$(strData, Len(strData) - 1)
    Loop
    StringToHex = strOutput
End Function

