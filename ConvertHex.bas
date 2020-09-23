Attribute VB_Name = "ConvertHex"
'**************************************
' Name: Convert Unsigned HEX values
' Description:Small function to convert
'     very LONG hex strings containing an unsi
'     gned value into a double. Problem with V
'     Bs native solution ("&H0AF") doing this:
'     It assumes 32-Bit signed! longs ... If y
'     ou are working with values larger than 2
'     .147.483.648 you get negative values, la
'     ter you get errors ... ___ When needed?
'     e.g. when working with SNMP timer ticks:
'     They are counted in seconds/100, so a fe
'     w weeks results in very large values ;)
'     ___ Keywords: HEX, CONVERT, CONVERSION,
'     UNSIGNED, SIGNED.
' By: Light Templer
'
' Inputs:A string with a hexadecimal val
'     ue
'
' Returns:A positive double
'
' Side Effects:Error when input isn't a
'     valid hex value, because I didn't includ
'     e validation for this.
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=58899&lngWId=1'for details.'**************************************



Public Function HexUnsigned2Dbl(sHex As String) As Double
    ' Long unsigned hex values to doubles (u
    '     sed to get with unsigned longs)
    '
    ' 02/14/2005 - LightTempler
    Dim i As Long
    Dim lExpCtr As Long
    sHex = "0" + sHex


    For i = Len(sHex) To 2 Step -2
        lExpCtr = lExpCtr + 1
        HexUnsigned2Dbl = HexUnsigned2Dbl + CDbl("&H" + Right$("00" + Mid$(sHex, i - 1, 2), 2)) * 256# ^ (lExpCtr - 1)
    Next i
End Function
