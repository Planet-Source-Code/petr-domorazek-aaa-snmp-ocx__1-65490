VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl FreeSnmpV1 
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   525
   InvisibleAtRuntime=   -1  'True
   Picture         =   "FreeSnmpV1.ctx":0000
   ScaleHeight     =   525
   ScaleWidth      =   525
   ToolboxBitmap   =   "FreeSnmpV1.ctx":0139
   Begin MSWinsockLib.Winsock WS1 
      Left            =   0
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "FreeSnmpV1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private OcxRemoteIP As String
Private OcxPort As String
Private OcxCommunity As String


Public Event Error(sErrMsg As String)
Public Event Result(sRemoteIP As String, sReqId As Long, sSnmpErr As Byte, sSnmpErrIndex As Byte, _
                    sSnmpOID() As String, sSnmpValueType() As Byte, sSnmpValue() As String)
                    

Public Property Get RemoteIP() As String
    RemoteIP = OcxRemoteIP
End Property

Public Property Let RemoteIP(ByVal vNewValue As String)
    OcxRemoteIP = vNewValue
    PropertyChanged "RemoteIP"
End Property

Public Property Get Port() As String
    Port = OcxPort
End Property

Public Property Let Port(ByVal vNewValue As String)
    OcxPort = vNewValue
    PropertyChanged "Port"
End Property

Public Property Get Community() As String
    Community = OcxCommunity
End Property

Public Property Let Community(ByVal vNewValue As String)
    OcxCommunity = vNewValue
    PropertyChanged "Community"
End Property

Public Sub SnmpGet(RequestID As Long, ParamArray sOID_arr() As Variant)

Dim OutPacket As String
Dim F As Integer
Dim sTempOID() As String

On Local Error GoTo error_handler

If UBound(sOID_arr) = -1 Then
     RaiseEvent Error("Error, OID is empty.")
     Exit Sub
     End If
If OcxRemoteIP = "" Then
     RaiseEvent Error("RemoteIP error")
     Exit Sub
     End If
If OcxPort < 1 Or OcxPort > 65535 Then
     RaiseEvent Error("Port error")
     Exit Sub
     End If
If OcxCommunity = "" Then
     RaiseEvent Error("Community error")
     Exit Sub
     End If
If RequestID < 1 Then
     RaiseEvent Error("ReqID error")
     Exit Sub
     End If
     
ReDim sTempOID(LBound(sOID_arr) To UBound(sOID_arr))
For F = LBound(sOID_arr) To UBound(sOID_arr)
    sTempOID(F) = CStr(sOID_arr(F))
    Next
   
OutPacket = GenMultiPacket(OcxCommunity, RequestID, 160, sTempOID())

With WS1
 .RemotePort = OcxPort
 .RemoteHost = OcxRemoteIP
 .SendData OutPacket
End With
        
        
Exit Sub
error_handler:
    
    RaiseEvent Error("[" & Err.Description & "] in 'SnmpGetNext()'")
        
    
End Sub



Public Sub SnmpGetNext(RequestID As Long, ParamArray sOID_arr() As Variant)

Dim OutPacket As String
Dim F As Integer
Dim sTempOID() As String

On Local Error GoTo error_handler

If UBound(sOID_arr) = -1 Then
     RaiseEvent Error("Error, OID is empty.")
     Exit Sub
     End If
If OcxRemoteIP = "" Then
     RaiseEvent Error("RemoteIP error")
     Exit Sub
     End If
If OcxPort < 1 Or OcxPort > 65535 Then
     RaiseEvent Error("Port error")
     Exit Sub
     End If
If OcxCommunity = "" Then
     RaiseEvent Error("Community error")
     Exit Sub
     End If
If RequestID < 1 Then
     RaiseEvent Error("ReqID error")
     Exit Sub
     End If
     
ReDim sTempOID(LBound(sOID_arr) To UBound(sOID_arr))
For F = LBound(sOID_arr) To UBound(sOID_arr)
    sTempOID(F) = CStr(sOID_arr(F))
    Next

OutPacket = GenMultiPacket(OcxCommunity, RequestID, 161, sTempOID())

With WS1
 .RemotePort = OcxPort
 .RemoteHost = OcxRemoteIP
 .SendData OutPacket
End With
        
        
Exit Sub
error_handler:
    
    RaiseEvent Error("[" & Err.Description & "] in 'SnmpGetNext()'")
        
    
End Sub

Public Sub SnmpSet(RequestID As Long, sOID_set As String, sCommunity_set As String, _
                    ValueType As Byte, Value As Variant)
Dim OutPacket As String
Dim F As Integer
    
On Local Error GoTo error_handler
        
If OcxRemoteIP = "" Then
     RaiseEvent Error("RemoteIP error")
     Exit Sub
     End If
If OcxPort = 0 Or OcxPort > 65535 Then
     RaiseEvent Error("Port error")
     Exit Sub
     End If
If sCommunity_set = "" Then
     RaiseEvent Error("Community error")
     Exit Sub
     End If
If sOID_set = "" Then
     RaiseEvent Error("OID error")
     Exit Sub
     End If
If RequestID > 1 Then
     RaiseEvent Error("ReqID error")
     Exit Sub
     End If
        
OutPacket = GenSimplePacket(Trim(sCommunity_set), Trim(sOID_set), RequestID, 163, ValueType, Value)


With WS1
 .RemotePort = OcxPort
 .RemoteHost = OcxRemoteIP
 .SendData OutPacket
End With
        
Exit Sub
error_handler:
    
    RaiseEvent Error("[" & Err.Description & "] in 'SnmpSet()'")
    
End Sub




'##### UserControl ############################################################

Private Sub UserControl_Initialize()
On Local Error GoTo error_handler

    With WS1
        .Protocol = sckUDPProtocol
        .LocalPort = 0
        .RemotePort = 161
    End With

    If OcxPort = "" Then OcxPort = "161"
    If OcxCommunity = "" Then OcxCommunity = "public"
    Exit Sub

error_handler:
    
    RaiseEvent Error("[" & Err.Description & "] in 'Control_Initialize()")
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    OcxRemoteIP = PropBag.ReadProperty("RemoteIP", "")
    OcxPort = PropBag.ReadProperty("Port", "161")
    OcxCommunity = PropBag.ReadProperty("Community", "public")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Community", OcxCommunity, "public")
    Call PropBag.WriteProperty("RemoteIP", OcxRemoteIP, "")
    Call PropBag.WriteProperty("Port", OcxPort, "161")
End Sub

Private Sub UserControl_Resize()
   UserControl.Height = 500
   UserControl.Width = 500
End Sub

'##### WS1 ############################################################

Private Sub WS1_DataArrival(ByVal bytesTotal As Long)
'DecodePacket
'************
'(0) = "255" 'No error
'(1) = SnmpPacketID
'(2) = SnmpOID
'(3) = SnmpVType
'(4) = SnmpVal
'(5) = ErrorByte
'(6) = OrgPacket
'(7) = SnmpStatusError
'(8) = SnmpErrorIndex
        
If bytesTotal > 1 Then
        
Dim AData As String
Dim DeAAata() As String
Dim RemIP As String
Dim sId As Long
Dim sError As Byte
Dim sErrorIndex As Byte
Dim sOID() As String
Dim sType() As Byte
Dim sValue() As String
Dim N, mTemp As Integer


    WS1.GetData AData
    RemIP = WS1.RemoteHostIP

    DeAAata = DecodePacket(AData)

    If DeAAata(0, 0) = 255 Then

        sId = DeAAata(0, 1)
        sError = DeAAata(0, 7)
        sErrorIndex = DeAAata(0, 8)
   
            mTemp = UBound(DeAAata)
            ReDim sOID(0 To mTemp)
            ReDim sType(0 To mTemp)
            ReDim sValue(0 To mTemp)
       
            For N = 0 To mTemp
                sOID(N) = DeAAata(N, 2)
                If sError = 0 Then
                   sType(N) = CByte(DeAAata(N, 3))
                   sValue(N) = ConvertSnmpValue(CByte(DeAAata(N, 3)), DeAAata(N, 4))
                Else
                   sType(N) = 255
                   sValue(N) = SnmpErrStatusDesc(sError) & ", OID Index: " & sErrorIndex
                   End If
                Next


        RaiseEvent Result(RemIP, sId, sError, sErrorIndex, sOID(), sType(), sValue())
    Else
        'Error in decode packet
        RaiseEvent Error("Error in packet: " & DeAAata(0, 4) & " Byte no. " & DeAAata(0, 5))
        End If

Else
        RaiseEvent Error("Reset Connection")
    End If

  
End Sub

Private Sub WS1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent Error("WinSock Error: " & Number & " " & Description)
    CancelDisplay = True
End Sub
