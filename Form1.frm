VERSION 5.00
Object = "*\AFreeSnmpOCX.vbp"
Begin VB.Form Form1 
   BorderStyle     =   1  'Pevný okraj
   Caption         =   "SnmpTest"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
   Begin FreeSnmp.FreeSnmpV1 FreeSnmpV11 
      Left            =   9840
      Top             =   120
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   8520
      TabIndex        =   0
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton CmdUptime 
      Caption         =   "Uptime"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton CmdRestart 
      Caption         =   "CM Restart"
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton CmdGet 
      Caption         =   "Test"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   180
      Width           =   1095
   End
   Begin VB.TextBox TxtCommunity 
      Height          =   345
      Left            =   3600
      TabIndex        =   5
      Text            =   "public"
      Top             =   180
      Width           =   1095
   End
   Begin VB.TextBox TxtIP 
      Height          =   345
      Left            =   600
      TabIndex        =   2
      Text            =   "255.255.255.255"
      Top             =   180
      Width           =   1695
   End
   Begin VB.TextBox TxtOut 
      Height          =   6315
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   65535
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Oba
      TabIndex        =   1
      Top             =   720
      Width           =   10215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Prùhledný
      Caption         =   "Community:"
      Height          =   195
      Left            =   2640
      TabIndex        =   4
      Top             =   240
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Prùhledný
      Caption         =   "IP:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   315
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const i = 2 'Integer
Const s = 4 'String
Const n = 5 'Null
Const o = 6 'OID
Const a = 64 'IP address
Const c = 65 'Counter
Const g = 66 'Gauge

Private Sub CmdClear_Click()
TxtOut.Text = ""
End Sub

Private Sub CmdGet_Click()
FreeSnmpV11.RemoteIP = TxtIP.Text
FreeSnmpV11.Community = TxtCommunity.Text
FreeSnmpV11.SnmpGet Int(Rnd * 999) + 1, ".1.3.6.1.2.1.1.1.0", ".1.3.6.1.2.1.1.3.0", _
            ".1.3.6.1.2.1.1.4.0", ".1.3.6.1.2.1.1.5.0", ".1.3.6.1.2.1.1.6.0"
End Sub

Private Sub CmdRestart_Click()
FreeSnmpV11.RemoteIP = "80.188.185.185"
FreeSnmpV11.SnmpSet 1, ".1.3.6.1.2.1.69.1.1.3.0", "private", i, 1
End Sub

Private Sub CmdUptime_Click()
FreeSnmpV11.RemoteIP = TxtIP.Text
FreeSnmpV11.Community = TxtCommunity.Text
FreeSnmpV11.SnmpGet Int(Rnd * 999) + 1001, ".1.3.6.1.2.1.1.3.0"
End Sub

Private Sub FreeSnmpV11_Error(sErrMsg As String)
MsgBox sErrMsg
End Sub

Private Sub FreeSnmpV11_Result(sRemoteIP As String, sReqId As Long, sSnmpErr As Byte, sSnmpErrIndex As Byte, sSnmpOID() As String, sSnmpValueType() As Byte, sSnmpValue() As String)
Dim F As Integer
Dim sTemp As String

For F = LBound(sSnmpValue) To UBound(sSnmpValue)
    sTemp = ConvSnmpType(sSnmpValueType(F)) & ":" & sSnmpValue(F) & " # " & sTemp
    Next
TxtOut.Text = sRemoteIP & " # " & sReqId & " # " & sTemp & vbCrLf & TxtOut.Text
End Sub


Function ConvSnmpType(sType As Byte) As String
Select Case sType
 Case 0
  ConvSnmpType = "TIMEOUT"
 Case 2
  ConvSnmpType = "INTEGER"
 Case 4
  ConvSnmpType = "OCTETSTRING"
 Case 5
  ConvSnmpType = "NULL"
 Case 6
  ConvSnmpType = "OBJECTIDENTIFIER"
 Case 64
  ConvSnmpType = "IPADDRESS"
 Case 65
  ConvSnmpType = "COUNTER"
 Case 66
  ConvSnmpType = "GAUGE"
 Case 67
  ConvSnmpType = "TIMETICKS"
 Case 68
  ConvSnmpType = "OPAQUE"
 Case Else
  ConvSnmpType = "Unknown"
 End Select
End Function



