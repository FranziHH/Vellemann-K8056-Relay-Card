VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsK8056"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event printComment(Message As String)
Public Event printHexData(Message As String)
Public Event commOutput(Message As String)
Public Event Error(Message As String)

Enum vmStatus
  LOW = 0
  HIGH = 1
End Enum

'Standard-Eigenschaftswerte:
Const m_def_Repeat = 1
'Eigenschaftsvariablen:
Dim m_Repeat As Integer

Function CheckSum(Address As Integer, Instruction As Integer, Relay As Integer) As Integer
  ' Calculate the checksum for the RS232 transmission:
  ' 2-complement of the sum of the 4 previous bytes + 1
  CheckSum = 255 - ((Asc(vbCr) + Address + Instruction + (Relay + 48)) Mod 255) + 1
End Function

Sub RelayOn(CardAddress As Integer, Relay As Integer)
  RaiseEvent printComment(Format_String("RelayOn", 14) & ": A" & Format(CardAddress, "000") & " R" & CStr(Relay))
  Call RS232Instruction(CardAddress, "S", Relay)
End Sub

Sub RelayOff(CardAddress As Integer, Relay As Integer)
  RaiseEvent printComment(Format_String("RelayOff", 14) & ": A" & Format(CardAddress, "000") & " R" & CStr(Relay))
  Call RS232Instruction(CardAddress, "C", Relay)
End Sub

Sub EmergencyStop()
  RaiseEvent printComment(Format_String("EmergencyStop", 14))
  Call RS232Instruction(2, "E", 1) ' any cardaddress will do
End Sub

Sub Toggle(CardAddress As Integer, Relay As Integer)
  RaiseEvent printComment(Format_String("Toggle", 14) & ": A" & Format(CardAddress, "000") & " R" & CStr(Relay))
  Call RS232Instruction(CardAddress, "T", Relay)
End Sub

Sub Push(CardAddress As Integer, Relay As Integer)
  RaiseEvent printComment(Format_String("Push", 14) & ": A" & Format(CardAddress, "000") & " R" & CStr(Relay))
  Call RS232Instruction(CardAddress, "P", Relay)
End Sub

Sub SetAddress(CardAddress As Integer, NewCardAddress As Integer)
  RaiseEvent printComment(Format_String("SetAddress", 14) & ": A" & Format(CardAddress, "000") & " N" & Format(NewCardAddress, "000"))
  Call RS232Instruction(CardAddress, "A", NewCardAddress)
End Sub

Sub DisplayAddress()
  RaiseEvent printComment(Format_String("DisplayAddress", 14))
  Call RS232Instruction(1, "D", 1)
End Sub

Sub ResetAddress()
  RaiseEvent printComment(Format_String("ResetAddress", 14))
  Call RS232Instruction(1, "F", 1)
End Sub

Sub digitalWrite(CardAddress As Integer, Relay As Integer, Status As vmStatus)
  RaiseEvent printComment(Format_String("digitalWrite: A", 14) & Format(CardAddress, "000") & " R" & CStr(Relay) & " S=" & IIf(Status = HIGH, "HIGH", "LOW"))
  If Status = HIGH Then
    Call RelayOn(CardAddress, Relay)
  ElseIf Status = LOW Then
    Call RelayOff(CardAddress, Relay)
  End If
End Sub

Sub RS232Instruction(Address As Integer, Instruction As String, Relay As Integer)
  
  Dim intCheckSum As Integer
  Dim NrSends As Integer
  Dim i As Integer
  Dim tmpCommand As String
  Dim SendCommand As String
  
  On Error GoTo ErrHandler
  
  'AddressChange is very special ....
  If Instruction = "A" Then
    intCheckSum = CheckSum(Address, Asc(Instruction), (Relay - 48))
  Else
    intCheckSum = CheckSum(Address, Asc(Instruction), Relay)
  End If
  
  tmpCommand = tmpCommand & vbCr
  tmpCommand = tmpCommand & Chr(Address)
  tmpCommand = tmpCommand & Instruction
  If Instruction = "A" Then
    tmpCommand = tmpCommand & Chr(Relay)
  Else
    tmpCommand = tmpCommand & CStr(Relay)
  End If
  tmpCommand = tmpCommand & Chr(intCheckSum)
  
  For i = 1 To m_Repeat
    SendCommand = SendCommand & tmpCommand
  Next i
  
  RaiseEvent printHexData(tmpCommand) ' prints only ONE Sequence
  RaiseEvent commOutput(SendCommand)
  
  Exit Sub
  
ErrHandler:
  RaiseEvent Error(Err.Number & ", " & Err.Description)
  
End Sub

Private Function Format_String(txt As String, mLen As Long) As String

  Dim mOUT As String
  
  If Len(txt) = mLen Then
    mOUT = txt
  ElseIf Len(txt) > mLen Then
    mOUT = Left(txt, mLen)
  Else
    mOUT = String(mLen - Len(txt), ".") & txt
  End If
  
  Format_String = mOUT
 
End Function

Private Sub Class_Initialize()
  m_Repeat = m_def_Repeat
End Sub

Public Property Get Repeat() As Integer
  Repeat = m_Repeat
End Property

Public Property Let Repeat(ByVal New_Repeat As Integer)
  m_Repeat = New_Repeat
End Property
