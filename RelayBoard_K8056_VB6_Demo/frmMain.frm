VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "8 Channel Relay Board"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame frameGetAddress 
      Caption         =   "Calc Display Address"
      Height          =   735
      Left            =   120
      TabIndex        =   59
      Top             =   4800
      Width           =   5475
      Begin VB.TextBox txtCalcAddress 
         Appearance      =   0  '2D
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "0"
         ToolTipText     =   "Double Click for RESET"
         Top             =   240
         Width           =   915
      End
      Begin VB.CheckBox chkRelay 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   60
         Top             =   300
         Width           =   195
      End
      Begin VB.CheckBox chkRelay 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   780
         TabIndex        =   61
         Top             =   300
         Width           =   195
      End
      Begin VB.CheckBox chkRelay 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1260
         TabIndex        =   62
         Top             =   300
         Width           =   195
      End
      Begin VB.CheckBox chkRelay 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1740
         TabIndex        =   63
         Top             =   300
         Width           =   195
      End
      Begin VB.CheckBox chkRelay 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   2220
         TabIndex        =   64
         Top             =   300
         Width           =   195
      End
      Begin VB.CheckBox chkRelay 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   2700
         TabIndex        =   65
         Top             =   300
         Width           =   195
      End
      Begin VB.CheckBox chkRelay 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   3180
         TabIndex        =   66
         Top             =   300
         Width           =   195
      End
      Begin VB.CheckBox chkRelay 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   3660
         TabIndex        =   67
         Top             =   300
         Width           =   195
      End
   End
   Begin VB.Frame frameDebug 
      Caption         =   "Debug"
      Height          =   5415
      Left            =   5760
      TabIndex        =   68
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Debug Window"
         Height          =   315
         Left            =   1740
         TabIndex        =   70
         Top             =   4980
         Width           =   2235
      End
      Begin VB.TextBox txtDataHex 
         Appearance      =   0  '2D
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4515
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   69
         Top             =   300
         Width           =   4155
      End
   End
   Begin VB.CommandButton cmdEmergency 
      Caption         =   "Emergency   S  T  O  P"
      Height          =   1935
      Left            =   4440
      TabIndex        =   58
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame frameAddress 
      Caption         =   "Address"
      Height          =   1215
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton cmdChangeAddress 
         Caption         =   "Change"
         Height          =   315
         Left            =   1380
         TabIndex        =   10
         Top             =   480
         Width           =   1275
      End
      Begin VB.CommandButton cmdDisplayAddress 
         Caption         =   "Display"
         Height          =   315
         Left            =   1380
         TabIndex        =   12
         Top             =   840
         Width           =   1275
      End
      Begin VB.CommandButton cmdResetAddress 
         Caption         =   "Reset All"
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   840
         Width           =   1275
      End
      Begin VB.CommandButton cmdSetAddress 
         Caption         =   "Set Address"
         Height          =   315
         Left            =   1380
         TabIndex        =   9
         Top             =   120
         Width           =   1275
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  '2D
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Text            =   "1"
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame frameToggle 
      Caption         =   "Toggle buttons"
      Height          =   735
      Left            =   120
      TabIndex        =   40
      Top             =   3120
      Width           =   4095
      Begin VB.CommandButton cmdToggle 
         Caption         =   "1"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   41
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdToggle 
         Caption         =   "2"
         Height          =   375
         Index           =   2
         Left            =   660
         TabIndex        =   42
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdToggle 
         Caption         =   "3"
         Height          =   375
         Index           =   3
         Left            =   1140
         TabIndex        =   43
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdToggle 
         Caption         =   "4"
         Height          =   375
         Index           =   4
         Left            =   1620
         TabIndex        =   44
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdToggle 
         Caption         =   "5"
         Height          =   375
         Index           =   5
         Left            =   2100
         TabIndex        =   45
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdToggle 
         Caption         =   "6"
         Height          =   375
         Index           =   6
         Left            =   2580
         TabIndex        =   46
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdToggle 
         Caption         =   "7"
         Height          =   375
         Index           =   7
         Left            =   3060
         TabIndex        =   47
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdToggle 
         Caption         =   "8"
         Height          =   375
         Index           =   8
         Left            =   3540
         TabIndex        =   48
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame framePush 
      Caption         =   "Momentary buttons"
      Height          =   735
      Left            =   120
      TabIndex        =   49
      Top             =   3960
      Width           =   4095
      Begin VB.CommandButton cmdPush 
         Caption         =   "1"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   50
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdPush 
         Caption         =   "2"
         Height          =   375
         Index           =   2
         Left            =   660
         TabIndex        =   51
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdPush 
         Caption         =   "3"
         Height          =   375
         Index           =   3
         Left            =   1140
         TabIndex        =   52
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdPush 
         Caption         =   "4"
         Height          =   375
         Index           =   4
         Left            =   1620
         TabIndex        =   53
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdPush 
         Caption         =   "5"
         Height          =   375
         Index           =   5
         Left            =   2100
         TabIndex        =   54
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdPush 
         Caption         =   "6"
         Height          =   375
         Index           =   6
         Left            =   2580
         TabIndex        =   55
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdPush 
         Caption         =   "7"
         Height          =   375
         Index           =   7
         Left            =   3060
         TabIndex        =   56
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdPush 
         Caption         =   "8"
         Height          =   375
         Index           =   8
         Left            =   3540
         TabIndex        =   57
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame frameOnOff 
      Caption         =   "ON/OFF"
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   5535
      Begin VB.CommandButton cmdOn 
         Caption         =   "SET ALL"
         Height          =   375
         Index           =   9
         Left            =   4320
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "CLEAR ALL"
         Height          =   375
         Index           =   9
         Left            =   4320
         TabIndex        =   31
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "1"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "O"
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "1"
         Height          =   375
         Index           =   2
         Left            =   660
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "O"
         Height          =   375
         Index           =   2
         Left            =   660
         TabIndex        =   17
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "1"
         Height          =   375
         Index           =   3
         Left            =   1140
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "O"
         Height          =   375
         Index           =   3
         Left            =   1140
         TabIndex        =   19
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "1"
         Height          =   375
         Index           =   4
         Left            =   1620
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "O"
         Height          =   375
         Index           =   4
         Left            =   1620
         TabIndex        =   21
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "1"
         Height          =   375
         Index           =   5
         Left            =   2100
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "O"
         Height          =   375
         Index           =   5
         Left            =   2100
         TabIndex        =   23
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "1"
         Height          =   375
         Index           =   6
         Left            =   2580
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "O"
         Height          =   375
         Index           =   6
         Left            =   2580
         TabIndex        =   25
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "1"
         Height          =   375
         Index           =   7
         Left            =   3060
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "O"
         Height          =   375
         Index           =   7
         Left            =   3060
         TabIndex        =   27
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdOn 
         Caption         =   "1"
         Height          =   375
         Index           =   8
         Left            =   3540
         TabIndex        =   28
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "O"
         Height          =   375
         Index           =   8
         Left            =   3540
         TabIndex        =   29
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Frame frameComPort 
      Caption         =   "Settings"
      Height          =   1215
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtRepeat 
         Appearance      =   0  '2D
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         TabIndex        =   5
         Text            =   "1"
         Top             =   780
         Width           =   555
      End
      Begin VB.TextBox txtComPort 
         Appearance      =   0  '2D
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         TabIndex        =   2
         Text            =   "1"
         Top             =   300
         Width           =   555
      End
      Begin VB.CommandButton cmdComPort 
         Appearance      =   0  '2D
         Caption         =   "Open Port"
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblSettings 
         Caption         =   "Command"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   6
         Top             =   840
         Width           =   915
      End
      Begin VB.Label lblSettings 
         Caption         =   "Repeat"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblSettings 
         Caption         =   "Com Port"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.Timer buttonTimer 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   0
      Top             =   1980
   End
   Begin MSCommLib.MSComm comm 
      Left            =   0
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label lblRelay 
      Caption         =   "8"
      Height          =   255
      Index           =   8
      Left            =   3780
      TabIndex        =   39
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label lblRelay 
      Caption         =   "7"
      Height          =   255
      Index           =   7
      Left            =   3300
      TabIndex        =   38
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label lblRelay 
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   2820
      TabIndex        =   37
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblRelay 
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   2340
      TabIndex        =   36
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblRelay 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   1860
      TabIndex        =   35
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblRelay 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   1380
      TabIndex        =   34
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblRelay 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   900
      TabIndex        =   33
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label lblRelay 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   32
      Top             =   2760
      Width           =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SendCommand As String
Private ShowCommand As String   ' Only 1 Sequence
Private Address As Integer
Private AppTitle As String

Private WithEvents K8056 As clsK8056
Attribute K8056.VB_VarHelpID = -1

' kann Fokus auf Control setzen, das noch nicht geladen ist
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
  
Private Sub EnableControls(Status As Boolean)

  Dim i As Integer
  
  cmdChangeAddress.Enabled = Status
  cmdDisplayAddress.Enabled = Status
  cmdEmergency.Enabled = Status
  cmdResetAddress.Enabled = Status
  cmdSetAddress.Enabled = Status
  
  For i = 1 To 9
    If i < 9 Then
      cmdToggle(i).Enabled = Status
      cmdPush(i).Enabled = Status
    End If
    cmdOn(i).Enabled = Status
    cmdOff(i).Enabled = Status
  Next i
  
  txtAddress.Enabled = Status
  txtRepeat.Enabled = Status
  
  If Status = False Then
    SetFocusAPI txtComPort.hwnd
  End If
  
End Sub

Private Sub InitDevice()
  'beim Starten die Einstellungen setzen
  
  Dim tmpStr As String
  
  myIniFile = App.Path & "\" & App.EXEName & ".ini"
  Set K8056 = New clsK8056
  
  SetFormPosition Me
  EnableControls False
  
  AppTitle = Me.Caption
  
  tmpStr = INIGetSettingString("Settings", "Repeat", myIniFile)
  If tmpStr <> "" Then
    txtRepeat.Text = tmpStr
    K8056.Repeat = Val(tmpStr)
  Else
    txtRepeat.Text = "2"
    K8056.Repeat = 2
  End If
  tmpStr = INIGetSettingString("Settings", "Address", myIniFile)
  If tmpStr <> "" Then
    txtAddress.Text = tmpStr
    Address = Val(tmpStr)
  Else
     Address = 1
    txtAddress.Text = "1"
  End If
  tmpStr = INIGetSettingString("Settings", "Port", myIniFile)
  If tmpStr <> "" Then txtComPort.Text = tmpStr
  
  SetTitle
  
End Sub

Private Sub SetTitle()

  Me.Caption = AppTitle & " - Device #" & AppInstanceNumber & " " & IIf(comm.PortOpen, "Online", "Offline") & " [Card #" & Format(Address, "000") & "]"

End Sub

Private Sub chkRelay_Click(Index As Integer)

  Dim i As Long
  Dim tmpVal As Long
   
  tmpVal = 0
  
  For i = 0 To 7
    If chkRelay(i).Value = vbChecked Then
      tmpVal = tmpVal + 2 ^ i
    End If
  Next i
  
  txtCalcAddress.Text = tmpVal
  
End Sub

Private Sub cmdChangeAddress_Click()

  Dim NewAddress As Integer
  
  NewAddress = Val(txtAddress.Text)
  If NewAddress < 1 Or NewAddress > 255 Then
    NewAddress = 1
    txtAddress.Text = CStr(NewAddress)
  End If
  
  If Address <> NewAddress Then
    Call K8056.SetAddress(Address, NewAddress)
    Address = NewAddress
    INIWriteSetting "Settings", "Address", CStr(Address), myIniFile
    SetTitle
  End If
  
End Sub

Private Sub cmdClear_Click()
  txtDataHex.Text = ""
End Sub

Private Sub cmdComPort_Click()

  On Error GoTo ErrHandler
  
  If comm.PortOpen Then
    comm.PortOpen = False
    txtComPort.Enabled = True
    EnableControls False
    cmdComPort.Caption = "Open Port"
    SetTitle
  Else
    comm.CommPort = txtComPort.Text
    comm.Handshaking = comNone
    comm.Settings = "2400,N,8,1"
    comm.OutBufferSize = 4096
    comm.InputLen = 0
    comm.RThreshold = 1
    comm.SThreshold = 1
    comm.DTREnable = True
    comm.PortOpen = True
    txtComPort.Enabled = False
    EnableControls True
    cmdComPort.Caption = "Close Port"
    SetTitle
    INIWriteSetting "Settings", "Port", txtComPort.Text, myIniFile
  End If
  
  Exit Sub
  
ErrHandler:
  If comm.PortOpen = True Then
    comm.PortOpen = False
    txtComPort.Enabled = True
    cmdComPort.Caption = "Open Port"
    EnableControls True
    SetTitle
  End If
  MsgboxEx Err.Number & ", " & Err.Description, vbExclamation, "Error", , , eCenterDialog
  
End Sub

Private Sub cmdDisplayAddress_Click()
  Call K8056.DisplayAddress
End Sub

Private Sub cmdEmergency_Click()
  Call K8056.EmergencyStop
End Sub

Private Sub cmdEmergency_GotFocus()
  buttonTimer.Enabled = False
End Sub

Private Sub cmdOff_Click(Index As Integer)
  Call K8056.RelayOff(Address, Index)
End Sub

Private Sub cmdOff_GotFocus(Index As Integer)
  buttonTimer.Enabled = False
End Sub

Private Sub cmdOn_Click(Index As Integer)
  Call K8056.RelayOn(Address, Index)
End Sub

Private Sub cmdOn_GotFocus(Index As Integer)
  buttonTimer.Enabled = False
End Sub

Private Sub cmdPush_LostFocus(Index As Integer)
  buttonTimer.Enabled = False
End Sub

Private Sub cmdResetAddress_Click()
  Call K8056.ResetAddress
  Address = 1
  txtAddress.Text = "1"
  INIWriteSetting "Settings", "Address", CStr(Address), myIniFile
  SetTitle
End Sub

Private Sub cmdSetAddress_Click()
  Address = Val(txtAddress.Text)
  If Address < 1 Or Address > 255 Then
    Address = 1
    txtAddress.Text = CStr(Address)
  End If
  INIWriteSetting "Settings", "Address", CStr(Address), myIniFile
  SetTitle
End Sub

Private Sub cmdToggle_Click(Index As Integer)
  ' Index 1 - 8
  Call K8056.Toggle(Address, Index)
End Sub

Private Sub cmdPush_Click(Index As Integer)
  ' Index 1 - 8
  cmdEmergency.SetFocus
  buttonTimer.Enabled = False
End Sub

Private Sub cmdPush_GotFocus(Index As Integer)
  ' Index 1 - 8
  Call K8056.Push(Address, Index)
  buttonTimer.Enabled = True
End Sub

Private Sub cmdToggle_GotFocus(Index As Integer)
  buttonTimer.Enabled = False
End Sub

Private Sub Form_Load()
  InitDevice
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  comm.DTREnable = False
  If comm.PortOpen = True Then comm.PortOpen = False
  
  SaveFormPosition Me
  'Instanz als frei deklarieren
  INIWriteSetting "App", "#" & CStr(AppInstanceNumber), "0", myIniFile
  
  Set K8056 = Nothing
  
End Sub

Private Sub buttonTimer_Timer()
  If comm.PortOpen = True Then
    printHexData ShowCommand
    comm.Output = SendCommand
  Else
    buttonTimer.Enabled = False
  End If
End Sub

Private Sub printHexData(Data As String)

  Dim hOut As String
  Dim tmpHex As String
  Dim i As Integer

  For i = 1 To Len(Data)
    tmpHex = Hex(Asc(Mid(Data, i, 1)))
    If Len(tmpHex) < 2 Then tmpHex = "0" & tmpHex
    If hOut <> "" Then hOut = hOut & " "
    hOut = hOut & tmpHex
  Next i
  txtDataHex.Text = txtDataHex.Text & hOut & vbCrLf
  If Len(txtDataHex.Text) > 65000 Then
    'Inhalt kann max 64KB aufnehmen, daher dann den Anfang immer wieder kürzen
    i = InStr(txtDataHex.Text, vbCrLf)
    If i > 0 Then
      txtDataHex.Text = Mid(txtDataHex.Text, i + 1)
    End If
  End If
  txtDataHex.SelLength = 0
  txtDataHex.SelStart = Len(txtDataHex.Text)
  
End Sub

Private Sub printComment(Data As String)
  
  Dim i As Integer

  txtDataHex.Text = txtDataHex.Text & Data & vbCrLf
  If Len(txtDataHex.Text) > 65000 Then
    'Inhalt kann max 64KB aufnehmen, daher dann den Anfang immer wieder kürzen
    i = InStr(txtDataHex.Text, vbCrLf)
    If i > 0 Then
      txtDataHex.Text = Mid(txtDataHex.Text, i + 1)
    End If
  End If
  txtDataHex.SelLength = 0
  txtDataHex.SelStart = Len(txtDataHex.Text)
  
End Sub

Private Sub K8056_commOutput(Message As String)

  On Error GoTo ErrHandler
  
  If comm.PortOpen Then
    SendCommand = Message
    comm.Output = Message
  End If
  
  Exit Sub
  
ErrHandler:
  MsgboxEx Err.Number & ", " & Err.Description, vbExclamation, "Error", , , eCenterDialog
  
End Sub

Private Sub K8056_Error(Message As String)
  MsgboxEx Message, , "Error", vbExclamation, , eCenterDialog
End Sub

Private Sub K8056_printComment(Message As String)
  printComment Message
End Sub

Private Sub K8056_printHexData(Message As String)
  ShowCommand = Message
  printHexData Message
End Sub

Private Sub txtAddress_Change()
  If Val(txtAddress.Text) < 1 Or Val(txtAddress.Text) > 255 Then
    MsgboxEx "we can only handle addresses from 1 to 255", vbInformation, "Notice", , , eCenterDialog
    txtAddress.Text = "1"
  End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
  NumericTextBox KeyAscii
End Sub

Private Sub txtCalcAddress_DblClick()

  Dim i As Long
  
  For i = 0 To 7
    chkRelay(i).Value = vbUnchecked
  Next i
  
End Sub

Private Sub txtComPort_Change()
  If Val(txtComPort.Text) < 1 Or Val(txtComPort.Text) > 16 Then
    MsgboxEx "we can only handle ports from 1 to 16", vbInformation, "Notice", , , eCenterDialog
    txtComPort.Text = "1"
  End If
End Sub

Private Sub txtComPort_KeyPress(KeyAscii As Integer)
  NumericTextBox KeyAscii
End Sub

Private Sub NumericTextBox(KeyAscii As Integer)
  Select Case KeyAscii
    'check if key pressed is number, ., - or backspace
    Case Asc("0") To Asc("9"), vbKeyBack, vbKeyDelete
      'don't do anything
    Case Else
      'set the keyascii to 0
      KeyAscii = 0
  End Select
End Sub

Private Sub txtRepeat_KeyPress(KeyAscii As Integer)
  NumericTextBox KeyAscii
End Sub

Private Sub txtRepeat_LostFocus()
  K8056.Repeat = Val(txtRepeat.Text)
  INIWriteSetting "Settings", "Repeat", txtRepeat.Text, myIniFile
End Sub
