Attribute VB_Name = "modINIFile"
Option Explicit

Private Declare Function WritePrivateProfileString _
  Lib "kernel32" _
  Alias "WritePrivateProfileStringA" _
  (ByVal strSection As String, _
   ByVal strKeyNam As String, _
   ByVal strValue As String, _
   ByVal strFileName As String) _
As Long

Private Declare Function GetPrivateProfileString _
  Lib "kernel32" _
  Alias "GetPrivateProfileStringA" _
  (ByVal strSection As String, _
   ByVal strKeyName As String, _
   ByVal strDefault As String, _
   ByVal strReturned As String, _
   ByVal lngSize As Long, _
   ByVal strFileName As String) _
As Long

Public myIniFile As String

Private Declare Function GetVersion Lib "kernel32" () As Long

Private Declare Function EnumDisplayMonitors Lib "user32" ( _
      ByVal hDC As Long, _
      lprcClip As Any, _
      ByVal lpfnEnum As Long, _
      ByVal dwData As Long _
   ) As Long
   
Private m_cM As clsMonitors

Private Function MonitorEnumProc( _
      ByVal hMonitor As Long, _
      ByVal hDCMonitor As Long, _
      ByVal lprcMonitor As Long, _
      ByVal dwData As Long _
   ) As Long
   m_cM.fAddMonitor hMonitor
   MonitorEnumProc = 1
End Function

Public Sub EnumMonitors(cM As clsMonitors)
   Set m_cM = cM
   EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, 0
End Sub

Public Function IsNt() As Boolean
Dim lVer As Long
   lVer = GetVersion()
   IsNt = ((lVer And &H80000000) = 0)
End Function


Public Function INIGetSettingString( _
  strSection As String, _
  strKeyName As String, _
  strFile As String) _
  As String
  ' Comments  : Returns a string value from the specified INI file
  ' Parameters: strSection - Name of the section to look in
  '             strKeyName - Name of the key to look for
  '             strFile - Path and name of the INI file to look in
  ' Returns   : String value
  ' Source    : Total VB SourceBook 6
  '
  Dim strBuffer As String * 256
  Dim intSize As Integer

  On Error GoTo PROC_ERR
  
  If AppInstanceNumber = 0 Then
    GetAppInstance
  End If
  If strSection <> "App" Then strSection = strSection & "_" & AppInstanceNumber
  
  intSize = GetPrivateProfileString(strSection, strKeyName, "", strBuffer, 256, strFile)

  INIGetSettingString = Left$(strBuffer, intSize)

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "INIGetSettingString"
  Resume PROC_EXIT
  
End Function

Public Function INIWriteSetting( _
  strSection As String, _
  strKeyName As String, _
  strValue As String, _
  strFile As String) _
  As Integer
  ' Comments  : Writes the specified value to the specified INI file
  ' Parameters: strSection - section to write into
  '             strKeyName - key to write into
  '             strValue - value to write
  '             strFile - path and name of the INI file to write to
  ' Returns   : True if successful, False otherwise
  ' Source    : Total VB SourceBook 6
  '
  Dim intStatus As Integer

  On Error GoTo PROC_ERR

  If AppInstanceNumber = 0 Then
    GetAppInstance
  End If
  If strSection <> "App" Then strSection = strSection & "_" & AppInstanceNumber
  
  intStatus = WritePrivateProfileString( _
    strSection, _
    strKeyName, _
    strValue, _
    strFile)

  INIWriteSetting = (intStatus <> 0)

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "INIWriteSetting"
  Resume PROC_EXIT
  
End Function

Public Sub SetFormPosition(frm As Form)

  Dim cMonFrom As clsMonitor
  Dim tmpStr As String
  Dim isMonitor As Boolean
  Dim newMonitor As String
  Dim newLeft As Long
  Dim newTop As Long
  Dim i As Long
  Dim lTop As Long
  Dim lLeft As Long
  Dim m_cM As clsMonitors
  Dim isMoved As Boolean
  
  Set m_cM = New clsMonitors
  Set cMonFrom = m_cM.MonitorForWindow(frm.hwnd)
  
  isMoved = False
  
  tmpStr = INIGetSettingString("Settings", "Monitor", myIniFile)
  If tmpStr <> "" Then
    newMonitor = tmpStr
  Else
    newMonitor = ""
  End If
  
  tmpStr = INIGetSettingString("Settings", "Top", myIniFile)
  If tmpStr <> "" Then
    newTop = CLng(tmpStr)
  Else
    newTop = 0
  End If
  
  tmpStr = INIGetSettingString("Settings", "Left", myIniFile)
  If tmpStr <> "" Then
    newLeft = CLng(tmpStr)
  Else
    newLeft = 0
  End If
  
  If newMonitor <> "" And newTop > 0 And newLeft > 0 Then
    'nur etwas tun, wenn auch was gespeichert wurde
    isMonitor = False
    
    For i = 1 To m_cM.MonitorCount
      If m_cM.Monitor(i).Name = newMonitor Then
        isMonitor = True
        Exit For
      End If
    Next i
    
    If isMonitor Then
      'die Form jetzt platzieren - korrekter monitor wurde gefunden
      lTop = frm.ScaleY(newTop, frm.ScaleMode, vbPixels) + 30
      lLeft = frm.ScaleX(newLeft, frm.ScaleMode, vbPixels) + 30
      
      If newMonitor <> cMonFrom.Name Then
         lLeft = lLeft - cMonFrom.Width
         lTop = lTop - cMonFrom.Top
      End If

      If lLeft < m_cM.Monitor(i).Width And lTop < m_cM.Monitor(i).Height Then
        'sicherstellen, dass wenn positioniert wird, im sichtbaren bereich ...
        frm.Move newLeft, newTop
        isMoved = True
      End If
    End If
  End If

  If Not isMoved Then
    'auf Desktopmitte setzen ...
    frm.Left = (Screen.Width - frm.Width) / 2
    frm.Top = (Screen.Height - frm.Height) / 2
  End If
  
  Set m_cM = Nothing
  Set cMonFrom = Nothing
  
End Sub
Public Sub SaveFormPosition(frm As Form)

  Dim m_cM As clsMonitors
  Dim Monitor As clsMonitor
  
  Set m_cM = New clsMonitors
  Set Monitor = m_cM.MonitorForWindow(frm.hwnd)

  INIWriteSetting "Settings", "Monitor", Monitor.Name, myIniFile
  INIWriteSetting "Settings", "Top", CStr(frm.Top), myIniFile
  INIWriteSetting "Settings", "Left", CStr(frm.Left), myIniFile
  
  Set m_cM = Nothing
  Set Monitor = Nothing
  
End Sub
