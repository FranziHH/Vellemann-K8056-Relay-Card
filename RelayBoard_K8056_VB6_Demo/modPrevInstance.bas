Attribute VB_Name = "modPrevInstance"
' -----------------------------------------------------------------
' Liste aktiver Anwendungsprozesse ermitteln
' Copyright © Mathias Schiffer 1999-2005
' -----------------------------------------------------------------
'
' KURZE FUNKTIONSBESCHREIBUNG:
'
' - Public Function GetProcessCollection() As Collection
'   Gibt eine String-Collection zurück, deren Einträge den
'   Aufbau "Prozessname|Prozess-ID" haben.
'
' - Public Function ProcessName(ByVal CollectionString As String) As String
'   Extrahiert aus einem String der Collection den Prozessnamen.
'
' - Public Function ProcessHandle(ByVal CollectionString As String) As Long
'   Extrahiert aus einem String der Collection die Prozess-ID.
'
' - Public Function KillProcessByPID(ByVal PID As Long) As Boolean
'   Terminiert einen Prozess auf Basis seiner Prozess-ID. Ein Prozess
'   sollte nur in "Notfällen" terminiert werden. Datenverluste der
'   terminierten Anwendung sind nicht ausgeschlossen.
'
' -----------------------------------------------------------------
  
Option Explicit

Public AppInstanceNumber As Long

' ------------------------- DEKLARATIONEN -------------------------
  
' Deklaration notwendiger API-Funktionen:
  
' GetVersionEx dient der Erkennung des Betriebssystems:
Private Declare Function GetVersionEx _
  Lib "kernel32" Alias "GetVersionExA" ( _
  ByRef lpVersionInformation As OSVERSIONINFO _
  ) As Long
' Toolhelp-Funktionen zur Prozessauflistung (Win9x):
Private Declare Function CreateToolhelp32Snapshot _
  Lib "kernel32" ( _
  ByVal dwFlags As Long, _
  ByVal th32ProcessID As Long _
  ) As Long
Private Declare Function Process32First _
  Lib "kernel32" ( _
  ByVal hSnapshot As Long, _
  ByRef lppe As PROCESSENTRY32 _
  ) As Long
Private Declare Function Process32Next _
  Lib "kernel32" ( _
  ByVal hSnapshot As Long, _
  ByRef lppe As PROCESSENTRY32 _
  ) As Long
' PSAPI-Funktionen zur Prozessauflistung (Windows NT)
Private Declare Function EnumProcesses _
  Lib "psapi.dll" ( _
  ByRef lpidProcess As Long, _
  ByVal cb As Long, _
  ByRef cbNeeded As Long _
  ) As Long
Private Declare Function GetModuleFileNameEx _
  Lib "psapi.dll" Alias "GetModuleFileNameExA" ( _
  ByVal hProcess As Long, _
  ByVal hModule As Long, _
  ByVal ModuleName As String, _
  ByVal nSize As Long _
  ) As Long
Private Declare Function EnumProcessModules _
  Lib "psapi.dll" ( _
  ByVal hProcess As Long, _
  ByRef lphModule As Long, _
  ByVal cb As Long, _
  ByRef cbNeeded As Long _
  ) As Long
' Win32-API-Funktionen für Prozessmanagement
Private Declare Function OpenProcess _
  Lib "Kernel32.dll" ( _
  ByVal dwDesiredAccess As Long, _
  ByVal bInheritHandle As Long, _
  ByVal dwProcId As Long _
  ) As Long
Private Declare Function TerminateProcess _
  Lib "kernel32" ( _
  ByVal hProcess As Long, _
  ByVal uExitCode As Long _
  ) As Long
Private Declare Function CloseHandle _
  Lib "Kernel32.dll" ( _
  ByVal Handle As Long _
  ) As Long
  
' Deklaration notwendiger Konstanter:
  
Private Const MAX_PATH                  As Long = 260
Private Const PROCESS_QUERY_INFORMATION As Long = 1024
Private Const PROCESS_VM_READ           As Long = 16
Private Const STANDARD_RIGHTS_REQUIRED  As Long = &HF0000
Private Const SYNCHRONIZE               As Long = &H100000
Private Const PROCESS_ALL_ACCESS        As Long = STANDARD_RIGHTS_REQUIRED _
                                               Or SYNCHRONIZE Or &HFFF
Private Const TH32CS_SNAPPROCESS        As Long = &H2&
  
' Konstante für die Erkennung des Betriebssystems:
Private Const VER_PLATFORM_WIN32_NT     As Long = 2
  
' Notwendige Typdeklarationen
  
Private Type PROCESSENTRY32 ' Prozesseintrag
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long
   th32DefaultHeapID As Long
   th32ModuleID     As Long
   cntThreads As Long
   th32ParentProcessID As Long
   pcPriClassBase As Long
   dwFlags As Long
   szExeFile As String * MAX_PATH ' = 260
End Type
  
Private Type OSVERSIONINFO ' Betriebssystemerkennung
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Public Sub GetAppInstance()

  'ermittelt die aktuelle Instanz
  Dim tmpAppName As String
  Dim myColl As Collection
  Dim i As Long
  Dim tmpStr As String
  Dim newNr As Long
  
  If isEXE Then
    'wenn kompiliert, die Nummer der aktiven Instanzen ermitteln
    tmpAppName = LCase(App.Path & "\" & App.EXEName & ".exe")
    AppInstanceNumber = 0
    
    Set myColl = GetProcessCollection
    
    For i = 1 To myColl.Count
      If LCase(ProcessName(myColl(i))) = tmpAppName Then
        AppInstanceNumber = AppInstanceNumber + 1
      End If
    Next i

    INIWriteSetting "App", "Instance", CStr(AppInstanceNumber), myIniFile
    
    'über INI Eintrag prüfen, ob eine "niedrigere" Instanz nicht läuft
    newNr = 0
    For i = 1 To AppInstanceNumber
      tmpStr = INIGetSettingString("App", "#" & CStr(i), myIniFile)
      If Val(tmpStr) = 0 Then
        'diese Instanz ist frei
        newNr = i
        Exit For
      End If
    Next i
    
    If newNr > 0 Then
      'eine niedrigere Instanz ist frei
      AppInstanceNumber = newNr
    End If
    
  Else
    'aus der IDE kann nur eine Instanz gestartet werden
    AppInstanceNumber = 1
  End If

  INIWriteSetting "App", "#" & CStr(AppInstanceNumber), "1", myIniFile
    
End Sub


Public Function GetProcessCollection() As Collection
' Ermittelt die abfragbaren laufenden Prozesse des lokalen
' Rechners. Jeder gefundene Prozess wird mit seiner ID
' als String in einem Element der Rückgabe-Collection
' gespeichert im Format "Prozessname|Prozess-ID".
Dim collProcesses As New Collection
Dim ProcID As Long
  
  If (Not IsWindowsNT) Then
  
    ' WINDOWS 95 / 98 / Me
    ' --------------------
  
    Dim sName As String
    Dim hSnap As Long
    Dim pEntry As PROCESSENTRY32
  
    ' Einen Snapshot der Prozessinformationen erstellen
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If hSnap = 0 Then
      Exit Function ' Pech gehabt
    End If
  
    pEntry.dwSize = Len(pEntry) ' Größe der Struktur zur Verfügung stellen
  
    ' Den ersten Prozess im Snapshot ermitteln
    ProcID = Process32First(hSnap, pEntry)
  
    ' Mittels Process32Next über alle weiteren Prozesse iterieren
    Do While (ProcID <> 0) ' Gibt es eine gültige Prozess-ID?
      sName = TrimNullChar(pEntry.szExeFile)  ' Rückgabestring stutzen
      collProcesses.Add sName & "|" & CStr(ProcID) ' Collection-Eintrag
      ProcID = Process32Next(hSnap, pEntry)   ' Nächste PID des Snapshots
    Loop
  
    ' Handle zum Snapshot freigeben
    CloseHandle hSnap
  
  Else
  
    ' WINDOWS NT / 2000 / XP / 2003 / Vista
    ' -------------------------------------
  
    Dim cb As Long
    Dim cbNeeded As Long
    Dim RetVal As Long
    Dim NumElements As Long
    Dim ProcessIDs() As Long
    Dim cbNeeded2 As Long
    Dim NumElements2 As Long
    Dim Modules(1) As Long
    Dim ModuleName As String
    Dim LenName As Long
    Dim hProcess As Long
    Dim i As Long
  
    cb = 8         ' "CountBytes": Größe des Arrays (in Bytes)
    cbNeeded = 9   ' cbNeeded muss initial größer als cb sein
  
    ' Schrittweise an die passende Größe des Prozess-ID-Arrays
    ' heranarbeiten. Dazu vergößern wir das Array großzügig immer
    ' weiter, bis der zur Verfügung gestellte Speicherplatz (cb)
    ' den genutzten (cbNeeded) überschreitet:
    Do While cb <= cbNeeded ' Alle Bytes wurden belegt -
                            ' es könnten also noch mehr sein
      cb = cb * 2                      ' Speicherplatz verdoppeln
      ReDim ProcessIDs(cb / 4) As Long ' Long = 4 Bytes
      EnumProcesses ProcessIDs(1), cb, cbNeeded ' Array abholen
    Loop
  
    ' In cbNeeded steht der übergebene Speicherplatz in Bytes.
    ' Da jedes Element des Arrays als Long aus 4 Bytes besteht,
    ' ermitteln wir die Anzahl der tatsächlich übergebenen
    ' Elemente durch entsprechende Division:
    NumElements = cbNeeded / 4
  
    ' Jede gefundene Prozess-ID des Arrays abarbeiten
    For i = 1 To NumElements
  
      ' Versuchen, den Prozess zu öffnen und ein Handle zu erhalten
      hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
                          Or PROCESS_VM_READ, _
                             0, ProcessIDs(i))
  
      If (hProcess <> 0) Then ' OpenProcess war erfolgreich
  
        ' EnumProcessModules gibt die dem Prozess angehörenden
        ' Module in einem Array zurück.
        RetVal = EnumProcessModules(hProcess, Modules(1), _
                                    1, cbNeeded2)
  
        If (RetVal <> 0) Then ' EnumProcessModules war erfolgreich
          ModuleName = Space$(MAX_PATH) ' Speicher reservieren
          ' Den Pfadnamen für das erste gefundene Modul bestimmen
          LenName = GetModuleFileNameEx(hProcess, Modules(1), _
                                        ModuleName, Len(ModuleName))
          ' Den gefundenen Pfad und die Prozess-ID unserer
          ' ProcessCollection hinzufügen (Trennzeichen "|")
          collProcesses.Add Left$(ModuleName, LenName) & "|" & _
                            CStr(ProcessIDs(i))
        End If
  
      End If
  
      CloseHandle hProcess ' Offenes Handle schließen
  
    Next i
  
  End If
  
  ' Zusammengestellte Collection übergeben
  Set GetProcessCollection = collProcesses
  
End Function
  
  
Public Function isEXE() As Boolean

  On Error Resume Next
  
  Debug.Print 1 / 0
  
  If Err.Number Then
    ' Anwendung läuft im Debug-Modus in der VB-IDE
    isEXE = False
  Else
    ' Anwendung läuft als kompilierte Anwendung
    isEXE = True
  End If

End Function

Public Function ProcessName(ByVal CollectionString As String) As String
' Extrahiert aus einem String der Collection den Prozessnamen.
Dim Pos1 As Long
  
  ' Trenner suchen
  Pos1 = InStr(CollectionString, "|")
  ' Wenn Trenner vorhanden, Eintrag zurückgeben (sonst vbNullString)
  If (Pos1 > 0) Then
    ProcessName = Left$(CollectionString, Pos1 - 1)
  End If
  
End Function
  
  
Public Function ProcessHandle(ByVal CollectionString As String) As Long
' Extrahiert aus einem String der Collection die Prozess-ID.
Dim Pos1 As Long
  
  ' Trenner suchen
  Pos1 = InStr(CollectionString, "|")
  ' Wenn Trenner vorhanden, Handle zurückgeben (sonst 0)
  If (Pos1 > 0) And (Len(CollectionString) > Pos1) Then
    ProcessHandle = CLng(Mid$(CollectionString, Pos1 + 1))
  End If
  
End Function
  
  
Public Function KillProcessByPID(ByVal PID As Long) As Boolean
' Versucht auf Basis einer Prozess-ID, den zugehörigen
' Prozess zu terminieren. Im Erfolgsfall wird True zurückgegeben.
Dim hProcess As Long
  
  ' Öffnen des Prozesses über seine Prozess-ID
  hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
  
  ' Gibt es ein Handle, wird der Prozess darüber abgeschossen
  If (hProcess <> 0) Then
    KillProcessByPID = TerminateProcess(hProcess, 1&) <> 0
    CloseHandle hProcess
  End If
  
End Function

Private Function TrimNullChar(ByVal s As String) As String
' Kürzt einen String s bis zum Zeichen vor einem vbNullChar
Dim Pos1 As Long
  
  ' vbNullChar = Chr$(0) im String suchen
  Pos1 = InStr(s, vbNullChar)
  ' Falls vorhanden, den String entsprechend kürzen
  If (Pos1 > 0) Then
    TrimNullChar = Left$(s, Pos1 - 1)
  Else
    TrimNullChar = s
  End If
  
End Function
  
  
Private Function IsWindowsNT() As Boolean
' Gibt True für Windows NT (und 2000, XP, 2003, Vista) zurück
Dim OSInfo As OSVERSIONINFO
  
  With OSInfo
    .dwOSVersionInfoSize = Len(OSInfo)  ' Angabe der Größe dieser Struktur
    .szCSDVersion = Space$(128)         ' Speicherreservierung für Angabe des Service Packs
    GetVersionEx OSInfo                 ' OS-Informationen ermitteln
    IsWindowsNT = (.dwPlatformId = VER_PLATFORM_WIN32_NT) ' für Windows NT
  End With
  
End Function

