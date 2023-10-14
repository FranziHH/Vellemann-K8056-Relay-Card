Attribute VB_Name = "modMessageBoxEx"
Option Explicit

'**********************************************************************
' * Comments         : Controlling the position of a MsgBox
' *
' * You can create a CBT hook for your application so that it receives
' * notifications when windows are created and destroyed. If you
' * display a message box with this CBT hook in place, your application
' * will receive a HCBT_ACTIVATE message when the message box is
' * activated. Once you receive this HCBT_ACTIVATE message, you can
' * position the window with the SetWindowPos API function and then
' * release the CBT hook if it is no longer needed. See the "Test"
' * routine for a demonstration.
' *
' **********************************************************************

'PLACE CODE IN A STANDARD MODULE

Public Enum ePosMsgBox
   eTopLeft
   eTopRight
   eTopCenter
   eBottomLeft
   eBottomRight
   eBottomCenter
   eCenterScreen
   eCenterDialog
End Enum

Private Type RECT
   Left                 As Long
   Top                  As Long
   Right                As Long
   Bottom               As Long
End Type

'Message API and constants
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal zlhHook As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Const GWL_HINSTANCE = (-6)
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const HCBT_ACTIVATE = 5
Private Const WH_CBT = 5

'Other APIs
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Private zlhHook         As Long
Private zePosition      As ePosMsgBox

'Purpose : Displays a msgbox at a specified location on the screen
'Inputs : As per a standard MsgBox +
' Position An enumerated type which controls the screen position of the MsgBox
'Outputs : As per a standard Msgbox
'Author : Andrew Baker
'Date : 25/05/2001
'Notes :

Function MsgboxEx(Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title As String, Optional HelpFile As String, Optional Context As Long, Optional Position As ePosMsgBox = eCenterScreen) As VbMsgBoxResult
   Dim lhInst           As Long
   Dim lThread          As Long

   'Set up the CBT hook
   lhInst = GetWindowLong(GetForegroundWindow, GWL_HINSTANCE)
   lThread = GetCurrentThreadId()
   zlhHook = SetWindowsHookEx(WH_CBT, AddressOf zWindowProc, lhInst, lThread)

   zePosition = Position

   'Display the message box
   MsgboxEx = MsgBox(Prompt, Buttons, Title, HelpFile, Context)
End Function

'Call back used by MsgboxEx
Function zWindowProc(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim tFormPos         As RECT, tMsgBoxPos As RECT, tScreenWorkArea As RECT
   Dim lLeft            As Long, lTop As Long

   If lMsg = HCBT_ACTIVATE Then
      On Error Resume Next
      'A new dialog has been displayed
      tScreenWorkArea = ScreenWorkArea
      'Get the coordinates of the form and the message box so that
      'you can determine where the center of the form is located
      GetWindowRect GetForegroundWindow, tFormPos
      GetWindowRect wParam, tMsgBoxPos

      Select Case zePosition
         Case eCenterDialog
            lLeft = (tFormPos.Left + (tFormPos.Right - tFormPos.Left) / 2) - ((tMsgBoxPos.Right - tMsgBoxPos.Left) / 2)
            lTop = (tFormPos.Top + (tFormPos.Bottom - tFormPos.Top) / 2) - ((tMsgBoxPos.Bottom - tMsgBoxPos.Top) / 2)

         Case eCenterScreen
            lLeft = ((tScreenWorkArea.Right - tScreenWorkArea.Left) - (tMsgBoxPos.Right - tMsgBoxPos.Left)) / 2
            lTop = ((tScreenWorkArea.Bottom - tScreenWorkArea.Top) - (tMsgBoxPos.Bottom - tMsgBoxPos.Top)) / 2

         Case eTopLeft
            lLeft = tScreenWorkArea.Left
            lTop = tScreenWorkArea.Top

         Case eTopRight
            lLeft = tScreenWorkArea.Right - (tMsgBoxPos.Right - tMsgBoxPos.Left)
            lTop = tScreenWorkArea.Top

         Case eTopCenter
            lLeft = ((tScreenWorkArea.Right - tScreenWorkArea.Left) - (tMsgBoxPos.Right - tMsgBoxPos.Left)) / 2
            lTop = tScreenWorkArea.Top

         Case eBottomLeft
            lLeft = tScreenWorkArea.Left
            lTop = tScreenWorkArea.Bottom - (tMsgBoxPos.Bottom - tMsgBoxPos.Top)

         Case eBottomRight
            lLeft = tScreenWorkArea.Right - (tMsgBoxPos.Right - tMsgBoxPos.Left)
            lTop = tScreenWorkArea.Bottom - (tMsgBoxPos.Bottom - tMsgBoxPos.Top)

         Case eBottomCenter
            lLeft = ((tScreenWorkArea.Right - tScreenWorkArea.Left) - (tMsgBoxPos.Right - tMsgBoxPos.Left)) / 2
            lTop = tScreenWorkArea.Bottom - (tMsgBoxPos.Bottom - tMsgBoxPos.Top)

      End Select

      'Position the msgbox
      SetWindowPos wParam, 0, lLeft, lTop, 10, 10, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE

      'Release the CBT hook
      UnhookWindowsHookEx zlhHook
   End If
   zWindowProc = False

End Function

'Purpose : Returns the screen dimensions, not including the tastbar
'Inputs : N/A
'Outputs : A type which defines the extent of the screen work area.
'Author : Andrew Baker
'Date : 25/05/2001
'Notes :

Function ScreenWorkArea() As RECT
   Dim tScreen          As RECT
   Dim lRet             As Long
   Const SPI_GETWORKAREA = 48

   lRet = SystemParametersInfo(SPI_GETWORKAREA, vbNull, tScreen, 0)
   ScreenWorkArea = tScreen
End Function

'Demonstration routine
Sub Test()
   MsgboxEx "Hello BottomCenter", , , , , eBottomCenter
   MsgboxEx "Hello BottomLeft", , , , , eBottomLeft
   MsgboxEx "Hello BottomRight", , , , , eBottomRight
   MsgboxEx "Hello CenterDialog", , , , , eCenterDialog
   MsgboxEx "Hello CenterScreen", , , , , eCenterScreen
   MsgboxEx "Hello TopCenter", , , , , eTopCenter
   MsgboxEx "Hello TopLeft", , , , , eTopLeft
   MsgboxEx "Hello TopRight", , , , , eTopRight
End Sub

