Attribute VB_Name = "EsiExit"
'2/19/05 New
Option Explicit
'Windows closing functions
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const hWnd_TopMost = -1

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOW = 5

Public Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Const WM_CLOSE = &H10

'Users
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Const ES_RED = &HC0 'ESI red for labels
Public Const ES_BLUE = &H800000 'ESI blue for labels

Public bUserAction As Byte
Public sAppTitle As String

Public sWorkStation As String * 20
Public sUserName As String * 20

Public Function GetWorkStation() As String
   Dim intZeroPos As Integer
   Dim gintMAX_SIZE As Long
   Dim strBuf As String
   gintMAX_SIZE = 255 'Maximum buffer size
   
   strBuf = Space$(gintMAX_SIZE)
   On Error GoTo ModErr1
   
   'Get the workstation and then trim the buffer to the exact length
   'returned and add a dir sep (backslash) if the API didn't return one
   '
   If GetComputerName(strBuf, gintMAX_SIZE) > 0 Then
      intZeroPos = InStr(strBuf, Chr$(0))
      If intZeroPos > 0 Then strBuf = Left$(strBuf, intZeroPos - 1)
      GetWorkStation = strBuf
   Else
      GetWorkStation = ""
   End If
   Exit Function
   
ModErr1:
   On Error GoTo 0
   GetWorkStation = ""
   
End Function


Public Function GetNetId() As String
   Dim intZeroPos As Integer
   Dim gintMAX_SIZE As Long
   Dim strBuf As String
   gintMAX_SIZE = 255 'Maximum buffer size
   
   strBuf = Space$(gintMAX_SIZE)
   '
   'Get the Net user login and then trim the buffer to the exact length
   'returned and add a dir sep (backslash) if the API didn't return one
   
   On Error GoTo ModErr1
   If GetUserName(strBuf, gintMAX_SIZE) > 0 Then
      intZeroPos = InStr(strBuf, Chr$(0))
      If intZeroPos > 0 Then strBuf = Left$(strBuf, intZeroPos - 1)
      GetNetId = strBuf
   Else
      GetNetId = ""
   End If
   Exit Function
   
ModErr1:
   On Error GoTo 0
   GetNetId = ""
   
End Function

'Syntax OpenWebHelp hWnd, "hs4101" where the Context Id (page) is hs4101
'Note temporary URL
'8/1/04 New

Public Sub OpenWebHelp(SendTo As String)
   Dim lTask As Long
   Dim sUrl As String
   If Right$(SendTo, 4) <> ".htm" Then SendTo = SendTo & ".htm"
   sUrl = "http://esisupport.home.comcast.net/"
   SendTo = sUrl & SendTo
   lTask = ShellExecute(diaExit.hWnd, "open", SendTo, _
           vbNullString, vbNullString, SW_SHOW)
   If lTask < 32 Then MsgBox "Error Opening The Requested Page.", _
      vbInformation, "ES/2000 ERP"
   
   '//ShellExecute(
   'hwnd>> handle to parent window
   'lpOperation>> pointer to string that specifies operation to perform
   'lpFile>> pointer to filename or folder name string
   'lpParameters>> pointer to string that specifies executable-file parameters
   'lpDirectory>> pointer to string that specifies default directory
   ' nShowCmd>>whether file is shown when opened
   ')
   
End Sub

'Closes the rest. Won't work if not open

Public Sub CloseRemainder()
   Dim iFreeFile As Integer
   Dim lClose As Long
   Dim sFilePath As String
   
   On Error Resume Next
   diaExit.Timer2.Enabled = False
   iFreeFile = FreeFile
   
   lClose = FindWindow(vbNullString, "ESI2000")
   If lClose > 0 Then SendMessage lClose, WM_CLOSE, 0&, 0&
   Sleep 1000
   diaExit.Height = 3700
   diaExit.lblShutDown.Caption = diaExit.lblTime & " All open Sections " _
                                 & "and Manager Were Closed."
   diaExit.lblShutDown.Refresh
   
   sFilePath = GetSetting("Esi2000", "System", "FilePath", sFilePath)
   Open sFilePath & "EsiClose.log" For Append Shared As iFreeFile
   lClose = LOF(iFreeFile)
   If lClose = 0 Then
      Print #iFreeFile, "ES/2005 ERP closed the following Workstations due to excessive idle time: "
      Print #iFreeFile, vbCr
      Print #iFreeFile, "Workstation        ", "Windows Log On"
      Print #iFreeFile, String$(110, "-")
   End If
   sWorkStation = GetWorkStation()
   sUserName = GetNetId()
   Print #iFreeFile, sWorkStation, sUserName, Format(Now, "mm/dd/yy hh:mm AMPM")
   Close
   
End Sub

Public Sub CloseCurrentApp()
   Static b As Byte
   Dim lClose As Long
   If b > 7 Then Exit Sub
   b = b + 1
   'Twice in case of a MessageBox showing
   lClose = FindWindow(vbNullString, sAppTitle)
   If lClose > 0 Then SendMessage lClose, WM_CLOSE, 0&, 0&
   Sleep 1000
   
   lClose = FindWindow(vbNullString, sAppTitle)
   If lClose > 0 Then SendMessage lClose, WM_CLOSE, 0&, 0&
   
   diaExit.Section(b) = sAppTitle
   diaExit.OffTime(b) = Format(Now, "hh:mm AMPM")
   diaExit.Timer3.Enabled = True
   bUserAction = 1
   
End Sub
