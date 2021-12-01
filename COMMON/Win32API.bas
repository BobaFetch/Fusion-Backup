Attribute VB_Name = "Win32API"
Option Explicit

Public Declare Function SQLSetConnectOption Lib "odbc32.dll" _
   (ByVal hdbc&, ByVal fOption%, ByVal vParam As Any) As Integer
Public Const SQL_PRESERVE_CURSORS As Long = 1204
Public Const SQL_PC_ON As Long = 1

'3/7/06 Refresh desktop on closing
Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
   lpRect As Long, ByVal bErase As Long) As Long

'memory/system DLL's
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function CreateDirectoryEx Lib "kernel32" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As Any) As Long
' To make a form active
Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long

'12/5/05 Not in use
'Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
'Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

'Combo Boxes and sample see AddComboStr
Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Const CB_ADDSTRING = &H143
'Public Const CB_SHOWDROPDOWN = &H14F
'
'Mail, Web, Apps etc
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOW = 5

'Lock Windows
Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long

'Version (Platform)
Public Declare Function GetVersionEx Lib "kernel32" Alias _
   "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const VER_PLATFORM_WIN32_WINDOWS = 1

Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

'
'Key Board
Declare Function GetKeyboardState Lib "user32" (kbKey As Byte) As Long
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function SetKeyboardState Lib "user32" (kbKey As Byte) As Long

Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
   ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const VK_TAB = &H9
Public Const VK_INSERT = &H2D
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2

'Find Open Sections
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'ODBC 11/15/05
' Public Declare Function SQLSetConnectOption Lib "odbc32.dll" _
'        (ByVal hdbc&, ByVal fOption%, ByVal vParam As Any) As Integer
' Public Const SQL_PRESERVE_CURSORS As Long = 1204
' Public Const SQL_PC_ON As Long = 1

'Column Formating
'##0.000 or ##0.0000
Public ES_TimeFormat As String
'Purchasing
'######0.000 or #####0.0000
Public ES_PurchasedDataFormat As String
'Selling
'######0.000 or #####0.0000
Public ES_SellingPriceFormat As String
Public Const ES_QuantityDataFormat As String = "#####0.0000"
Public Const ES_MoneyFormat As String = "0.00"

'ES Color Formats
Public Const ES_RED = &HC0          'red
Public Const ES_WHITE = &HFFFFFF    'white
Public Const ES_BLUE = &H800000     'blue
Public Const ES_BLACK = &H0         'black
Public Const Es_FormBackColor = &H8000000F
Public Const Es_CheckBoxForeColor = &H8000000F
Public Const Es_HelpBackGroundColor = &H80000018
Public Const Es_TextBackColor = &HFFFFFF
Public Const Es_TextForeColor = &H0
Public Const Es_TextDisabled = &HE0E0E0
Public Const ES_ViewBackColor = &HE0FFFF '11/17/04 RGB(255, 253, 223) Views
Public Const ES_SystemBackcolor = vbButtonFace
'Form Constants
Public Const ES_RESIZE = 0
Public Const ES_DONTRESIZE = -1
Public Const ES_LIST = 0
Public Const ES_DONTLIST = -1
Public Const ES_IGNOREDASHES = 1 'Compress routine
'MsgBox
Public Const ES_NOQUESTION = &H124 'Question and return (Default NO)
Public Const ES_YESQUESTION = &H24 'Question and return (Default YES)

'About box

Public Type MEMORYSTATUS
   dwLength As Long
   dwMemoryLoad As Long
   dwTotalPhys As Long
   dwAvailPhys As Long
   dwTotalPageFile As Long
   dwAvailPageFile As Long
   dwTotalVirtual As Long
   dwAvailVirtual As Long
End Type

'Window status
Public Const Swp_NOSIZE = &H1
Public Const Swp_NOMOVE = &H2
Public Const hWnd_TopMost = -1
Public Const Hwnd_NOTOPMOST = -2
Public Const Flags = Swp_NOSIZE Or Swp_NOMOVE
