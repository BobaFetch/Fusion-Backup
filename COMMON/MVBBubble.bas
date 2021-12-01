Attribute VB_Name = "MVBBubble"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Name:         MVBBubble
'' Filename:     MVBBubble.bas
'' Author:       Mattias Sjögren (mattiass@hem.passagen.se)
''               http://www.msjogren.net/dotnet/
''
'' Description:  Adds multiline and text-alignment support to VB tooltips.
''
''               To center the tip contents, begin the tip text with "<c>".
''               To right align it, begin the text with "<r>".
''
''               Use the MaxTipWidth property to limit the tip width, and the
''               HideToolTips proeprty to prevent any tooltips from showing.
''
'' Dependencies: None
''
''
'' Copyright ©2000-2001, Mattias Sjögren
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit


'''''''''''''''''''''
'''   Constants   '''
'''''''''''''''''''''

' CBT hook constants
Private Const WH_CBT = 5
Private Const HCBT_CREATEWND = 3
Private Const HCBT_DESTROYWND = 4

' Get/SetWindowLong constants
Private Const GWL_WNDPROC = (-4)

' Window messages
Private Const WM_DESTROY = &H2
Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_PAINT = &HF
Private Const WM_WINDOWPOSCHANGING = &H46
Private Const WM_SETTINGCHANGE = &H1A

' SetWindowPos flags
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80

' DrawText flags
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_WORDBREAK = &H10
Private Const DT_EXPANDTABS = &H40
Private Const DT_CALCRECT = &H400

' System color constants
Private Const COLOR_INFOTEXT = 23
Private Const COLOR_INFOBK = 24

' SetBkMode background modes
Private Const TRANSPARENT = 1

' SystemParametersInfo constants
Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const SPI_SETNONCLIENTMETRICS = 42

' GetSystemMetrics constants
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6


' Minimum space between tooltip border and text
Private Const HORIZ_MARGIN = 1
Private Const VERT_MARGIN = 1


'''''''''''''''''
'''   Types   '''
'''''''''''''''''

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Type SIZE
  cx As Long
  cy As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type WINDOWPOS
  hwnd As Long
  hWndInsertAfter As Long
  x As Long
  y As Long
  cx As Long
  cy As Long
  flags As Long
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(1 To 32) As Byte   ' LF_FACESIZE = 32
End Type

Private Type NONCLIENTMETRICS
  cbSize As Long
  iBorderWidth As Long
  iScrollWidth As Long
  iScrollHeight As Long
  iCaptionWidth As Long
  iCaptionHeight As Long
  lfCaptionFont As LOGFONT
  iSMCaptionWidth As Long
  iSMCaptionHeight As Long
  lfSMCaptionFont As LOGFONT
  iMenuWidth As Long
  iMenuHeight As Long
  lfMenuFont As LOGFONT
  lfStatusFont As LOGFONT
  lfMessageFont As LOGFONT
End Type


''''''''''''''''''''
'''   Declares   '''
''''''''''''''''''''

Private Declare Function EnumThreadWindows Lib "user32" ( _
  ByVal dwThreadId As Long, _
  ByVal lpfn As Long, _
  ByVal lParam As Long) As Long

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
  ByVal idHook As Long, _
  ByVal lpfn As Long, _
  ByVal hMod As Long, _
  ByVal dwThreadId As Long) As Long
  
Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
  ByVal hhk As Long) As Long
  
Private Declare Function CallNextHookEx Lib "user32" ( _
  ByVal hhk As Long, _
  ByVal nCode As Long, _
  ByVal wParam As Long, _
  lParam As Any) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
  ByVal hwnd As Long, _
  ByVal nIndex As Long, _
  ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
  ByVal lpPrevWndFunc As Long, _
  ByVal hwnd As Long, _
  ByVal Msg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
  ByVal hwnd As Long, _
  ByVal Msg As Long, _
  ByVal wParam As Long, _
  lParam As Any) As Long

Private Declare Function SetWindowPos Lib "user32" ( _
  ByVal hwnd As Long, _
  ByVal hWndInsertAfter As Long, _
  ByVal x As Long, _
  ByVal y As Long, _
  ByVal cx As Long, _
  ByVal cy As Long, _
  ByVal uFlags As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
  ByVal hwnd As Long, _
  ByVal lpClassName As String, _
  ByVal nMaxCount As Long) As Long

Private Declare Function GetClientRect Lib "user32" ( _
  ByVal hwnd As Long, _
  lpRect As RECT) As Long

Private Declare Function InflateRect Lib "user32" ( _
  lprc As RECT, _
  ByVal dx As Long, _
  ByVal dy As Long) As Long

Private Declare Function FillRect Lib "user32" ( _
  ByVal hdc As Long, _
  lprc As RECT, _
  ByVal hbr As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" ( _
  ByVal hdc As Long, _
  ByVal lpString As String, _
  ByVal nCount As Long, _
  lpRect As RECT, _
  ByVal uFormat As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" ( _
  ByVal hdc As Long, _
  ByVal iBkMode As Long) As Long

Private Declare Function GetDC Lib "user32" ( _
  ByVal hwnd As Long) As Long
  
Private Declare Function ReleaseDC Lib "user32" ( _
  ByVal hwnd As Long, _
  ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" ( _
  ByVal hdc As Long, _
  ByVal hgdiobj As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" ( _
  ByVal hObject As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" ( _
  lplf As LOGFONT) As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
  ByVal uiAction As Long, _
  ByVal uiParam As Long, _
  pvParam As Any, _
  ByVal fWinIni As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" ( _
  ByVal nIndex As Long) As Long

Private Declare Function GetCursorPos Lib "user32" ( _
  lpPoint As POINTAPI) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
  pDest As Any, _
  pSource As Any, _
  ByVal cb As Long)


'''''''''''''''''''''
'''   Variables   '''
'''''''''''''''''''''

Private m_hwndTT As Long
Private m_hCBTHook As Long
Private m_pfnOldWndProc As Long
Private m_hFont As Long

Private m_fHideTips As Boolean
Private m_nMaxWidth As Long


''''''''''''''''''''''''''
'''   Public methods   '''
''''''''''''''''''''''''''

'
' Function HookToolTips
'
' Description:  Call this function to initialize to the module.
'
' Returns:      True on success, False on failure.
'
Public Function HookToolTips() As Boolean
    
  If (m_hCBTHook <> 0) Or (m_pfnOldWndProc <> 0) Then Exit Function
    
  ' First enumerate current windows to see it the tooltip is already created
  Call EnumThreadWindows(App.ThreadID, AddressOf EnumThreadWndProc, 0&)
  
  If m_hwndTT Then
    ' Found an existing tooltip window - subclass it
    m_pfnOldWndProc = SetWindowLong(m_hwndTT, GWL_WNDPROC, AddressOf ToolTipWndProc)
    CreateFont
    HookToolTips = True
  Else
    ' No tooltip window created yet. Set up hook so we get notified when it is
    m_hCBTHook = SetWindowsHookEx(WH_CBT, AddressOf CBTProc, 0&, App.ThreadID)
    HookToolTips = CBool(m_hCBTHook)
  End If
  
  ' Start with a huge max width
  m_nMaxWidth = &H7FFFFFFF
  
End Function

'
' Sub UnhookToolTips
'
' Description:  This procedure should be called before the application terminates
'               to stop the tooltip window subclassing. If the subclass isn't removed
'               when the tooltip window is destroyed, the application is likely to crash.
'
Public Sub UnhookToolTips()
  
  If m_hCBTHook Then Call UnhookWindowsHookEx(m_hCBTHook)
  If m_pfnOldWndProc Then Call SetWindowLong(m_hwndTT, GWL_WNDPROC, m_pfnOldWndProc)
  If m_hFont Then Call DeleteObject(m_hFont)
  m_hCBTHook = 0
  m_pfnOldWndProc = 0
  m_hwndTT = 0
  m_hFont = 0

End Sub

'
' Property HideToolTips
'
' Description:  Used to temporarily disable all tooltips from appearing,
'               without having to clear the ToolTipText property for every
'               control.
'
Public Property Get HideToolTips() As Boolean
  HideToolTips = m_fHideTips
End Property

Public Property Let HideToolTips(ByVal NewHideToolTips As Boolean)
  m_fHideTips = NewHideToolTips
End Property

'
' Property MaxTipWidth
'
' Description:  Retrieves or sets the maximum allowed width of the tooltip
'               window. If the tip text is wider than this, it will be wrapped
'               to the next line.
'
Public Property Get MaxTipWidth() As Long
  MaxTipWidth = m_nMaxWidth
End Property

Public Property Let MaxTipWidth(ByVal NewMaxTipWidth As Long)
  m_nMaxWidth = NewMaxTipWidth
End Property

'
' Function EnumThreadWndProc
'
' Description:  EnumThreadWindows callback function.
'
Public Function EnumThreadWndProc(ByVal hwnd As Long, ByVal lParam As Long) As Long

  EnumThreadWndProc = 1   ' default is to continue enumaration - return TRUE ...
  If IsBubbleWindow(hwnd) Then EnumThreadWndProc = 0    ' ... but stop if we find the tooltip
  
End Function

'
' Function CBTProc
'
' Description:  Hook procedure used to catch the creation of the tooltip window.
'
Public Function CBTProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
  If nCode = HCBT_CREATEWND Then
    If IsBubbleWindow(wParam) Then
      
      ' Subclass the tooltip window
      m_pfnOldWndProc = SetWindowLong(wParam, GWL_WNDPROC, AddressOf ToolTipWndProc)
      
      CreateFont
      
      ' Now that the one and only VBBubble tooltip window is created,
      ' we don't need the hook anymore.
      Call UnhookWindowsHookEx(m_hCBTHook)
      m_hCBTHook = 0
      
    End If  ' IsBubbleWindow()
  End If  ' nCode = HCBT_CREATEWND
  
  CBTProc = CallNextHookEx(m_hCBTHook, nCode, wParam, ByVal lParam)
  
End Function

'
' Function ToolTipWndProc
'
' Description:  Window procedure of the subclassed tooltip window.
'
Public Function ToolTipWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Static wp As WINDOWPOS
  Static ptCursor As POINTAPI
  
  
  Select Case uMsg
    Case WM_SETTINGCHANGE

      If wParam = SPI_SETNONCLIENTMETRICS Then Call CreateFont

    ' End WM_SETTINGCHANGE
    
    
    Case WM_PAINT

      ' Let VB handle the message first
      Call CallWindowProc(m_pfnOldWndProc, hwnd, uMsg, wParam, lParam)

      ' Do our own drawing
      Call DrawToolTip

      ToolTipWndProc = 0
      Exit Function

    ' End WM_PAINT
    
    
    Case WM_WINDOWPOSCHANGING

      ' Let VB handle the message first
      Call CallWindowProc(m_pfnOldWndProc, hwnd, uMsg, wParam, lParam)

      Call CopyMemory(wp, ByVal lParam, Len(wp))

      If m_fHideTips Then
        ' Hide all tooltips
        wp.flags = SWP_HIDEWINDOW
      Else
        GetToolTipSize wp.cx, wp.cy

        ' Center tip under cursor
        Call GetCursorPos(ptCursor)
        wp.x = ptCursor.x - wp.cx \ 2
        If wp.x < 0 Then wp.x = 0
        If (wp.x + wp.cx) > (Screen.Width \ Screen.TwipsPerPixelX) Then _
            wp.x = (Screen.Width \ Screen.TwipsPerPixelX) - wp.cx

      End If

      CopyMemory ByVal lParam, wp, Len(wp)

      ToolTipWndProc = 0
      Exit Function
      
    ' End WM_WINDOWPOSCHANGING
    
  End Select
  
  ToolTipWndProc = CallWindowProc(m_pfnOldWndProc, hwnd, uMsg, wParam, lParam)
        
End Function


''''''''''''''''''''''''''
'''   Private methods   '''
''''''''''''''''''''''''''

'
' Function IsBubbleWindow
'
' Description:  Checks if the window is a VB tooltip.
'
Private Function IsBubbleWindow(ByVal hwnd As Long) As Boolean

  Static sClass As String
  
  
  IsBubbleWindow = False
  
  sClass = String$(20, 0)
  Call GetClassName(hwnd, sClass, 20)
    
  ' Tooltip class names
  '  In VB5/6 IDE: "VBBubble"
  '  In VB5 runtime: "VBBubbleRT5"
  '  In VB6 runtime: "VBBubbleRT6"
          
  If StrComp(Left$(sClass, 8), "VBBubble", vbTextCompare) = 0 Then
    IsBubbleWindow = True
    m_hwndTT = hwnd
  End If

End Function

'
' Sub CreateFont
'
' Description:  (Re)creates the font used to draw the tooltip text,
'               based on the system tooltip font settings.
'
Private Sub CreateFont()
  
  Dim ncm As NONCLIENTMETRICS
  
  
  If m_hFont Then Call DeleteObject(m_hFont)
  
  ncm.cbSize = Len(ncm)
  Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, Len(ncm), ncm, 0&)
  m_hFont = CreateFontIndirect(ncm.lfStatusFont)
  
End Sub

'
' Sub DrawToolTip
'
' Description:  Draws the tooltip.
'
Private Sub DrawToolTip()

  Dim rc As RECT
  Dim sTipText As String, sAlignTag As String
  Dim nTipTextLen As Long
  Dim hdcTT As Long
  Dim nOldBkMode As Long
  Dim hOldFont As Long
  Dim nDTFlags As Long
  
  
  Call GetClientRect(m_hwndTT, rc)

  ' Get tip text
  nTipTextLen = SendMessage(m_hwndTT, WM_GETTEXTLENGTH, 0&, ByVal 0&)
  sTipText = String$(nTipTextLen, 0)
  Call SendMessage(m_hwndTT, WM_GETTEXT, nTipTextLen + 1, ByVal sTipText)

  ' Default is left aligned text
  nDTFlags = DT_LEFT
  
  ' Check if an alignment tag is available
  sAlignTag = LCase$(Left$(sTipText, 3))
  Select Case sAlignTag
    Case "<c>"
      nDTFlags = DT_CENTER
      sTipText = Mid$(sTipText, 4)

    Case "<r>"
      nDTFlags = DT_RIGHT
      sTipText = Mid$(sTipText, 4)
  End Select
  
  hdcTT = GetDC(m_hwndTT)
  
  nOldBkMode = SetBkMode(hdcTT, TRANSPARENT)
  hOldFont = SelectObject(hdcTT, m_hFont)

  ' Fill entire tip window with the correct color to erase the drawing made by VB
  Call FillRect(hdcTT, rc, COLOR_INFOBK + 1)
  
  Call InflateRect(rc, -HORIZ_MARGIN, -VERT_MARGIN)
  Call DrawText(hdcTT, sTipText, -1&, rc, nDTFlags Or DT_WORDBREAK Or DT_EXPANDTABS)

  Call SelectObject(hdcTT, hOldFont)
  Call SetBkMode(hdcTT, nOldBkMode)
  Call ReleaseDC(m_hwndTT, hdcTT)
  
End Sub

'
' Sub GetToolTipSize
'
' Description:  Calculates the size of the tooltip window so that it's large enough
'               to display the current tip text.
'
' x/y:          [out] Recieves width and height of the window.
'
Private Sub GetToolTipSize(ByRef x As Long, ByRef y As Long)

  Dim rc As RECT
  Dim sTipText As String, sAlignTag As String
  Dim nTipTextLen As Long
  Dim hdcTT As Long
  Dim hOldFont As Long
  
  
  ' Get tip text
  nTipTextLen = SendMessage(m_hwndTT, WM_GETTEXTLENGTH, 0&, ByVal 0&)
  sTipText = String$(nTipTextLen, 0)
  Call SendMessage(m_hwndTT, WM_GETTEXT, nTipTextLen + 1, ByVal sTipText)
    
  ' Remove alignment tag
  sAlignTag = LCase$(Left$(sTipText, 3))
  If sAlignTag = "<c>" Or sAlignTag = "<r>" Then sTipText = Mid$(sTipText, 4)
  
  hdcTT = GetDC(m_hwndTT)
  hOldFont = SelectObject(hdcTT, m_hFont)
  
  rc.Right = m_nMaxWidth
  Call DrawText(hdcTT, sTipText, -1&, rc, DT_CALCRECT Or DT_WORDBREAK Or DT_EXPANDTABS)
  
  Call SelectObject(hdcTT, hOldFont)
  Call ReleaseDC(m_hwndTT, hdcTT)
  
  x = (rc.Right - rc.Left) + 2 * (GetSystemMetrics(SM_CXBORDER) + HORIZ_MARGIN)
  y = (rc.Bottom - rc.Top) + 2 * (GetSystemMetrics(SM_CYBORDER) + VERT_MARGIN)
    
End Sub
