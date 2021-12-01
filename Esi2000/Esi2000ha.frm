VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form Esi2000ha 
   AutoRedraw      =   -1  'True
   Caption         =   "Fusion ERP"
   ClientHeight    =   585
   ClientLeft      =   4860
   ClientTop       =   765
   ClientWidth     =   7215
   HelpContextID   =   50
   Icon            =   "Esi2000ha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   585
   ScaleWidth      =   7215
   Begin VB.CommandButton cmdDataCol 
      Height          =   600
      Left            =   6000
      Picture         =   "Esi2000ha.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Data Collection"
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdPOM 
      Height          =   600
      Left            =   5400
      Picture         =   "Esi2000ha.frx":053D
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "POM"
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdTime 
      Height          =   600
      Left            =   3000
      Picture         =   "Esi2000ha.frx":0CAD
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Time Management"
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdInvc 
      Height          =   600
      Left            =   3600
      Picture         =   "Esi2000ha.frx":11F3
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Inventory"
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdSle 
      Height          =   600
      Left            =   1200
      Picture         =   "Esi2000ha.frx":17D4
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Sales"
      Top             =   0
      Width           =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   65000
      Left            =   2400
      Top             =   600
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1800
      Top             =   600
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   585
      FormDesignWidth =   7215
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Log Off"
      Height          =   600
      Left            =   6600
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Log Off Of ESI2000 - Right Click For Setup"
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdQual 
      Height          =   600
      Left            =   4200
      Picture         =   "Esi2000ha.frx":1CF5
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Quality Assurance"
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdFin 
      Height          =   600
      Left            =   4800
      Picture         =   "Esi2000ha.frx":228F
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Finance"
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdEng 
      Height          =   600
      Left            =   1800
      Picture         =   "Esi2000ha.frx":281D
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Engineering"
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdPrd 
      Height          =   600
      Left            =   2400
      Picture         =   "Esi2000ha.frx":2ECF
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Production"
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdAdm 
      Height          =   600
      Left            =   600
      Picture         =   "Esi2000ha.frx":343F
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Administration"
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Log On"
      Enabled         =   0   'False
      Height          =   600
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Log On To ESI2000 - Settings"
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "Esi2000ha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of          ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'9/26/06 See Form_Activate Block inadvertant TaskBar firing
Option Explicit
Dim bResized As Byte

Private Sub cmdAdm_Click()
   
   SectionButtonClick SECTION_ADMN, Me
End Sub

Private Sub cmdAdm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   ShowChoices Button, cmdAdm, SECTION_ADMN
   
End Sub

Private Sub cmdCan_Click()
   Dim sMsg As String
   Dim bResponse As Byte
   Timer1.Enabled = False
   sMsg = "Do You Want To Quit " & sSysCaption & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, sSysCaption)
   If bResponse = vbYes Then
      If Top > Screen.Height - Height Then Top = 0
      If Left > Screen.Width - Width Then Left = 0
      If Top >= 0 And Left >= 0 Then
         SaveSetting "Esi2000", "mngr", "hMngrTop", Top
         SaveSetting "Esi2000", "mngr", "hMngrLeft", Left
      End If
      bResponse = CloseManager()
   End If
   
End Sub

Private Sub cmdDataCol_Click()
   SectionButtonClick SECTION_DATACOL, Me
End Sub

Private Sub cmdEng_Click()
   
   SectionButtonClick SECTION_ENGR, Me
   
End Sub

Private Sub cmdEng_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   ShowChoices Button, cmdEng, SECTION_ENGR
   
End Sub


Private Sub cmdFin_Click()
   
   SectionButtonClick SECTION_FINA, Me
   
End Sub

Private Sub cmdFin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   ShowChoices Button, cmdFin, SECTION_FINA
   
End Sub


Private Sub cmdInvc_Click()
   
   SectionButtonClick SECTION_INVC, Me
   
End Sub

Private Sub cmdInvc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   ShowChoices Button, cmdInvc, SECTION_INVC
   
End Sub

Private Sub cmdOk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      SubjHelp.Timer1.Interval = 7000
      SubjHelp.optFrom = vbUnchecked
      SubjHelp.Section = "ESI Manager"
      SubjHelp.hlp = "To Change the Bar from the " & vbCr _
                     & "top to the side, double click " & vbCr _
                     & "the Menu Bar or stretch" & vbCr _
                     & "the height to the top or to the Bottom."
      On Error Resume Next
      SubjHelp.Left = Esi2000ha.Left + 100
      SubjHelp.Show
      SubjHelp.Height = 1400
   End If
   
End Sub


Private Sub cmdPOM_Click()
'   Dim strPOM As String
'   strPOM = "EsiPOM.exe"
'   Shell sFilePath & strPOM & " " & Command, vbMaximizedFocus
   SectionButtonClick SECTION_POM, Me
   
End Sub

Private Sub cmdPOM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Dim strPOM As String
'   strPOM = "EsiPOM.exe"
'   Shell sFilePath & strPOM & " " & Command, vbMaximizedFocus
   ShowChoices Button, cmdPOM, SECTION_POM

End Sub

Private Sub cmdPrd_Click()
   
   SectionButtonClick SECTION_PROD, Me
   
End Sub



Private Sub cmdPrd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   ShowChoices Button, cmdPrd, SECTION_PROD
End Sub


Private Sub cmdQual_Click()
   
   SectionButtonClick SECTION_QUAL, Me
End Sub

Private Sub cmdQual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   ShowChoices Button, cmdQual, SECTION_QUAL
   
End Sub


Private Sub cmdSle_Click()
   
   SectionButtonClick SECTION_SALE, Me
End Sub

Private Sub cmdSle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   ShowChoices Button, cmdSle, SECTION_SALE
End Sub

Private Sub cmdTime_Click()
   
   SectionButtonClick SECTION_TIME, Me
   
End Sub

Private Sub cmdTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   ShowChoices Button, cmdTime, SECTION_TIME
   
End Sub

Private Sub Form_Activate()
   'If bWeAreLoaded = 1 Then WindowState = vbNormal
   bVerticalLoaded = 0
   bShowVertical = GetSetting("Esi2000", "mngr", "ShowVertical", bShowVertical)
   If bShowVertical = 1 Then
      bVerticalLoaded = 1
      Hide
      Esi2000v.Show
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Initialize()
   Dim sMin$
   Dim sFrmTop$
   Dim sFrmLeft$
   
   Dim wFlags As Long
   Dim a As Long
   Dim hMenu As Long
   wFlags = 1
   
   hMenu = GetSystemMenu(hwnd, 0)
   SetWindowPos hwnd, hWnd_TopMost, 0, 0, 0, 0, wFlags
   a& = ModifyMenu&(hMenu, SC_CLOSE, MF_GRAYED, -10, "Close")
   Top = GetSetting("Esi2000", "mngr", "hMngrTop", Top)
   If Top < 0 Then Top = 100
   If Top > Screen.Height Then Top = 100
   
   Left = GetSetting("Esi2000", "mngr", "hMngrLeft", Left)
   If Left < 0 Then Left = 100
   If Left > (Screen.Width - Width) Then Left = Screen.Width - Width
   bVerticalLoaded = 0
   bResized = 0
   'If sServer = "JEVCO2" Then Timer1.Enabled = True _
Else Timer1.Enabled = False
   If bUserLoggedOn = 0 Then EsiLogon.Show
   
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' If the user hits CTRL-R then show the registration form
    If Shift = 2 And KeyCode = 82 Then
        ESIRegister.Show
    End If
End Sub

Private Sub Form_Load()
   Dim i%
   Caption = GetSystemCaption()
   For i% = 0 To 6
      SectionInfo(SECTION_INVC).SectionOpen = 0
   Next
   GetAppTitles
   If bWeAreLoaded = 1 Then
      WindowState = vbNormal
      Show
   End If
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim b As Byte
   Timer1.Enabled = False
   If UnloadMode = vbAppWindows Then
      MsgBox "Please Close " & sSysCaption & " Before Exiting Windows.", vbExclamation, sSysCaption
      Cancel = True
      Exit Sub
   End If
   If bVerticalLoaded = 0 Then
      bShowVertical = 0
   End If
   
   If CloseManager() = 1 Then
      Cancel = True
   End If
   
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   If WindowState = vbNormal Then
      If Top + Height > Screen.Height Then Top = Screen.Height - Height
      If Top < 0 Then Top = 100
      If Left < 0 Then Left = 100
      If Left > (Screen.Width - Width) Then Left = Screen.Width - Width
      If Height > 1600 Or Width < 6000 Then
         SaveSetting "Esi2000", "mngr", "ShowVertical", 1
         Hide
         bVerticalLoaded = 1
         Esi2000v.Show
      End If
      
      'eight = 930
      'idth = 4545
      SizeMe
   End If
   If WindowState = vbMaximized Then
      WindowState = vbNormal
      Hide
      bVerticalLoaded = 1
      Esi2000v.Show
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' Set Esi2000ha = Nothing
   
End Sub







'Changed 12000 2/7/02

Private Sub SizeMe()
   
   Height = 1095
   Width = 6105
   '    Dim iScreen As Integer
   '    iScreen = Screen.Width \ Screen.TwipsPerPixelX
   '    Select Case iScreen
   '        Case 1024
   '            Height = 900
   '            Width = 5238
   '        Case 1280
   '            Height = 1048
   '            Width = 6318
   '         Case Is > 1590
   '            Height = 1125
   '            Width = 7245
   '        Case Else
   '            Height = 930
   '            Width = 5120
   '    End Select
   
End Sub

Private Sub Timer1_Timer()
   Static b As Byte
   b = b + 1
   If b = 2 Then
      GetAppTitles
      b = 0
   End If
   GetSystemMessage
   
End Sub
