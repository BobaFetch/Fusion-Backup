VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form Esi2000v 
   AutoRedraw      =   -1  'True
   Caption         =   "Fusion ERP"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   2145
   ClientWidth     =   600
   ControlBox      =   0   'False
   HelpContextID   =   50
   Icon            =   "Esi2000v.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6795
   ScaleWidth      =   600
   Begin VB.CommandButton cmdPOM 
      Height          =   600
      Left            =   0
      Picture         =   "Esi2000v.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "POM Module"
      Top             =   5595
      Width           =   600
   End
   Begin VB.CommandButton cmdTime 
      Height          =   600
      Left            =   0
      Picture         =   "Esi2000v.frx":103A
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Time Management"
      Top             =   3195
      Width           =   600
   End
   Begin VB.CommandButton cmdInvc 
      Height          =   600
      Left            =   0
      Picture         =   "Esi2000v.frx":1580
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Inventory Control"
      Top             =   3795
      Width           =   600
   End
   Begin VB.CommandButton cmdSle 
      Height          =   600
      Left            =   0
      Picture         =   "Esi2000v.frx":1B61
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Sales"
      Top             =   1395
      Width           =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   65000
      Left            =   900
      Top             =   3360
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   960
      Top             =   1500
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6795
      FormDesignWidth =   600
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "---"
      Height          =   195
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Log Off"
      Height          =   600
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Log Off Of ESI2000 - Right Click For Setup"
      Top             =   6195
      Width           =   600
   End
   Begin VB.CommandButton cmdQual 
      Height          =   600
      Left            =   0
      Picture         =   "Esi2000v.frx":2082
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Quality Assurance"
      Top             =   4395
      Width           =   600
   End
   Begin VB.CommandButton cmdFin 
      Height          =   600
      Left            =   0
      Picture         =   "Esi2000v.frx":261C
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Financial Accounting"
      Top             =   4995
      Width           =   600
   End
   Begin VB.CommandButton cmdEng 
      Height          =   600
      Left            =   0
      Picture         =   "Esi2000v.frx":2BAA
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Engineering"
      Top             =   1995
      Width           =   600
   End
   Begin VB.CommandButton cmdPrd 
      Height          =   600
      Left            =   0
      Picture         =   "Esi2000v.frx":325C
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Production"
      Top             =   2595
      Width           =   600
   End
   Begin VB.CommandButton cmdAdm 
      Height          =   600
      Left            =   0
      Picture         =   "Esi2000v.frx":37CC
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Administration"
      Top             =   795
      Width           =   600
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Log On"
      Enabled         =   0   'False
      Height          =   600
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Log On To ESI2000 - Settings"
      Top             =   195
      Width           =   600
   End
End
Attribute VB_Name = "Esi2000v"
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

Private Sub SizeMe()
   Height = 6705
   Width = 720
End Sub

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
         SaveSetting "Esi2000", "mngr", "vMngrTop", Top
         SaveSetting "Esi2000", "mngr", "vMngrLeft", Left
      End If
      bResponse = CloseManager()
   End If
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


Private Sub cmdMin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Then WindowState = vbMinimized
End Sub

Private Sub cmdMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   WindowState = vbMinimized
End Sub

Private Sub cmdOk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      SubjHelp.Timer1.Interval = 7000
      SubjHelp.optFrom = vbChecked
      SubjHelp.Section = "ESI Manager"
      SubjHelp.hlp = "To Change the Bar from the " & vbCr _
                     & "side to the top, double click " & vbCr _
                     & "the Menu Bar or stretch" & vbCr _
                     & "the width to the right or to the left."
      If (Left + Width + 2175) > (Screen.Width) Then
         SubjHelp.Left = Esi2000v.Left - 2475
      Else
         SubjHelp.Left = Esi2000v.Left + 700
      End If
      On Error Resume Next
      SubjHelp.Top = Esi2000v.Top
      SubjHelp.Show
      SubjHelp.Height = 1400
   End If
   
End Sub

Private Sub cmdPOM_Click()
   SectionButtonClick SECTION_POM, Me
End Sub

Private Sub cmdPOM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
   bVerticalLoaded = 1
   bShowVertical = GetSetting("Esi2000", "mngr", "ShowVertical", bShowVertical)
   If bShowVertical = 0 Then
      bVerticalLoaded = 0
      Hide
      Esi2000h.Show
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
   
   iMinimize = GetSetting("Esi2000", "mngr", "MinOnOpen", iMinimize)
   Top = GetSetting("Esi2000", "mngr", "vMngrTop", Top)
   If Top < 0 Then Top = 100
   If Top > Screen.Height Then Top = 100
   Left = GetSetting("Esi2000", "mngr", "vMngrLeft", Left)
   If Left < 0 Then Left = 100
   If Left > (Screen.Width - Width) Then Left = Screen.Width - Width
   bVerticalLoaded = 1
   If bUserLoggedOn = 0 Then EsiLogon.Show
   
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' If the user hits CTRL-R then show the registration form
    If Shift = 2 And KeyCode = 82 Then
        ESIRegister.Show
    End If
End Sub

Private Sub Form_Load()
   Dim I%
   Caption = GetSystemCaption()
   For I% = 0 To 6
      SectionInfo(SECTION_INVC).SectionOpen = 0
   Next
   GetAppTitles
   If bWeAreLoaded = 1 Then
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
   
   If bVerticalLoaded = 1 Then
      bShowVertical = 1
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
      If Width > 1600 Or Height < 6200 Then
         SaveSetting "Esi2000", "mngr", "ShowVertical", 0
         Hide
         bVerticalLoaded = 0
         Esi2000h.Show
      End If
      SizeMe
   Else
      If WindowState = vbMaximized Then
         WindowState = vbNormal
         Hide
         bVerticalLoaded = 0
         Esi2000h.Show
      End If
   End If
   
End Sub

'9/17/04 Jevco patch for Inventory Update

Private Sub Timer1_Timer()
   Static b As Byte
   b = b + 1
   If b = 2 Then
      GetAppTitles
      b = 0
   End If
   GetSystemMessage
   
End Sub
