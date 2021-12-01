VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDISect 
   BackColor       =   &H8000000C&
   Caption         =   "Administration"
   ClientHeight    =   9630
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   13050
   HelpContextID   =   1000
   Icon            =   "MDIAdmn.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox BotPanel 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   340
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   13050
      TabIndex        =   22
      Top             =   9285
      Width           =   13050
      Begin VB.Label OvrPanel 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OVER"
         Height          =   324
         Left            =   9384
         TabIndex        =   26
         Top             =   24
         Width           =   780
      End
      Begin VB.Label tmePanel 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   324
         Left            =   10224
         TabIndex        =   25
         Top             =   24
         Width           =   720
      End
      Begin VB.Label lblBotPanel 
         Height          =   336
         Left            =   100
         TabIndex        =   24
         Top             =   24
         Width           =   2532
      End
      Begin VB.Label SystemMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   2664
         TabIndex        =   23
         Top             =   24
         Width           =   6300
      End
   End
   Begin MSComctlLib.StatusBar RightPanel 
      Align           =   4  'Align Right
      Height          =   8685
      Left            =   12840
      TabIndex        =   21
      Top             =   600
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   15319
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2
            MinWidth        =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox LeftBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   8685
      Left            =   1980
      ScaleHeight     =   8685
      ScaleWidth      =   195
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.PictureBox TopBar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   13050
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   13050
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MDIAdmn.frx":014A
         Height          =   828
         Index           =   13
         Left            =   5580
         Picture         =   "MDIAdmn.frx":0CB0
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Inventory Management"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MDIAdmn.frx":6AB2
         Height          =   588
         Index           =   11
         Left            =   4800
         Picture         =   "MDIAdmn.frx":76F4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "System Help - All Sections"
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MDIAdmn.frx":8336
         Height          =   588
         Index           =   9
         Left            =   3960
         Picture         =   "MDIAdmn.frx":8E9C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Inventory Management"
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MDIAdmn.frx":9A02
         Height          =   588
         Index           =   7
         Left            =   3120
         Picture         =   "MDIAdmn.frx":A824
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Time Management"
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MDIAdmn.frx":B646
         Height          =   588
         Index           =   5
         Left            =   2280
         Picture         =   "MDIAdmn.frx":B950
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Production Control Administration"
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MDIAdmn.frx":BC5A
         Height          =   588
         Index           =   3
         Left            =   1440
         Picture         =   "MDIAdmn.frx":BF64
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Sales Administration"
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MDIAdmn.frx":C26E
         Height          =   588
         Index           =   2
         Left            =   600
         Picture         =   "MDIAdmn.frx":CB38
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "System Administration"
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.Image Logo 
         Height          =   765
         Index           =   1
         Left            =   10080
         Picture         =   "MDIAdmn.frx":D402
         Top             =   0
         Width           =   2430
      End
   End
   Begin VB.PictureBox SideBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   8685
      Left            =   0
      ScaleHeight     =   8685
      ScaleWidth      =   1980
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1980
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MDIAdmn.frx":1357C
         Height          =   828
         Index           =   12
         Left            =   120
         Picture         =   "MDIAdmn.frx":13D2B
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Inventory Management"
         Top             =   4280
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MDIAdmn.frx":144DA
         Height          =   828
         Index           =   10
         Left            =   1020
         Picture         =   "MDIAdmn.frx":14DDD
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "System Help - All Sections"
         Top             =   3000
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MDIAdmn.frx":156E0
         Height          =   828
         Index           =   8
         Left            =   120
         Picture         =   "MDIAdmn.frx":16050
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Inventory Management"
         Top             =   3000
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MDIAdmn.frx":169C0
         Height          =   828
         Index           =   6
         Left            =   1020
         Picture         =   "MDIAdmn.frx":171E2
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Time Management"
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MDIAdmn.frx":17A04
         Height          =   828
         Index           =   4
         Left            =   120
         Picture         =   "MDIAdmn.frx":181AE
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Production Control Administration"
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MDIAdmn.frx":18958
         Height          =   828
         Index           =   1
         Left            =   1020
         Picture         =   "MDIAdmn.frx":1916A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Sales Administration"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MDIAdmn.frx":1997C
         Height          =   828
         Index           =   0
         Left            =   120
         Picture         =   "MDIAdmn.frx":1A2CD
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "System Administration"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin ResizeLibCtl.ReSize ReSize1 
         Left            =   1260
         Top             =   4320
         _Version        =   196615
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
         Enabled         =   0   'False
         FormMinWidth    =   0
         FormMinHeight   =   0
         FormDesignHeight=   9630
         FormDesignWidth =   13050
      End
      Begin Crystal.CrystalReport Crw 
         Left            =   1260
         Top             =   4800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowLeft      =   110
         WindowTop       =   35
         WindowWidth     =   460
         WindowHeight    =   410
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         WindowControls  =   -1  'True
         WindowState     =   1
         PrintFileLinesPerPage=   60
      End
      Begin VB.Image imgPartList 
         Height          =   300
         Left            =   600
         Picture         =   "MDIAdmn.frx":1AC1E
         Top             =   7440
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgNewPart 
         Height          =   300
         Left            =   120
         Picture         =   "MDIAdmn.frx":1B0BC
         Top             =   7320
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgPartFind 
         Height          =   300
         Left            =   1440
         Picture         =   "MDIAdmn.frx":1B557
         Top             =   7080
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XPrinter_small 
         Height          =   300
         Left            =   480
         Picture         =   "MDIAdmn.frx":1B991
         Stretch         =   -1  'True
         Top             =   7080
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgStandardComment 
         Height          =   300
         Left            =   960
         Picture         =   "MDIAdmn.frx":1BDD1
         Top             =   7080
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Database Maint"
         Height          =   495
         Index           =   6
         Left            =   60
         TabIndex        =   28
         ToolTipText     =   "Inventory Management"
         Top             =   5100
         Width           =   1005
      End
      Begin VB.Image XDisplay 
         Height          =   300
         Left            =   1140
         Picture         =   "MDIAdmn.frx":1C229
         Top             =   6060
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XPrinter 
         Height          =   300
         Left            =   1500
         Picture         =   "MDIAdmn.frx":1C646
         Top             =   6060
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sales"
         Height          =   492
         Index           =   1
         Left            =   1020
         TabIndex        =   6
         ToolTipText     =   "Sales Administration"
         Top             =   1200
         Width           =   860
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "System"
         Height          =   492
         Index           =   0
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "System Settings And Administration"
         Top             =   1200
         Width           =   860
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Production Control"
         Height          =   492
         Index           =   2
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Production Control Management"
         Top             =   2520
         Width           =   860
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time Mgmt"
         Height          =   492
         Index           =   3
         Left            =   1020
         TabIndex        =   3
         ToolTipText     =   "Time Cards And Employees"
         Top             =   2520
         Width           =   860
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory Mgmt"
         Height          =   492
         Index           =   4
         Left            =   80
         TabIndex        =   2
         ToolTipText     =   "Inventory Management"
         Top             =   3840
         Width           =   1000
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "System Help"
         Height          =   492
         Index           =   5
         Left            =   1020
         TabIndex        =   1
         ToolTipText     =   "System Help Files"
         Top             =   3840
         Width           =   860
      End
      Begin VB.Image XPHelpDn 
         Height          =   300
         Left            =   1500
         Picture         =   "MDIAdmn.frx":1CA91
         Top             =   5700
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image XPHelpUp 
         Height          =   300
         Left            =   1140
         Picture         =   "MDIAdmn.frx":1CFDB
         Top             =   5700
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image XPPrinterDn 
         Height          =   300
         Left            =   1500
         Picture         =   "MDIAdmn.frx":1D525
         Top             =   5340
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XPPrinterUp 
         Height          =   300
         Left            =   1140
         Picture         =   "MDIAdmn.frx":1D970
         Top             =   5340
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image Logo 
         Height          =   1500
         Index           =   0
         Left            =   120
         Picture         =   "MDIAdmn.frx":1DDBB
         Top             =   5520
         Width           =   1935
      End
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   3960
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   2520
      Top             =   1800
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2040
      Top             =   1800
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3000
      Top             =   1800
   End
   Begin MSComDlg.CommonDialog Cdi 
      Left            =   1080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   2.33096e-38
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   5040
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "MDIAdmn.frx":21662
   End
End
Attribute VB_Name = "MDISect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'7/22/04 Demo Help Loaded in Help Contents
Public bUnloading As Byte
Dim bOnLoad As Byte

Public Sub CheckSectionPermissions()
   'Dim bByte As Byte
   
   Dim iList As Integer
   Dim iHideModBtn As Integer
   On Error Resume Next
   
   'let the programmer see everything
   If RunningInIDE Then
      InitializePermissions Secure, 1
   End If
   If bSecSet = 1 Then
      
      ' Check flag to Hide module buttons if user don't have permission
      iHideModBtn = GetHideModule()
      
      If Secure.UserAdmn <> 1 Then
         For iList = 0 To 11
            cmdSect(iList).Enabled = False
         Next
         For iList = 0 To 4
            MDISect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(iList))).Visible = False
         Next
         For iList = 1 To 6
            MDISect.ActiveBar1.Bands("mnuWindow").Tools("Window" & Trim(str(iList))).Visible = False
         Next
         For iList = 1 To 12
            MDISect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(str(iList))).Visible = False
         Next
         MDISect.SystemMsg = "There Are No Section User Permissions"
      Else
         If Secure.UserAdmnG1 <> 1 Then
         
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(0).Visible = False
                cmdSect(2).Visible = False
                z1(0).Visible = False
            Else
                cmdSect(0).Enabled = False
                cmdSect(2).Enabled = False
                MDISect.ActiveBar1.Bands("mnuWindow").Tools("Window1").Visible = False
            End If
         End If
         
         If Secure.UserAdmnG2 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(1).Visible = False
                cmdSect(3).Visible = False
                z1(1).Visible = False
            Else
                cmdSect(1).Enabled = False
                cmdSect(3).Enabled = False
                MDISect.ActiveBar1.Bands("mnuWindow").Tools("Window2").Visible = False
            End If
         End If
         
         If Secure.UserAdmnG3 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(4).Visible = False
                cmdSect(5).Visible = False
                z1(2).Visible = False
            Else
                cmdSect(4).Enabled = False
                cmdSect(5).Enabled = False
                MDISect.ActiveBar1.Bands("mnuWindow").Tools("Window3").Visible = False
            End If
         End If
         
         If Secure.UserAdmnG4 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(6).Visible = False
                cmdSect(7).Visible = False
                z1(3).Visible = False
            Else
                cmdSect(6).Enabled = False
                cmdSect(7).Enabled = False
                MDISect.ActiveBar1.Bands("mnuWindow").Tools("Window4").Visible = False
            End If
         End If
         
         If Secure.UserAdmnG5 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(8).Visible = False
                cmdSect(9).Visible = False
                z1(4).Visible = False
            Else
                cmdSect(8).Enabled = False
                cmdSect(9).Enabled = False
                MDISect.ActiveBar1.Bands("mnuWindow").Tools("Window5").Visible = False
            End If
         End If
         
         If Secure.UserAdmnG6 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(10).Visible = False
                cmdSect(11).Visible = False
                z1(5).Visible = False
            Else
                cmdSect(10).Enabled = False
                cmdSect(11).Enabled = False
                MDISect.ActiveBar1.Bands("mnuWindow").Tools("Window6").Visible = False
            End If
         End If
      
         If Secure.UserAdmnG7 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(12).Visible = False
                cmdSect(13).Visible = False
                z1(6).Visible = False
            Else
                cmdSect(12).Enabled = False
                cmdSect(13).Enabled = False
            End If
         End If
      
      End If
      iUserIdx = GetSetting("Esi2000", "System", "UserProfileRec", iUserIdx)
   End If
   
End Sub

Private Sub HelpSearch()
   OpenHelpContext 5, True
   
End Sub

Private Sub EditSettings()
   '   On Error GoTo MenuEdit1
   '        If TypeOf ActiveForm.ActiveControl Is TextBox _
   '            Or TypeOf ActiveForm.ActiveControl Is ComboBox Then
   '                ActiveBar1.Bands("mnuEdit").Tools("EditSelect").Enabled = True
   '            If ActiveForm.ActiveControl.SelText = "" Then
   '                ActiveBar1.Bands("mnuEdit").Tools("EditCut").Enabled = False
   '                ActiveBar1.Bands("mnuEdit").Tools("EditCopy").Enabled = False
   '                ActiveBar1.Bands("mnuEdit").Tools("EditDelete").Enabled = False
   '            Else
   '                ActiveBar1.Bands("mnuEdit").Tools("EditCut").Enabled = True
   '                ActiveBar1.Bands("mnuEdit").Tools("EditCopy").Enabled = True
   '                ActiveBar1.Bands("mnuEdit").Tools("EditDelete").Enabled = True
   '            End If
   '            If Clipboard.GetText = "" Then
   '                ActiveBar1.Bands("mnuEdit").Tools("EditPaste").Enabled = False
   '            Else
   '                ActiveBar1.Bands("mnuEdit").Tools("EditPaste").Enabled = True
   '            End If
   '        Else
   '            ActiveBar1.Bands("mnuEdit").Tools("EditCut").Enabled = False
   '            ActiveBar1.Bands("mnuEdit").Tools("EditCopy").Enabled = False
   '            ActiveBar1.Bands("mnuEdit").Tools("EditDelete").Enabled = False
   '            ActiveBar1.Bands("mnuEdit").Tools("EditPaste").Enabled = False
   '            ActiveBar1.Bands("mnuEdit").Tools("EditSelect").Enabled = False
   '        End If
   '    Exit Sub
   '
   'MenuEdit1:
   '    On Error GoTo 0
   '    ActiveBar1.Bands("mnuEdit").Tools("EditCut").Enabled = False
   '    ActiveBar1.Bands("mnuEdit").Tools("EditCopy").Enabled = False
   '    ActiveBar1.Bands("mnuEdit").Tools("EditDelete").Enabled = False
   '    ActiveBar1.Bands("mnuEdit").Tools("EditPaste").Enabled = False
   '    ActiveBar1.Bands("mnuEdit").Tools("EditSelect").Enabled = False
   
End Sub


Private Sub HelpAbout()
   SysAbout.Show
   
End Sub


Private Sub HelpContents()
   OpenHelpContext 1000, True
   
End Sub


Private Sub WindowSettings()
   Dim iList As Integer
   If bUnloading = 0 Then CloseForms
   'See who's here
   
   iList = GetSetting("Esi2000", "Sections", "sale", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection1").Enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection1").Enabled = False
   
   iList = GetSetting("Esi2000", "Sections", "prod", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection2").Enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection2").Enabled = False
   
   iList = GetSetting("Esi2000", "Sections", "engr", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection3").Enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection3").Enabled = False
   
   iList = GetSetting("Esi2000", "Sections", "fina", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection4").Enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection4").Enabled = False
   
   iList = GetSetting("Esi2000", "Sections", "qual", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection5").Enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection5").Enabled = False
   
   iList = GetSetting("Esi2000", "Sections", "invc", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection6").Enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection6").Enabled = False
   
End Sub


Private Sub ActiveBar1_BandOpen(ByVal Band As ActiveBarLibraryCtl.Band)
   Select Case Band.Name
      Case "Sections"
         GetAppTitles
         '        Case "mnuEdit"
         '            EditSettings
      Case "mnuWindow"
         WindowSettings
         'Current Group is canceled if selected from bar
         cUR.CurrentGroup = ""
   End Select
   
End Sub

Private Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
   Dim iList As Integer
   On Error Resume Next
   Sleep 100
   Select Case Tool.Name
      Case "ReleaseNotes"
          Dim ret As Long
          ret = ShellExecute(Me.hwnd, "open", "http://www.fusionerp.net/category/release-notes/", vbNullString, vbNullString, 3)
          If ret < 32 Then MsgBox "There was an error when trying to open a default browser", vbCritical, "Error"
      Case "Databases"
         Load SysData
      Case "FileExit"
         CloseForms
         Unload Me
      Case "FilePrint"
         Cdi.ShowPrinter
      Case "FileReport"
         MouseCursor 13
         SysZoom.Show
      Case "FileSettings"
         SysSettings.Show 1
         '      Case "Window0"
         '          iHideRecent = 0
         '          SaveSetting "Esi2000", "Programs", "HideRecent", 0
         '          Recent.WindowState = 0
         '          Recent.Show
      Case "Window1"
         zGr1Admn.Show
      Case "Window2"
         zGr2Sale.Show
      Case "Window3"
         zGr3Padm.Show
      Case "Window4"
         zGr4Humn.Show
      Case "Window5"
         zGr5Matl.Show
      Case "Window6"
         zGr6Help.Show
      Case "DatabaseMaint"
         zGr7Maint.Show
      Case "HelpContents"
         HelpContents
      Case "HelpSearch"
         HelpSearch
      Case "HelpAbout"
         HelpAbout
      Case "HelpStatus"
      Case "FileRecent0"
         OpenFavorite Trim(Tool.Caption)
      Case "FileRecent1"
         OpenFavorite Trim(Tool.Caption)
      Case "FileRecent2"
         OpenFavorite Trim(Tool.Caption)
      Case "FileRecent3"
         OpenFavorite Trim(Tool.Caption)
      Case "FileRecent4"
         OpenFavorite Trim(Tool.Caption)
      Case "EditCut"
         Clipboard.SetText ActiveForm.ActiveControl.SelText
         ActiveForm.ActiveControl.SelText = ""
      Case "EditCopy"
         Clipboard.Clear
         Clipboard.SetText ActiveForm.ActiveControl.SelText
      Case "EditPaste"
         ActiveForm.ActiveControl.SelText = Clipboard.GetText(vbCFText)
      Case "EditDelete"
         ActiveForm.ActiveControl.SelText = ""
      Case "EditSelect"
         iList = Len(ActiveForm.ActiveControl.Text)
         ActiveForm.ActiveControl.SelStart = 0
         ActiveForm.ActiveControl.SelLength = iList
      Case "Favor1"
         OpenFavorite Trim(Tool.Caption)
      Case "Favor2"
         OpenFavorite Trim(Tool.Caption)
      Case "Favor3"
         OpenFavorite Trim(Tool.Caption)
      Case "Favor4"
         OpenFavorite Trim(Tool.Caption)
      Case "Favor5"
         OpenFavorite Trim(Tool.Caption)
      Case "Favor6"
         OpenFavorite Trim(Tool.Caption)
      Case "Favor7"
         OpenFavorite Trim(Tool.Caption)
      Case "Favor8"
         OpenFavorite Trim(Tool.Caption)
      Case "Favor9"
         OpenFavorite Trim(Tool.Caption)
      Case "Favor10"
         OpenFavorite Trim(Tool.Caption)
      Case "Favor11"
         OpenFavorite Trim(Tool.Caption)
      Case "Favor12"
         OpenFavorite Trim(Tool.Caption)
      Case "FavorBar"
         If iBarOnTop = 0 Then
            iBarOnTop = 1
            SideBar.Visible = False
            LeftBar.Visible = True
            TopBar.Visible = True
            ActiveBar1.Bands("Options").Tools("FavorBar").Caption = "Bar On Side"
         Else
            iBarOnTop = 0
            SideBar.Visible = True
            LeftBar.Visible = False
            TopBar.Visible = False
            ActiveBar1.Bands("Options").Tools("FavorBar").Caption = "Bar On Top"
         End If
         SaveSetting "Esi2000", "Programs", "BarOnTop", iBarOnTop
      Case "FavorTips"
         If iAutoTips = 0 Then
            iAutoTips = 1
            ActiveBar1.Bands("Options").Tools("FavorTips").Caption = "Auto Tips On"
         Else
            iAutoTips = 0
            SideBar.Visible = True
            ActiveBar1.Bands("Options").Tools("FavorTips").Caption = "Auto Tips Off"
         End If
         SaveSetting "Esi2000", "Programs", "AutoTipOn", iAutoTips
      Case "FavorOptAdd"
         SysFavorite.Show 1
      Case "WindowSection1"
         'AppActivate "ESI Sales", True
         AppActivate sAppTitles(1), True
         SendKeys "% x", True
      Case "WindowSection2"
         'AppActivate "ESI Production", True
         AppActivate sAppTitles(3), True
         SendKeys "% x", True
      Case "WindowSection3"
         'AppActivate "ESI Engineering", True
         AppActivate sAppTitles(2), True
         SendKeys "% x", True
      Case "WindowSection4"
         'AppActivate "ESI Finance", True
         AppActivate sAppTitles(6), True
         SendKeys "% x", True
      Case "WindowSection5"
         'AppActivate "ESI Quality", True
         AppActivate sAppTitles(5), True
         SendKeys "% x", True
      Case "WindowSection6"
         'AppActivate "ESI Inventory", True
         AppActivate sAppTitles(4), True
         SendKeys "% x", True
   End Select
   MouseCursor 0
   
End Sub

Private Sub ActiveBar1_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
   Cancel = True
   
End Sub


Private Sub cmdSect_Click(Index As Integer)
   bUserAction = True
   If cmdSect(Index) Then
      If Not bUnloading Then CloseForms
      Select Case Index
         Case 0, 2
            'NewTabs.Show
            zGr1Admn.Show
         Case 1, 3
            zGr2Sale.Show
         Case 4, 5
            zGr3Padm.Show
         Case 6, 7
            zGr4Humn.Show
         Case 8, 9
            zGr5Matl.Show
         Case 10, 11
            zGr6Help.Show
         Case 12, 13
            zGr7Maint.Show
      End Select
   End If
   
End Sub




Private Sub Logo_Click(Index As Integer)
   ClickedOnLogo
  
End Sub

Private Sub MDIForm_Activate()
    If bOnLoad = 1 Then
        bOnLoad = 0
        ActivateSection "EsiAdmn"
        UpdateDatabase
        'MM Not need here 9/5/2009
        'CheckSectionPermissions
        'New 1/29/04 to open last form
        If bOpenLastForm = 1 Then
           ' 3/19/05
           sCurrForm = sRecent(0)
           OpenFavorite sCurrForm
        Else
           OpenFavorite ""
        End If
    End If
    MouseCursor 0
   
End Sub

Private Sub MDIForm_Click()
   '
End Sub

Private Sub MDIForm_Initialize()
   bUserAction = True
   On Error Resume Next
   bResize = GetSetting("Esi2000", "System", "ResizeForm", bResize)
   If bResize = 0 Then ReSize1.Enabled = False
   FormInitialize
   
End Sub

Private Sub MDIForm_Load()
    MouseCursor 13
    GetRecentList "EsiAdmn"
    '11/23/04
    On Error Resume Next
    If bUnloading = 0 Then
        bOnLoad = 1
        ' MM 9/5/2009
        ' Open the database connection
        If Not OpenDBServer(False) Then
            End
        End If
        ' Check the security
        CheckSectionPermissions
        ' Show the MDI form
        Show
    End If
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   bUserAction = True
   
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   On Error Resume Next
'   sSql = "UPDATE Alerts SET ALERTMSG='' " _
'          & "WHERE ALERTREF=1"
'   RdoCon.Execute sSql, rdExecDirect
   bUnloading = 1
   If bOpenLastForm = 0 Then sCurrForm = ""
   UnLoadSection "admn", "EsiAdmn"
   
End Sub


Private Sub MDIForm_Resize()
   lScreenWidth = Screen.Width
   If WindowState <> 1 Then ResizeSection
   
End Sub

Private Sub MDIForm_Terminate()
   'End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   CloseFiles
   
End Sub

Private Sub OvrPanel_Click()
   If bInsertOn Then _
      ToggleInsertKey False _
      Else ToggleInsertKey True
   
End Sub

Private Sub SideBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   bUserAction = True
   
End Sub


Private Sub Timer1_Timer()
   tmePanel = Format(Time, "h:mm AM/PM")
   tmePanel.ToolTipText = Format(Date, "dddd mmmm dd,yyyy")
   
End Sub

Private Sub Timer2_Timer()
   
   If ShutdownTest Then
      Unload Me
   End If
   
'   Static iTimer As Integer
'   Static sLast As String
'   Dim bByte As Byte
'   Dim bResponse As Byte
'   Dim CloseApp As Long
'   Dim sMsg As String
'   Dim CurSection As String
'
'   If bUserAction Then
'      iTimer = 0
'      bUserAction = False
'      sLast = Format$(Time, "hh:mm AM/PM")
'   Else
'      iTimer = iTimer + 1
'   End If
'   If iTimer = 57 Then '57
'      If Not bUserAction Then
'         sMsg = GetTimeOut(sLast)
'         On Error Resume Next
'         bUserAction = True
'         bByte = InStr(LTrim$(MDISect.Caption), "-")
'         CurSection = " " & Left$(MDISect.Caption, bByte - 2)
'         CloseForms
'         RdoCon.Close
'         If tmePanel > "4:53 PM" Then
'            SaveSetting "Esi2000", "System", "CloseSection", App.Title
'            CloseApp = FindWindow(vbNullString, "ESI CloseSections")
'            If CloseApp = 0 Then
'               If Dir(sFilePath & "EsiExit.exe") <> "" Then _
'                      Shell sFilePath & "EsiExit.exe", vbNormalFocus
'            Else
'               AppActivate "ESI CloseSections", True
'               SendKeys "% x", True
'            End If
'         End If
'         bResponse = MsgBox(sMsg, ES_YESQUESTION, sSysCaption & CurSection)
'         If bResponse = vbYes Then
'            iTimer = 0
'            bUserAction = True
'            OpenSqlServer True
'         Else
'            Unload Me
'         End If
'      End If
'   End If
'
End Sub

Private Sub Timer3_Timer()
   'check the state of the insert key
   Dim iState As Integer
   If Forms.Count > 1 Then
      If Forms(1).Tag <> "TAB" Then _
               Timer3.Interval = 2000 Else _
               Timer3.Interval = 1000
   Else
      Timer3.Interval = 1000
   End If
   
   Dim bytKeys(255) As Byte
   'Get status of the 256 virtual keys
   GetKeyboardState bytKeys(0)
   
   ' Force the key state to Insert
   ' MM KEY_STATUS
   bytKeys(VK_INSERT) = 1
   'Set the keyboard state
   SetKeyboardState bytKeys(0)
   
   
   'Change a key
   iState = bytKeys(VK_INSERT)
   'iState = GetKeyState(vbKeyInsert)
   If iState = 1 Then
      bInsertOn = True
      OvrPanel = "INSERT"
      OvrPanel.ToolTipText = "Insert Text Is On (Click me) "
   Else
      bInsertOn = False
      OvrPanel = "OVER"
      OvrPanel.ToolTipText = "Overtype Text Is On (Click me)"
   End If
   
End Sub

Private Sub Timer4_Timer()
   Static b As Byte
   b = b + 1
   If b > 5 Then
      Timer4.Enabled = False
      Exit Sub
   End If
   
End Sub

Private Sub Timer5_Timer()
   Dim iRed As Integer
   Dim iGreen As Integer
   Dim iBlue As Integer
   iRed = GetSetting("Esi2000", "System", "SectionBackColorR", iRed)
   iGreen = GetSetting("Esi2000", "System", "SectionBackColorG", iGreen)
   iBlue = GetSetting("Esi2000", "System", "SectionBackColorB", iBlue)
   If iRed + iGreen + iBlue = 0 Then
      MDISect.BackColor = vbApplicationWorkspace
   Else
      MDISect.BackColor = RGB(iRed, iGreen, iBlue)
   End If
   GetSystemMessage
   
End Sub

Private Sub TopBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   bUserAction = True
   
End Sub
