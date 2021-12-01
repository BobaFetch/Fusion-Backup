VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MdiSect 
   BackColor       =   &H8000000C&
   Caption         =   "Sales"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   255
   ClientWidth     =   11880
   Icon            =   "MdiSale.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox LeftBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7890
      Left            =   1980
      ScaleHeight     =   7890
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   645
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.PictureBox BotPanel 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   340
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11880
      TabIndex        =   14
      Top             =   8535
      Width           =   11880
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
         TabIndex        =   18
         Top             =   24
         Width           =   6300
      End
      Begin VB.Label lblBotPanel 
         Height          =   336
         Left            =   100
         TabIndex        =   17
         Top             =   24
         Width           =   2532
      End
      Begin VB.Label tmePanel 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   324
         Left            =   10224
         TabIndex        =   16
         Top             =   24
         Width           =   720
      End
      Begin VB.Label OvrPanel 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OVER"
         Height          =   324
         Left            =   9384
         TabIndex        =   15
         Top             =   24
         Width           =   780
      End
   End
   Begin VB.PictureBox TopBar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   650
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   11880
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   11880
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiSale.frx":014A
         Height          =   588
         Index           =   5
         Left            =   3120
         Picture         =   "MdiSale.frx":0D8C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiSale.frx":19CE
         Height          =   588
         Index           =   7
         Left            =   2280
         Picture         =   "MdiSale.frx":26F4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiSale.frx":341A
         Height          =   588
         Index           =   4
         Left            =   1440
         Picture         =   "MdiSale.frx":4140
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiSale.frx":4E66
         Height          =   588
         Index           =   2
         Left            =   600
         Picture         =   "MdiSale.frx":59CC
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.Image Logo 
         Appearance      =   0  'Flat
         Height          =   765
         Index           =   1
         Left            =   7440
         Picture         =   "MdiSale.frx":6532
         Top             =   30
         Width           =   2430
      End
   End
   Begin VB.PictureBox SideBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7890
      Left            =   0
      ScaleHeight     =   7890
      ScaleWidth      =   1980
      TabIndex        =   8
      Top             =   648
      Visible         =   0   'False
      Width           =   1980
      Begin VB.CommandButton cmdSect 
         Height          =   348
         Index           =   11
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3720
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdSect 
         Height          =   348
         Index           =   9
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3720
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdSect 
         Height          =   348
         Index           =   8
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3240
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdSect 
         Height          =   348
         Index           =   10
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3240
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiSale.frx":C6AC
         Height          =   828
         Index           =   3
         Left            =   1020
         Picture         =   "MdiSale.frx":CECB
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiSale.frx":D6EA
         Height          =   828
         Index           =   6
         Left            =   100
         Picture         =   "MdiSale.frx":DF9D
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiSale.frx":E850
         Height          =   828
         Index           =   1
         Left            =   1020
         Picture         =   "MdiSale.frx":F184
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiSale.frx":FAB8
         Height          =   828
         Index           =   0
         Left            =   100
         Picture         =   "MdiSale.frx":103D3
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin Crystal.CrystalReport Crw 
         Left            =   1320
         Top             =   3450
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
         WindowState     =   1
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowProgressCtls=   0   'False
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Image imgPartList 
         Height          =   300
         Left            =   0
         Picture         =   "MdiSale.frx":10CEE
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgNewPart 
         Height          =   300
         Left            =   0
         Picture         =   "MdiSale.frx":1118C
         Top             =   4840
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgPartFind 
         Height          =   300
         Left            =   1440
         Picture         =   "MdiSale.frx":11627
         Top             =   5280
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgStandardComment 
         Height          =   300
         Left            =   1080
         Picture         =   "MdiSale.frx":11A61
         Top             =   5280
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XPrinter_small 
         Height          =   300
         Left            =   600
         Picture         =   "MdiSale.frx":11EB9
         Stretch         =   -1  'True
         Top             =   5280
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XDisplay 
         Height          =   315
         Left            =   600
         Picture         =   "MdiSale.frx":122F9
         Stretch         =   -1  'True
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image XPrinter 
         Height          =   300
         Left            =   960
         Picture         =   "MdiSale.frx":12716
         Stretch         =   -1  'True
         Top             =   4920
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Slips"
         Height          =   408
         Index           =   1
         Left            =   1020
         TabIndex        =   12
         ToolTipText     =   "Packing Slip Management"
         Top             =   1200
         Width           =   860
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Order Processing"
         Height          =   408
         Index           =   0
         Left            =   100
         TabIndex        =   11
         ToolTipText     =   "Customer Order Processing"
         Top             =   1200
         Width           =   860
      End
      Begin VB.Image XPPrinterUp 
         Height          =   300
         Left            =   600
         Picture         =   "MdiSale.frx":12B61
         Stretch         =   -1  'True
         Top             =   4320
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XPPrinterDn 
         Height          =   300
         Left            =   960
         Picture         =   "MdiSale.frx":12FAC
         Stretch         =   -1  'True
         Top             =   4290
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XPHelpUp 
         Height          =   300
         Left            =   600
         Picture         =   "MdiSale.frx":133F7
         Stretch         =   -1  'True
         Top             =   4650
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XPHelpDn 
         Height          =   300
         Left            =   960
         Picture         =   "MdiSale.frx":1352A
         Stretch         =   -1  'True
         Top             =   4650
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblBut 
         BackStyle       =   0  'Transparent
         Caption         =   "Commissions"
         Height          =   408
         Index           =   2
         Left            =   1000
         TabIndex        =   10
         ToolTipText     =   "Sales Commissions"
         Top             =   2520
         Width           =   9000
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bookings/ Backlog"
         Height          =   405
         Index           =   3
         Left            =   75
         TabIndex        =   9
         ToolTipText     =   "Booking And Backlog Reports"
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Logo 
         Height          =   1500
         Index           =   0
         Left            =   0
         Picture         =   "MdiSale.frx":1365D
         Top             =   6090
         Width           =   1935
      End
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   3960
      Top             =   2040
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   2040
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3000
      Top             =   2040
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   2040
      Top             =   1200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   0   'False
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8880
      FormDesignWidth =   11880
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   2520
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2040
      Top             =   2040
   End
   Begin MSComDlg.CommonDialog Cdi 
      Left            =   2040
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   4.38642e-38
   End
   Begin MSComctlLib.StatusBar RightPanel 
      Align           =   4  'Align Right
      Height          =   7890
      Left            =   11670
      TabIndex        =   24
      Top             =   645
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   13917
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
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   2040
      Top             =   600
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
      Bands           =   "MdiSale.frx":16F04
   End
End
Attribute VB_Name = "MdiSect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'7/22/04 Demo Help Loaded in Help Contents
Option Explicit
Public bUnloading As Byte
Dim bOnLoad As Byte

Private Sub CheckSectionPermissions()
   'Dim bByte As Byte
   Dim iList As Integer
   Dim iHideModBtn As Integer
   On Error Resume Next
   
   'let the programmer see everything
   If RunningInIDE Then
      InitializePermissions Secure, 1
      Secure.UserInitials = "MGR"         'masquerade as SYSMGR
   End If
   
   
   If bSecSet = 1 Then
      ' Check flag to Hide module buttons if user don't have permission
      iHideModBtn = GetHideModule()
      
      If Secure.UserSale <> 1 Then
         For iList = 0 To 7
            cmdSect(iList).Enabled = False
         Next
         For iList = 0 To 5
            MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(iList))).Visible = False
         Next
         For iList = 1 To 5
            MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window" & Trim(str(iList))).Visible = False
         Next
         For iList = 1 To 12
            MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(str(iList))).Visible = False
         Next
         SystemMsg.ForeColor = vbRed
         SystemMsg = "There Are No Section User Permissions"
      Else
         If Secure.UserSaleG1 <> 1 Then
            
            If (iHideModBtn = 1) Then
                cmdSect(0).Visible = False
                cmdSect(2).Visible = False
                lblBut(0).Visible = False
            Else
                cmdSect(0).Enabled = False
                cmdSect(2).Enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window1").Visible = False
            End If
         End If
         
         If Secure.UserSaleG2 <> 1 Then
            If (iHideModBtn = 1) Then
                cmdSect(1).Visible = False
                cmdSect(4).Visible = False
                lblBut(1).Visible = False
            Else
                cmdSect(1).Enabled = False
                cmdSect(4).Enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window2").Visible = False
            End If
         End If
         
         If Secure.UserSaleG3 <> 1 Then
            If (iHideModBtn = 1) Then
                cmdSect(6).Visible = False
                cmdSect(7).Visible = False
                lblBut(3).Visible = False
            Else
                cmdSect(6).Enabled = False
                cmdSect(7).Enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window3").Visible = False
            End If
            
         End If
         
         If Secure.UserSaleG4 <> 1 Then
            If (iHideModBtn = 1) Then
                cmdSect(3).Visible = False
                cmdSect(5).Visible = False
                lblBut(2).Visible = False
            Else
                cmdSect(3).Enabled = False
                cmdSect(5).Enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window4").Visible = False
            End If
         End If
         
         If Secure.UserSaleG5 <> 1 Then
            'cmdSect(6).Enabled = False
            'cmdSect(7).Enabled = False
            'MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window5").Visible = False
         End If
         
      End If
   End If
   
End Sub



Private Sub ActiveBar1_BandOpen(ByVal Band As ActiveBarLibraryCtl.Band)
   Select Case Band.Name
      Case "Sections"
         GetAppTitles
         ' Case "mnuFile"
         '      Sleep 100
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
          ret = ShellExecute(Me.hWnd, "open", "http://www.fusionerp.net/category/release-notes/", vbNullString, vbNullString, 3)
          If ret < 32 Then MsgBox "There was an error when trying to open a default browser", vbCritical, "Error"
      Case "Databases"
         Load SysData
      Case "FileExit"
         Unload Me
      Case "FilePrint"
         Cdi.ShowPrinter
      Case "FileReport"
         MouseCursor 13
         SysZoom.Show
      Case "FileSettings"
         SysSettings.Show 1
      Case "Window1"
         zGr1Sale.Show
      Case "Window2"
         zGr2Pack.Show
      Case "Window3"
         zGr3Book.Show
      Case "Window4"
         zGr4Comm.Show
      Case "HelpContents"
         HelpContents
      Case "HelpSearch"
         HelpSearch
      Case "HelpAbout"
         HelpAbout
      Case "HelpStatus"
      Case "FileRecent0"
         OpenFavorite Tool.Caption
      Case "FileRecent1"
         OpenFavorite Tool.Caption
      Case "FileRecent2"
         OpenFavorite Tool.Caption
      Case "FileRecent3"
         OpenFavorite Tool.Caption
      Case "FileRecent4"
         OpenFavorite Tool.Caption
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
         OpenFavorite Tool.Caption
      Case "Favor2"
         OpenFavorite Tool.Caption
      Case "Favor3"
         OpenFavorite Tool.Caption
      Case "Favor4"
         OpenFavorite Tool.Caption
      Case "Favor5"
         OpenFavorite Tool.Caption
      Case "Favor6"
         OpenFavorite Tool.Caption
      Case "Favor7"
         OpenFavorite Tool.Caption
      Case "Favor8"
         OpenFavorite Tool.Caption
      Case "Favor9"
         OpenFavorite Tool.Caption
      Case "Favor10"
         OpenFavorite Tool.Caption
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
         ' AppActivate "ESI Production", True
         AppActivate sAppTitles(3), True
         SendKeys "% x", True
      Case "WindowSection2"
         ' AppActivate "ESI Engineering", True
         AppActivate sAppTitles(2), True
         SendKeys "% x", True
      Case "WindowSection3"
         ' AppActivate "ESI Administration", True
         AppActivate sAppTitles(0), True
         SendKeys "% x", True
      Case "WindowSection4"
         ' AppActivate "ESI Finance", True
         AppActivate sAppTitles(6), True
         SendKeys "% x", True
      Case "WindowSection5"
         ' AppActivate "ESI Quality", True
         AppActivate sAppTitles(5), True
         SendKeys "% x", True
      Case "WindowSection6"
         ' AppActivate "ESI Inventory", True
         AppActivate sAppTitles(4), True
         SendKeys "% x", True
   End Select
   MouseCursor 0
   
End Sub

Private Sub ActiveBar1_GotFocus(ByVal Tool As ActiveBarLibraryCtl.Tool)
   '
   
End Sub

Private Sub ActiveBar1_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
   Cancel = True
   
End Sub





Private Sub cmdSect_Click(Index As Integer)
   If cmdSect(Index) Then
      If Not bUnloading Then CloseForms
      Select Case Index
         Case 0, 2
            zGr1Sale.Show
         Case 1, 4
            zGr2Pack.Show
         Case 3, 5
            zGr4Comm.Show
         Case 6, 7
            zGr3Book.Show
      End Select
   End If
   
End Sub




Private Sub Logo_Click(Index As Integer)
    ClickedOnLogo
End Sub

Private Sub MDIForm_Activate()
   If bOnLoad = 1 Then
      bOnLoad = 0
      ActivateSection "EsiSale"
      ' database update
      UpdateDatabase
      ' MM 9/5/2009
      'CheckSectionPermissions
      'New 1/29/04 to open last form
      If bOpenLastForm = 1 Then
         sCurrForm = sRecent(0)
         OpenFavorite sCurrForm
      Else
         OpenFavorite ""
      End If
   End If
   MouseCursor 0
   
End Sub

Private Sub MDIForm_Initialize()
   On Error Resume Next
   bResize = GetSetting("Esi2000", "System", "ResizeForm", bResize)
   If bResize = 0 Then ReSize1.Enabled = False
   bUnloading = False
   FormInitialize
   
End Sub

Private Sub MDIForm_Load()

    MouseCursor 13
    GetRecentList "EsiSale"
    '11/23/04
    On Error Resume Next
    If bUnloading = 0 Then
        bOnLoad = 1
        If Not OpenDBServer(False) Then
            End
        End If
        CheckSectionPermissions
        Show
    End If
   
End Sub


Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bUserAction = True
   
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   bUnloading = 1
   If bOpenLastForm = 0 Then sCurrForm = ""
   UnLoadSection "sale", "EsiSale"
   
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

Private Sub SideBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bUserAction = True
   
End Sub


Private Sub SystemMsg_Click()
   SystemMsg = ""
   
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
'         bByte = InStr(LTrim$(MdiSect.Caption), "-")
'         CurSection = " " & Left$(MdiSect.Caption, bByte - 2)
'         CloseForms
'         RdoCon.Close
'         If tmePanel > "4:51 PM" Then
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

Private Sub WindowSettings()
   Dim iList As Integer
   If Not bUnloading Then CloseForms
   'See who's here
   iList = GetSetting("Esi2000", "Sections", "prod", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection1").Enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection1").Enabled = False
   
   iList = GetSetting("Esi2000", "Sections", "engr", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection2").Enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection2").Enabled = False
   
   iList = GetSetting("Esi2000", "Sections", "admn", iList)
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

Private Sub HelpAbout()
   SysAbout.Show
   
End Sub

Private Sub HelpContents()
   OpenHelpContext 2000, True
   
End Sub

Private Sub HelpSearch()
   OpenHelpContext 5, True
   
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
   'Change a key
   
   ' Force the key state to Insert
   ' MM KEY_STATUS
   bytKeys(VK_INSERT) = 1
   'Set the keyboard state
   SetKeyboardState bytKeys(0)
   
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
      SystemMsg.Visible = True
      Timer4.Enabled = False
      Exit Sub
   End If
   If SystemMsg.Visible Then
      SystemMsg.Visible = False
   Else
      SystemMsg.Visible = True
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
      MdiSect.BackColor = vbApplicationWorkspace
   Else
      MdiSect.BackColor = RGB(iRed, iGreen, iBlue)
   End If
   GetSystemMessage
   
End Sub

Private Sub Toolbar1_Click()
   
End Sub

