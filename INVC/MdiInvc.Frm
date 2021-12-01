VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MdiSect 
   BackColor       =   &H8000000C&
   Caption         =   "Inventory Control"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "MdiInvc.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox LeftBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7995
      Left            =   1980
      ScaleHeight     =   7995
      ScaleWidth      =   195
      TabIndex        =   22
      Top             =   648
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.PictureBox BotPanel 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   340
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11280
      TabIndex        =   7
      Top             =   8640
      Width           =   11280
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
         Height          =   340
         Left            =   2540
         TabIndex        =   11
         Top             =   20
         Width           =   6300
      End
      Begin VB.Label lblBotPanel 
         Height          =   336
         Left            =   100
         TabIndex        =   10
         Top             =   20
         Width           =   2532
      End
      Begin VB.Label tmePanel 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   324
         Left            =   10100
         TabIndex        =   9
         Top             =   20
         Width           =   720
      End
      Begin VB.Label OvrPanel 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OVER"
         Height          =   324
         Left            =   9260
         TabIndex        =   8
         ToolTipText     =   "Insert Or Overtype Text - Click To Change"
         Top             =   20
         Width           =   780
      End
   End
   Begin VB.PictureBox TopBar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   650
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   11280
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   11280
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiInvc.frx":014A
         Height          =   588
         Index           =   11
         Left            =   3960
         Picture         =   "MdiInvc.frx":0D8C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Lot Tracking Administration"
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.CommandButton cmdSect 
         DisabledPicture =   "MdiInvc.frx":19CE
         DownPicture     =   "MdiInvc.frx":2534
         Height          =   588
         Index           =   1
         Left            =   3120
         Picture         =   "MdiInvc.frx":309A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Inventory Managment - Cycle Counting"
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiInvc.frx":3C00
         Height          =   588
         Index           =   2
         Left            =   2280
         Picture         =   "MdiInvc.frx":4BDE
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Material Goods Receiving"
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiInvc.frx":5BBC
         Height          =   588
         Index           =   3
         Left            =   1440
         Picture         =   "MdiInvc.frx":6B9A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Material Control"
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiInvc.frx":7B78
         Height          =   588
         Index           =   9
         Left            =   600
         Picture         =   "MdiInvc.frx":7E82
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Inventory Control"
         Top             =   20
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.Image Logo 
         Height          =   765
         Index           =   1
         Left            =   7800
         Picture         =   "MdiInvc.frx":818C
         Top             =   15
         Width           =   2430
      End
   End
   Begin VB.PictureBox SideBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7995
      Left            =   0
      ScaleHeight     =   7995
      ScaleWidth      =   1980
      TabIndex        =   0
      Top             =   648
      Visible         =   0   'False
      Width           =   1980
      Begin VB.CommandButton cmdSect 
         Height          =   348
         Index           =   5
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4200
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdSect 
         Height          =   348
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4200
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiInvc.frx":E306
         Height          =   828
         Index           =   10
         Left            =   120
         Picture         =   "MdiInvc.frx":EAA6
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Lot Tracking Administration"
         Top             =   2760
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiInvc.frx":F246
         Height          =   828
         Index           =   0
         Left            =   1020
         Picture         =   "MdiInvc.frx":FBB6
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Inventory Managment - Cycle Counting"
         Top             =   1560
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiInvc.frx":10526
         Height          =   828
         Index           =   6
         Left            =   120
         Picture         =   "MdiInvc.frx":10E26
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Material Goods Receiving"
         Top             =   1560
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiInvc.frx":11726
         Height          =   828
         Index           =   7
         Left            =   1020
         Picture         =   "MdiInvc.frx":11FB9
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Material Control"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiInvc.frx":1284C
         Height          =   828
         Index           =   8
         Left            =   120
         Picture         =   "MdiInvc.frx":130DE
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Inventory Control"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.Image imgPartList 
         Height          =   300
         Left            =   0
         Picture         =   "MdiInvc.frx":13970
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgNewPart 
         Height          =   300
         Left            =   0
         Picture         =   "MdiInvc.frx":13E0E
         Top             =   4840
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgStandardComment 
         Height          =   300
         Left            =   1080
         Picture         =   "MdiInvc.frx":142A9
         Top             =   6120
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XPrinter_small 
         Height          =   300
         Left            =   0
         Picture         =   "MdiInvc.frx":14701
         Stretch         =   -1  'True
         Top             =   6000
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgPartFind 
         Height          =   300
         Left            =   600
         Picture         =   "MdiInvc.frx":14B41
         Top             =   6120
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XPrinter 
         Height          =   300
         Left            =   480
         Picture         =   "MdiInvc.frx":14F7B
         Top             =   5640
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XDisplay 
         Height          =   300
         Left            =   120
         Picture         =   "MdiInvc.frx":153C6
         Top             =   5640
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Receiving"
         Height          =   288
         Index           =   2
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Material Receiving"
         Top             =   2376
         Width           =   840
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
         Height          =   288
         Index           =   3
         Left            =   1020
         TabIndex        =   4
         ToolTipText     =   "Material Control"
         Top             =   1176
         Width           =   840
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory"
         Height          =   288
         Index           =   4
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Inventory Control"
         Top             =   1176
         Width           =   900
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Tracking"
         Height          =   288
         Index           =   5
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Lot Tracking"
         Top             =   3696
         Width           =   900
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory Mgmt"
         Height          =   492
         Index           =   4
         Left            =   960
         TabIndex        =   1
         ToolTipText     =   "Inventory Management - Cycle Counting"
         Top             =   2376
         Width           =   1008
      End
      Begin VB.Image XPHelpDn 
         Height          =   300
         Left            =   480
         Picture         =   "MdiInvc.frx":157E3
         Top             =   5250
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image XPHelpUp 
         Height          =   300
         Left            =   120
         Picture         =   "MdiInvc.frx":15916
         Top             =   5250
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image XPPrinterDn 
         Height          =   300
         Left            =   480
         Picture         =   "MdiInvc.frx":15A49
         Top             =   4890
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XPPrinterUp 
         Height          =   300
         Left            =   120
         Picture         =   "MdiInvc.frx":15E94
         Top             =   4890
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image Logo 
         Height          =   1500
         Index           =   0
         Left            =   0
         Picture         =   "MdiInvc.frx":162DF
         Top             =   6570
         Width           =   1935
      End
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   3960
      Top             =   1800
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   1800
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3000
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   2520
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2040
      Top             =   1800
   End
   Begin MSComDlg.CommonDialog Cdi 
      Left            =   1920
      Top             =   3210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   4.38642e-38
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   2000
      Top             =   1200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   0   'False
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8985
      FormDesignWidth =   11280
   End
   Begin MSComctlLib.StatusBar RightPanel 
      Align           =   4  'Align Right
      Height          =   7995
      Left            =   11070
      TabIndex        =   25
      Top             =   645
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   14102
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
      Left            =   2400
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
      Bands           =   "MdiInvc.frx":19B86
   End
End
Attribute VB_Name = "MdiSect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
Option Explicit
Public bUnloading As Byte
Dim bOnLoad As Byte
Public cCRViewer As EsCrystalRptViewer
Public Function CreateCRObject()

    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
End Function
Private Sub CheckSectionPermissions()
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
      
      If Secure.UserInvc <> 1 Then
         For iList = 0 To 11
            cmdSect(iList).Enabled = False
         Next
         For iList = 0 To 4
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
         If Secure.UserInvcG1 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(8).Visible = False
                cmdSect(9).Visible = False
                lblBut(4).Visible = False
            Else
                cmdSect(8).Enabled = False
                cmdSect(9).Enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window1").Visible = False
            End If
         End If
         
         If Secure.UserInvcG2 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(3).Visible = False
                cmdSect(7).Visible = False
                lblBut(3).Visible = False
            Else
                cmdSect(3).Enabled = False
                cmdSect(7).Enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window2").Visible = False
            End If
         End If
         
         If Secure.UserInvcG3 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(2).Visible = False
                cmdSect(6).Visible = False
                lblBut(2).Visible = False
            Else
                cmdSect(2).Enabled = False
                cmdSect(6).Enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window3").Visible = False
            End If
         End If
         
         If Secure.UserInvcG4 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(0).Visible = False
                cmdSect(1).Visible = False
                Z1(4).Visible = False
            Else
                cmdSect(0).Enabled = False
                cmdSect(1).Enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window4").Visible = False
            End If
         End If
         
         If Secure.UserInvcG5 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(10).Visible = False
                cmdSect(11).Visible = False
                Z1(5).Visible = False
            Else
                cmdSect(10).Enabled = False
                cmdSect(11).Enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window5").Visible = False
            End If
         End If
      End If
   End If
   
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
          ret = ShellExecute(Me.hWnd, "open", "http://www.fusionerp.net/category/release-notes/", vbNullString, vbNullString, 3)
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
         '       Case "Window0"
         '           iHideRecent = 0
         '           SaveSetting "Esi2000", "Programs", "HideRecent", 0
         '           Recent.WindowState = 0
         '           Recent.Show
      Case "Window1"
         zGr1Invc.Show
      Case "Window2"
         zGr2Pick.Show
      Case "Window3"
         zGr3Recv.Show
      Case "Window4"
         zGr4Matl.Show
      Case "Window5"
         zGr5Lots.Show
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
            LeftBar.Visible = True
            SideBar.Visible = False
            TopBar.Visible = True
            ActiveBar1.Bands("Options").Tools("FavorBar").Caption = "Bar On Side"
         Else
            iBarOnTop = 0
            LeftBar.Visible = False
            SideBar.Visible = True
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
         'AppActivate "ESI Administration", True
         AppActivate sAppTitles(0), True
         SendKeys "% x", True
      Case "WindowSection5"
         'AppActivate "ESI Finance", True
         AppActivate sAppTitles(6), True
         SendKeys "% x", True
      Case "WindowSection6"
         'AppActivate "ESI Quality", True
         AppActivate sAppTitles(5), True
         SendKeys "% x", True
   End Select
   MouseCursor 0
   
End Sub

Private Sub ActiveBar1_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
   Cancel = True
   
End Sub


Private Sub cmdSect_Click(Index As Integer)
   If cmdSect(Index) Then
      If Not bUnloading Then CloseForms
      Select Case Index
         Case 0, 1
            zGr4Matl.Show
         Case 2, 6
            zGr3Recv.Show
         Case 10, 11
            zGr5Lots.Show
         Case 3, 7
            zGr2Pick.Show
         Case 8, 9
            zGr1Invc.Show
      End Select
   End If
   
End Sub





Private Sub Logo_Click(Index As Integer)
    ClickedOnLogo
End Sub

Private Sub MDIForm_Activate()
    If bOnLoad Then
        bOnLoad = 0
        ActivateSection "EsiInvc"
        ' Update the database
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

Private Sub MDIForm_Initialize()
   On Error Resume Next
   bResize = GetSetting("Esi2000", "System", "ResizeForm", bResize)
   If bResize = 0 Then ReSize1.Enabled = False
   bUnloading = False
   FormInitialize
   
End Sub

Private Sub MDIForm_Load()
    MouseCursor 13
    GetRecentList "EsiInvc"
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


Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bUserAction = True
   
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   bUnloading = 1
   If bOpenLastForm = 0 Then sCurrForm = ""
   UnLoadSection "invc", "EsiInvc"
   
End Sub

Private Sub MDIForm_Resize()
   lScreenWidth = Screen.Width
   If WindowState <> 1 Then ResizeSection
   
End Sub

Private Sub MDIForm_Terminate()
   'end
   
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   Dim sWindows As String
   On Error Resume Next
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
End Sub

Private Sub WindowSettings()
   Dim iList As Integer
   If Not bUnloading Then CloseForms
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
   
   iList = GetSetting("Esi2000", "Sections", "admn", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection4").Enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection4").Enabled = False
   
   iList = GetSetting("Esi2000", "Sections", "fina", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection5").Enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection5").Enabled = False
   
   iList = GetSetting("Esi2000", "Sections", "qual", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection6").Enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection6").Enabled = False
   
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

Private Sub HelpContents()
   OpenHelpContext 5000, True
   
End Sub

Private Sub HelpSearch()
   OpenHelpContext 5, True
   
End Sub

Private Sub HelpAbout()
   SysAbout.Show
   
End Sub

Private Sub Timer3_Timer()
   'check the state of the insert key
   Dim iState As Integer
   If Forms.count > 1 Then
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
      OvrPanel.ToolTipText = "Insert Text Is On (Click me)"
   Else
      bInsertOn = False
      OvrPanel = "OVER"
      OvrPanel.ToolTipText = "Overtype Text Is On (Click me) "
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

