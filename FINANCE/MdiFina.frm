VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.MDIForm MdiSect 
   BackColor       =   &H8000000C&
   Caption         =   "Financial Accounting"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11085
   Icon            =   "MdiFina.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox BotPanel 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   340
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11085
      TabIndex        =   30
      Top             =   8685
      Width           =   11085
      Begin VB.Label OvrPanel 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OVER"
         Height          =   330
         Left            =   8220
         TabIndex        =   34
         Top             =   30
         Width           =   780
      End
      Begin VB.Label tmePanel 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   9300
         TabIndex        =   33
         Top             =   30
         Width           =   1050
      End
      Begin VB.Label lblBotPanel 
         Height          =   330
         Left            =   105
         TabIndex        =   32
         Top             =   30
         Width           =   3600
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
         TabIndex        =   31
         Top             =   24
         Width           =   6300
      End
   End
   Begin VB.PictureBox LeftBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7995
      Left            =   1980
      ScaleHeight     =   7995
      ScaleWidth      =   195
      TabIndex        =   29
      Top             =   690
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.PictureBox Sidebar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7995
      Left            =   0
      ScaleHeight     =   7995
      ScaleWidth      =   1980
      TabIndex        =   10
      Top             =   690
      Width           =   1980
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":014A
         Height          =   828
         Index           =   0
         Left            =   60
         Picture         =   "MdiFina.frx":0AB9
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":1428
         Height          =   828
         Index           =   2
         Left            =   960
         Picture         =   "MdiFina.frx":1D75
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":26C2
         Height          =   828
         Index           =   4
         Left            =   60
         Picture         =   "MdiFina.frx":3020
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2880
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":397E
         Height          =   828
         Index           =   6
         Left            =   960
         Picture         =   "MdiFina.frx":424C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1560
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":4B1A
         Height          =   828
         Index           =   8
         Left            =   60
         Picture         =   "MdiFina.frx":51E0
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1560
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":58A6
         Height          =   828
         Index           =   10
         Left            =   960
         Picture         =   "MdiFina.frx":6013
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2880
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":6780
         Height          =   828
         Index           =   12
         Left            =   60
         Picture         =   "MdiFina.frx":6FC8
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4200
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":7810
         Height          =   828
         Index           =   14
         Left            =   960
         Picture         =   "MdiFina.frx":83F2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6600
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   860
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":8FD4
         Height          =   828
         Index           =   16
         Left            =   960
         Picture         =   "MdiFina.frx":9797
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4200
         UseMaskColor    =   -1  'True
         Width           =   860
      End
      Begin VB.Image imgNewPart 
         Height          =   300
         Left            =   0
         Picture         =   "MdiFina.frx":9F5A
         Top             =   7320
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgPartList 
         Height          =   300
         Left            =   360
         Picture         =   "MdiFina.frx":A3F5
         Top             =   7680
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgPartFind 
         Height          =   300
         Left            =   1560
         Picture         =   "MdiFina.frx":A893
         Top             =   7440
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgStandardComment 
         Height          =   300
         Left            =   720
         Picture         =   "MdiFina.frx":ACCD
         Top             =   7200
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XPrinter_small 
         Height          =   300
         Left            =   960
         Picture         =   "MdiFina.frx":B125
         Top             =   7680
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accounts Receivable"
         Height          =   405
         Index           =   0
         Left            =   40
         TabIndex        =   28
         Top             =   1095
         Width           =   865
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accounts Payable  "
         Height          =   405
         Index           =   1
         Left            =   960
         TabIndex        =   27
         Top             =   1095
         Width           =   825
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Closing"
         Height          =   405
         Index           =   2
         Left            =   0
         TabIndex        =   26
         Top             =   3720
         Width           =   1005
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Journals"
         Height          =   405
         Index           =   3
         Left            =   960
         TabIndex        =   25
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "General Ledger"
         Height          =   405
         Index           =   4
         Left            =   0
         TabIndex        =   24
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Product Costing"
         Height          =   405
         Index           =   5
         Left            =   960
         TabIndex        =   23
         Top             =   3720
         Width           =   1005
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Job Costing"
         Height          =   285
         Index           =   6
         Left            =   0
         TabIndex        =   22
         Top             =   5040
         Width           =   1005
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Analysis"
         Height          =   405
         Index           =   8
         Left            =   960
         TabIndex        =   21
         Top             =   5040
         Width           =   825
      End
      Begin VB.Image XPHelpUp 
         Height          =   300
         Left            =   120
         Picture         =   "MdiFina.frx":B565
         Top             =   7080
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image XPHelpDn 
         Height          =   300
         Left            =   480
         Picture         =   "MdiFina.frx":B698
         Top             =   7080
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image XPPrinterUp 
         Height          =   300
         Left            =   120
         Picture         =   "MdiFina.frx":B7CB
         Top             =   6720
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XPPrinterDn 
         Height          =   300
         Left            =   480
         Picture         =   "MdiFina.frx":BC16
         Top             =   6720
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblBut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Costing"
         Height          =   285
         Index           =   7
         Left            =   900
         TabIndex        =   20
         Top             =   7440
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Logo 
         Height          =   1500
         Index           =   0
         Left            =   0
         Picture         =   "MdiFina.frx":C061
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Image XDisplay 
         Height          =   300
         Left            =   180
         Picture         =   "MdiFina.frx":F908
         Top             =   7440
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image XPrinter 
         Height          =   300
         Left            =   480
         Picture         =   "MdiFina.frx":FD25
         Top             =   7440
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   7500
      Top             =   2520
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   7980
      Top             =   2520
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   8460
      Top             =   2520
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8940
      Top             =   2520
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   9420
      Top             =   2520
   End
   Begin Threed.SSPanel TopBar 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11085
      _Version        =   65536
      _ExtentX        =   19553
      _ExtentY        =   1217
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   0   'False
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":10170
         Height          =   615
         Index           =   17
         Left            =   5760
         Picture         =   "MdiFina.frx":10D52
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":11934
         Height          =   615
         Index           =   15
         Left            =   6480
         Picture         =   "MdiFina.frx":12516
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":130F8
         Height          =   615
         Index           =   13
         Left            =   5040
         Picture         =   "MdiFina.frx":13F0A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":14D1C
         Height          =   615
         Index           =   11
         Left            =   4320
         Picture         =   "MdiFina.frx":1595E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":165A0
         Height          =   615
         Index           =   9
         Left            =   2160
         Picture         =   "MdiFina.frx":171E2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":17E24
         Height          =   615
         Index           =   7
         Left            =   2880
         Picture         =   "MdiFina.frx":18A66
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":196A8
         Height          =   615
         Index           =   5
         Left            =   3600
         Picture         =   "MdiFina.frx":1A24E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":1ADF4
         Height          =   615
         Index           =   3
         Left            =   1440
         Picture         =   "MdiFina.frx":1BA36
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.CommandButton cmdSect 
         DownPicture     =   "MdiFina.frx":1C678
         Height          =   615
         Index           =   1
         Left            =   720
         Picture         =   "MdiFina.frx":1D2BA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin Crystal.CrystalReport crw 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Image Logo 
         Height          =   495
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   1515
      End
      Begin VB.Image Logo 
         Height          =   765
         Index           =   1
         Left            =   7320
         Picture         =   "MdiFina.frx":1DEFC
         Top             =   60
         Width           =   2430
      End
   End
   Begin MSComDlg.CommonDialog Cdi 
      Left            =   1920
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "IIF"
      DialogTitle     =   "Export AR Activity To ."
      Filter          =   "IIF"
      FontSize        =   4.38642e-38
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   0   'False
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9030
      FormDesignWidth =   11085
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   1920
      Top             =   720
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
      Bands           =   "MdiFina.frx":24076
   End
End
Attribute VB_Name = "MdiSect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bUnloading As Byte
Dim bOnLoad As Byte
Dim bOpenLastForm As Byte

Public Sub CheckSectionPermissions()
   Dim i As Integer
   Dim iHideModBtn As Integer
   
   
   'let the programmer see everything
   'if Runninginide Then
   If RunningInIDE Then
      InitializePermissions Secure, 1
      Secure.UserInitials = "MGR"         'masquerade as SYSMGR
   End If
   
   If bSecSet = 1 Then
      
      ' Check flag to Hide module buttons if user don't have permission
      iHideModBtn = GetHideModule()
      
      If Secure.UserFina <> 1 Then
         For i = 0 To 15
            cmdSect(i).enabled = False
         Next
         For i = 0 To 4
            MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(i))).Visible = False
         Next
         For i = 1 To 7
            MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window" & Trim(str(i))).Visible = False
         Next
         For i = 1 To 12
            MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(str(i))).Visible = False
         Next
         SystemMsg.ForeColor = vbRed
         SystemMsg = "There Are No Section User Permissions"
      Else
         
         If Secure.UserFinaG1 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(0).Visible = False
                cmdSect(1).Visible = False
                lblBut(0).Visible = False
            Else
                cmdSect(0).enabled = False
                cmdSect(1).enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window1").Visible = False
            End If
         End If
         
         If Secure.UserFinaG2 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(2).Visible = False
                cmdSect(3).Visible = False
                lblBut(1).Visible = False
            Else
                cmdSect(2).enabled = False
                cmdSect(3).enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window2").Visible = False
            End If
         End If
         
         If Secure.UserFinaG3 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(8).Visible = False
                cmdSect(9).Visible = False
                lblBut(4).Visible = False
            Else
                cmdSect(8).enabled = False
                cmdSect(9).enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window3").Visible = False
            End If
         End If
         
         If Secure.UserFinaG4 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(6).Visible = False
                cmdSect(7).Visible = False
                lblBut(3).Visible = False
            Else
                cmdSect(6).enabled = False
                cmdSect(7).enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window4").Visible = False
            End If
         End If
         
         If Secure.UserFinaG5 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(4).Visible = False
                cmdSect(5).Visible = False
                lblBut(2).Visible = False
            Else
                cmdSect(4).enabled = False
                cmdSect(5).enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window5").Visible = False
            End If
         End If
         
         If Secure.UserFinaG6 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(10).Visible = False
                cmdSect(11).Visible = False
                lblBut(5).Visible = False
            Else
                cmdSect(10).enabled = False
                cmdSect(11).enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window6").Visible = False
            End If
         End If
         
         If Secure.UserFinaG7 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(12).Visible = False
                cmdSect(13).Visible = False
                lblBut(6).Visible = False
            Else
                cmdSect(12).enabled = False
                cmdSect(13).enabled = False
                MdiSect.ActiveBar1.Bands("mnuWindow").Tools("Window7").Visible = False
            End If
         End If
         
         If Secure.UserFinaG8 <> 1 Then
            If (iHideModBtn = 1) Then
                ' Hide this button if hide flag is set
                cmdSect(16).Visible = False
                cmdSect(17).Visible = False
                lblBut(8).Visible = False
            Else
                cmdSect(16).enabled = False
                cmdSect(17).enabled = False
            End If
         End If
      End If
   End If
   
End Sub

Public Sub CheckButton(Index)
   MouseCursor 13
   'Select Case cmdOpt(Index).Caption
   '    Case "Edit Tools"
   '        diaEdtEx.Show
   '    Case "New Tool"
   '        diaNewEx.Show
   'End Select
   MouseCursor 0
End Sub

Private Sub ActiveBar1_BandOpen(ByVal Band As ActiveBarLibraryCtl.Band)
   Select Case Band.Name
      Case "Sections"
         GetAppTitles
      Case "mnuFile"
      Case "mnuEdit"
         EditSettings
      Case "mnuWindow"
         WindowSettings
         'Current Group is canceled if selected from bar
         cUR.CurrentGroup = ""
   End Select
   
End Sub

Private Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
   Dim i%
   On Error Resume Next
   Select Case Tool.Name
      Case "ReleaseNotes"
         Dim ret As Long
         ret = ShellExecute(Me.hWnd, "open", "http://www.fusionerp.net/category/release-notes/", vbNullString, vbNullString, 3)
         If ret < 32 Then MsgBox "There was an error when trying to open a default browser", vbCritical, "Error"
      Case "FileExit"
         CloseForms
         Unload Me
      Case "FilePrint"
         Cdi.ShowPrinter
      Case "FileReport"
         MouseCursor 13
         diaZoom.Show
      Case "FileSettings"
         diaSetng.Show 1
      Case "Databases"
         Load SysData
      Case "Window1"
         tabArec.Show
      Case "Window2"
         tabApay.Show
      Case "Window3"
         tabClos.Show
      Case "Window4"
         tabJorn.Show
      Case "Window5"
         tabGenl.Show
      Case "Window6"
         tabScst.Show
      Case "Window7"
         tabJcst.Show
      Case "Window8"
         tabLcst.Show
      Case "SalesAnalysis"
         tabSalesAnalysis.Show
      Case "HelpContents"
         HelpContents
      Case "HelpSearch"
         HelpSearch
      Case "HelpAbout"
         HelpAbout
      Case "HelpStatus"
         Ready.msg = MdiSect.Caption & " Is Ready"
         Ready.Show
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
         i = Len(ActiveForm.ActiveControl.Text)
         ActiveForm.ActiveControl.SelStart = 0
         ActiveForm.ActiveControl.SelLength = i%
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
         OpenFavorite Trim(Tool.Caption)
      Case "Favor11"
         OpenFavorite Trim(Tool.Caption)
      Case "Favor12"
         OpenFavorite Trim(Tool.Caption)
      Case "FavorBar"
         If iBarOnTop = 0 Then
            iBarOnTop = 1
            Sidebar.Visible = False
            TopBar.Visible = True
            ActiveBar1.Bands("Options").Tools("FavorBar").Caption = "Bar On Side"
         Else
            iBarOnTop = 0
            Sidebar.Visible = True
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
            Sidebar.Visible = True
            ActiveBar1.Bands("Options").Tools("FavorTips").Caption = "Auto Tips Off"
         End If
         SaveSetting "Esi2000", "Programs", "AutoTipOn", iAutoTips
      Case "FavorOptAdd"
         diaFavor.Show 1
      Case "WindowSection1"
         AppActivate "ESI Sales", True
         SendKeys "% x", True
      Case "WindowSection2"
         AppActivate "ESI Production", True
         SendKeys "% x", True
      Case "WindowSection3"
         AppActivate "ESI Engineering", True
         SendKeys "% x", True
      Case "WindowSection4"
         AppActivate "ESI Administration", True
         SendKeys "% x", True
      Case "WindowSection5"
         AppActivate "ESI Financial Accounting", True
         SendKeys "% x", True
      Case "WindowSection6"
         AppActivate "ESI Quality", True
         SendKeys "% x", True
      Case "WindowSection7"
         AppActivate "ESI Inventory", True
         SendKeys "% x", True
   End Select
   
End Sub

Private Sub ActiveBar1_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
   Cancel = True
   
End Sub


Private Sub cmdSect_Click(Index As Integer)
   cmdSect(Index) = False
   If Not bUnloading Then CloseForms
   Select Case Index
      Case 0, 1
         tabArec.Show
      Case 2, 3
         tabApay.Show
      Case 4, 5
         tabClos.Show
      Case 6, 7
         tabJorn.Show
      Case 8, 9
         tabGenl.Show
      Case 10, 11
         tabScst.Show
      Case 12, 13
         tabJcst.Show
      Case 14, 15
         tabLcst.Show
      Case 16, 17
         tabSalesAnalysis.Show
   End Select
   
End Sub

Private Sub Image1_Click(Index As Integer)
   HelpAbout
   
End Sub

Private Sub Image2_Click(Index As Integer)
   HelpAbout
   
End Sub


Private Sub Logo_Click(Index As Integer)
    ClickedOnLogo
   
End Sub

Private Sub MDIForm_Activate()
    If bOnLoad = 1 Then
        bOnLoad = 0
        ActivateSection "EsiFina"
        ' Update database
        UpdateDatabase
        'MM Not need here 9/5/2009
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
   Dim wFlags As Long
   Dim hMenu As Long
   
   On Error Resume Next
   bOpenLastForm = GetSetting("Esi2000", "System", "Reopenforms", bOpenLastForm)
   bResize = GetSetting("Esi2000", "System", "ResizeForm", bResize)
   If bResize = 0 Then ReSize1.enabled = False
   bUnloading = False
   wFlags = 0
   hMenu = GetSystemMenu(hWnd, wFlags)
   FormInitialize
   
   'Image1.Left = 100
   'Image2.Left = 580
   'Image3.Left = 1060
   'Image4.Left = 1660
End Sub

Private Sub MDIForm_Load()
    MouseCursor 13
    Dim sYear As String
    
    bResize = GetSetting("Esi2000", "System", "ResizeForm", bResize)
    SaveSetting "Esi2000", "AppTitle", "fina", "ESI Finance"
    
    MouseCursor 13
    sYear = Format$(Now, "yyyy")
    GetRecentList "EsiFina"
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

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, _
                              X As Single, Y As Single)
   bUserAction = True
   
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   bUnloading = 1
   If bOpenLastForm = 0 Then sCurrForm = ""
   UnLoadSection "fina", "EsiFina"
End Sub

Private Sub MDIForm_Resize()
   lScreenWidth = Screen.Width
   If WindowState <> 1 Then ResizeSection
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   Dim sWindows As String
   On Error Resume Next
   JetDb.Close
   Set JetWkSpace = Nothing
   Set JetDb = Nothing
   'Snuff the temp database
   'Make a brand new one on the next visit
   sWindows = GetWindowsDir()
   
   'leave for Larry to debug
   'If Dir(sWindows & "\temp\esifina.mdb") <> "" Then _
   'Kill sWindows & "\temp\esifina.mdb"
   CloseFiles
   
End Sub

Private Sub OvrPanel_Click()
   
   'User clicks the Panel
   If bInsertOn Then _
      ToggleInsertKey False _
      Else ToggleInsertKey True
   
   
End Sub

Private Sub SideBar_MouseMove(Button As Integer, Shift As Integer, _
                              X As Single, Y As Single)
   bUserAction = True
   
End Sub


Private Sub SystemMsg_DblClick()
   SystemMsg = ""
   
End Sub

Public Sub EditSettings()
   On Error GoTo MenuEdit1
   If TypeOf ActiveForm.ActiveControl Is TextBox _
         Or TypeOf ActiveForm.ActiveControl Is ComboBox Then
      ActiveBar1.Bands("mnuEdit").Tools("EditSelect").enabled = True
      If ActiveForm.ActiveControl.SelText = "" Then
         ActiveBar1.Bands("mnuEdit").Tools("EditCut").enabled = False
         ActiveBar1.Bands("mnuEdit").Tools("EditCopy").enabled = False
         ActiveBar1.Bands("mnuEdit").Tools("EditDelete").enabled = False
      Else
         ActiveBar1.Bands("mnuEdit").Tools("EditCut").enabled = True
         ActiveBar1.Bands("mnuEdit").Tools("EditCopy").enabled = True
         ActiveBar1.Bands("mnuEdit").Tools("EditDelete").enabled = True
      End If
      If Clipboard.GetText = "" Then
         ActiveBar1.Bands("mnuEdit").Tools("EditPaste").enabled = False
      Else
         ActiveBar1.Bands("mnuEdit").Tools("EditPaste").enabled = True
      End If
   Else
      ActiveBar1.Bands("mnuEdit").Tools("EditCut").enabled = False
      ActiveBar1.Bands("mnuEdit").Tools("EditCopy").enabled = False
      ActiveBar1.Bands("mnuEdit").Tools("EditDelete").enabled = False
      ActiveBar1.Bands("mnuEdit").Tools("EditPaste").enabled = False
      ActiveBar1.Bands("mnuEdit").Tools("EditSelect").enabled = False
   End If
   Exit Sub
   
MenuEdit1:
   ActiveBar1.Bands("mnuEdit").Tools("EditCut").enabled = False
   ActiveBar1.Bands("mnuEdit").Tools("EditCopy").enabled = False
   ActiveBar1.Bands("mnuEdit").Tools("EditDelete").enabled = False
   ActiveBar1.Bands("mnuEdit").Tools("EditPaste").enabled = False
   ActiveBar1.Bands("mnuEdit").Tools("EditSelect").enabled = False
   
End Sub

Public Sub WindowSettings()
   Dim iList
   If Not bUnloading Then CloseForms
   'See who's here
   
   iList = GetSetting("Esi2000", "Sections", "sale", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection1").enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection1").enabled = False
   
   iList = GetSetting("Esi2000", "Sections", "prod", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection2").enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection2").enabled = False
   
   iList = GetSetting("Esi2000", "Sections", "engr", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection3").enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection3").enabled = False
   
   iList = GetSetting("Esi2000", "Sections", "admn", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection4").enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection4").enabled = False
   
   '   iList = GetSetting("Esi2000", "Sections", "fina", iList)
   '   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection5").Enabled = True _
   '        Else ActiveBar1.Bands("Sections").Tools("WindowSection5").Enabled = False
   
   iList = GetSetting("Esi2000", "Sections", "qual", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection6").enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection6").enabled = False
   
   iList = GetSetting("Esi2000", "Sections", "invc", iList)
   If iList = 1 Then ActiveBar1.Bands("Sections").Tools("WindowSection7").enabled = True _
              Else ActiveBar1.Bands("Sections").Tools("WindowSection7").enabled = False
   
End Sub

Public Sub HelpContents()
   On Error GoTo DiaErr1
   OpenHelpContext "ES2000"
   Exit Sub
DiaErr1:
   MsgBox "That Web Page Cannot Be Located.", _
      vbInformation, Caption
End Sub

Public Sub HelpSearch()
   On Error GoTo DiaErr1
   OpenHelpContext "ES2000"
   Exit Sub
DiaErr1:
   MsgBox "That Web Page Cannot Be Located.", _
      vbInformation, Caption
End Sub

Public Sub HelpAbout()
   SysAbout.Show
   
End Sub

Private Sub Timer1_Timer()
   tmePanel = Format(Time, "h:mm AM/PM")
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
'   Timer2.enabled = False '@@@
'   If bUserAction Then
'      iTimer = 0
'      bUserAction = False
'      sLast = Format$(Time, "hh:mm AM/PM")
'   Else
'      iTimer = iTimer + 1
'   End If
'   If iTimer > 56 Then '57
'      If Not bUserAction Then
'         sMsg = GetTimeOut(sLast)
'         On Error Resume Next
'         bUserAction = True
'         bByte = InStr(LTrim$(MdiSect.Caption), "-")
'         CurSection = " " & Left$(MdiSect.Caption, bByte - 2)
'         CloseForms
'         RdoCon.Close
'         If tmePanel > "4:55 PM" Then
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
   'check the state of the insert key
   Dim iState As Integer
   If Forms.count > 1 Then
      If Forms(1).Tag <> "TAB" Then _
               Timer3.Interval = 2000 Else _
               Timer3.enabled = False
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
      OvrPanel.ToolTipText = "Overtype Text Is On (Click me)"
   End If
   
End Sub


Private Sub Timer4_Timer()
   Static b As Byte
   b = b + 1
   If b > 5 Then
      SystemMsg.Visible = True
      Timer4.enabled = False
      Exit Sub
   End If
   If SystemMsg.Visible Then
      SystemMsg.Visible = False
   Else
      SystemMsg.Visible = True
   End If
   
End Sub


Private Sub Timer5_Timer()
   GetSystemMessage
End Sub


Private Sub TmePanel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   tmePanel.ToolTipText = Format(Date, "dddd, mmmm dd,yyyy")
End Sub
