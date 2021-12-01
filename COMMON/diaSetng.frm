VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaSetng
   BorderStyle = 3 'Fixed Dialog
   Caption = "Workstation Settings"
   ClientHeight = 4635
   ClientLeft = 780
   ClientTop = 855
   ClientWidth = 6795
   ControlBox = 0 'False
   Icon = "Diasetng.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 4635
   ScaleWidth = 6795
   ShowInTaskbar = 0 'False
   Begin VB.CheckBox optOpen
      Alignment = 1 'Right Justify
      Caption = "Re-Open The Last Form When Starting The Section"
      Height = 255
      Left = 2400
      TabIndex = 4
      ToolTipText = "Workstation Setting - Where Possible, Re-Opens The Last Entry Form On Section Startup"
      Top = 2640
      Width = 4100
   End
   Begin VB.CheckBox optCaps
      Alignment = 1 'Right Justify
      Caption = "Turn Off Proper Casing Of Descriptions, Names, etc"
      Height = 255
      Left = 2400
      TabIndex = 3
      ToolTipText = "Turns Off The Auto Capitalization Feature"
      Top = 2280
      Width = 4100
   End
   Begin VB.CheckBox optSize
      Alignment = 1 'Right Justify
      Caption = "Form Resizing Is On"
      Height = 255
      Left = 2400
      TabIndex = 2
      ToolTipText = "When Checked Form Resizing Is On. Takes Affect After Restart."
      Top = 1920
      Width = 4100
   End
   Begin VB.CheckBox optTab
      Alignment = 1 'Right Justify
      Caption = "Enter Key Emulates Tab    "
      Height = 255
      Left = 2400
      TabIndex = 1
      ToolTipText = "Sets Enter Key To Act Like Tab (except multi-line Text Boxes)"
      Top = 1560
      Width = 4100
   End
   Begin VB.CheckBox optTips
      Alignment = 1 'Right Justify
      Caption = "Auto Tool Tips"
      Enabled = 0 'False
      Height = 375
      Left = 4920
      TabIndex = 7
      ToolTipText = "Show all tool tips"
      Top = 3840
      Value = 1 'Checked
      Width = 1575
   End
   Begin VB.Frame Frame2
      Caption = "Sections Tool Bar"
      Height = 615
      Left = 2400
      TabIndex = 6
      ToolTipText = "Show the Tool Bar across the top or on the side"
      Top = 3720
      Width = 2295
      Begin VB.OptionButton optSide
         Caption = "On Side"
         Height = 255
         Left = 120
         TabIndex = 17
         ToolTipText = "Sections Tool Bar on the side"
         Top = 240
         Value = -1 'True
         Width = 975
      End
      Begin VB.OptionButton optTop
         Caption = "On Top"
         Height = 255
         Left = 1200
         TabIndex = 16
         ToolTipText = "Sections Tool Bar on the top"
         Top = 240
         Width = 975
      End
   End
   Begin VB.Frame Frame1
      Caption = "Manager Bar Selection"
      Height = 615
      Left = 2400
      TabIndex = 5
      ToolTipText = "Effective Next Start - Double Click Manager Bar To Change"
      Top = 3000
      Width = 4095
      Begin VB.CheckBox optSve
         Caption = "Save Current"
         Height = 255
         Left = 2520
         TabIndex = 15
         ToolTipText = "Save Current Bar When Closing"
         Top = 240
         Value = 1 'Checked
         Width = 1335
      End
      Begin VB.OptionButton optVert
         Caption = "Vertical"
         Height = 255
         Left = 1440
         TabIndex = 14
         ToolTipText = "Open With A Vertical Bar"
         Top = 240
         Width = 1215
      End
      Begin VB.OptionButton optHorz
         Caption = "Horizontal"
         Height = 255
         Left = 120
         TabIndex = 13
         ToolTipText = "Open With Horizontal Bar"
         Top = 240
         Width = 1215
      End
   End
   Begin VB.CheckBox optMin
      Alignment = 1 'Right Justify
      Caption = "Minimize ESI2000 Manager On Selection? "
      Height = 255
      Left = 2400
      TabIndex = 0
      ToolTipText = "Move The Manager To The Task Bar After Program Selection"
      Top = 1200
      Width = 4100
   End
   Begin VB.FileListBox File1
      Enabled = 0 'False
      Height = 1845
      Left = 360
      Pattern = "*.rpt"
      TabIndex = 10
      TabStop = 0 'False
      Top = 2400
      Width = 1935
   End
   Begin VB.CommandButton cmdRept
      Caption = "&Reports"
      Height = 255
      Left = 2400
      TabIndex = 11
      TabStop = 0 'False
      Top = 720
      Width = 975
   End
   Begin VB.TextBox txtRept
      Height = 285
      Left = 3480
      TabIndex = 18
      Top = 720
      Width = 3015
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 435
      Left = 5880
      TabIndex = 12
      TabStop = 0 'False
      Top = 120
      Width = 875
   End
   Begin VB.DirListBox Dir1
      Height = 1665
      Left = 360
      TabIndex = 9
      TabStop = 0 'False
      Top = 720
      Width = 1935
   End
   Begin VB.DriveListBox Drive1
      Height = 315
      Left = 360
      TabIndex = 8
      TabStop = 0 'False
      Top = 360
      Width = 2175
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 19
      ToolTipText = "System Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "Diasetng.frx":08CA
      PictureDn = "Diasetng.frx":0A10
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 4920
      Top = 120
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 4635
      FormDesignWidth = 6795
   End
End
Attribute VB_Name = "diaSetng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Dim bOpenForm As Byte
Dim iShowVertical As Integer






Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs925"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdRept_Click()
   On Error Resume Next
   txtRept = Dir1 & "\"
   
End Sub

Private Sub Dir1_Change()
   On Error Resume Next
   File1 = Dir1
   
End Sub

Private Sub Dir1_Click()
   On Error Resume Next
   File1 = Dir1
   
End Sub


Private Sub Dir1_Scroll()
   On Error Resume Next
   File1 = Dir1
   
End Sub


Private Sub Drive1_Change()
   On Error Resume Next
   Dir1 = Drive1
   
End Sub

Private Sub Form_Activate()
   '
End Sub

Private Sub Form_Initialize()
   CloseForms
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   If iBarOnTop = 1 Then
      Move MdiSect.Left + 100, 0
   Else
      Move 0, 0
   End If
   optMin.Value = GetSetting("Esi2000", "mngr", "MinOnOpen", optMin.Value)
   If iBarOnTop = 1 Then optTop = True
   If iAutoTips = 1 Then optTips.Value = vbChecked Else optTips.Value = vbUnchecked
   
   iShowVertical = GetSetting("Esi2000", "mngr", "ShowVertical", iShowVertical)
   If iShowVertical = 1 Then optVert.Value = True Else optHorz.Value = True
   If bEnterAsTab Then optTab.Value = vbChecked Else optTab.Value = vbUnchecked
   If bResize Then optSize.Value = vbChecked Else optSize.Value = vbUnchecked
   bOpenForm = GetSetting("Esi2000", "System", "Reopenforms", bOpenForm)
   optOpen = bOpenForm
   optSve.Value = GetSetting("Esi2000", "mngr", "CurrentBar", optSve.Value)
   sReportPath = GetSetting("Esi2000", "System", "ReportPath", sReportPath)
   If sReportPath = "" Then sReportPath = App.Path & "\"
   txtRept = sReportPath
   Drive1 = "c:"
   'caps
   optCaps.Value = GetSetting("Esi2000", "mngr", "AutoCaps", optCaps.Value)
   
   Show
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Right(Trim(txtRept), 1) <> "\" Then txtRept = txtRept & "\"
   sReportPath = txtRept
   SaveSetting "Esi2000", "mngr", "ShowVertical", iShowVertical
   SaveSetting "Esi2000", "System", "ReportPath", sReportPath
   SaveSetting "Esi2000", "mngr", "CurrentBar", optSve.Value
   SaveSetting "Esi2000", "Programs", "BarOnTop", iBarOnTop
   SaveSetting "Esi2000", "System", "ResizeForm", bResize
   SaveSetting "Esi2000", "mngr", "AutoCaps", optCaps.Value
   SaveSetting "Esi2000", "System", "Reopenforms", optOpen.Value
   bAutoCaps = optCaps.Value
   If optTab.Value = vbUnchecked Then
      bEnterAsTab = False
      SaveSetting "Esi2000", "System", "EnterAsTab", "0"
   Else
      bEnterAsTab = True
      SaveSetting "Esi2000", "System", "EnterAsTab", "1"
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   WindowState = 1
   Set diaSetng = Nothing
   
End Sub





Private Sub optHorz_Click()
   If optHorz.Value = True Then iShowVertical = 0
   
End Sub

Private Sub optMin_Click()
   SaveSetting "Esi2000", "mngr", "MinOnOpen", optMin.Value
   
End Sub


Private Sub optMin_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optOpen_Click()
   bOpenForm = optOpen.Value
   
End Sub

Private Sub optSide_Click()
   If optSide.Value = True Then iBarOnTop = 0 Else iBarOnTop = 1
   If iBarOnTop = 1 Then
      MdiSect.Sidebar.Visible = False
      MdiSect.TopBar.Visible = True
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorBar").Caption = "Bar On Side"
      Move MdiSect.Left + 100, MdiSect.Top + 1280
   Else
      MdiSect.Sidebar.Visible = True
      MdiSect.TopBar.Visible = False
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorBar").Caption = "Bar On Top"
      Move MdiSect.Left + 2000, MdiSect.Top + 650
   End If
   
End Sub

Private Sub optSize_Click()
   bResize = optSize.Value
   
End Sub

Private Sub optSize_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optTab_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optTips_Click()
   If optTips.Value = vbChecked Then iAutoTips = 1 Else iAutoTips = 0
   If iAutoTips = 0 Then
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorTips").Caption = "Auto Tips On"
   Else
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorTips").Caption = "Auto Tips Off"
   End If
   SaveSetting "Esi2000", "Programs", "AutoTips", iAutoTips
   
End Sub


Private Sub optTop_Click()
   If optTop.Value = True Then iBarOnTop = 1 Else iBarOnTop = 0
   If iBarOnTop = 1 Then
      MdiSect.Sidebar.Visible = False
      MdiSect.TopBar.Visible = True
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorBar").Caption = "Bar On Side"
      Move MdiSect.Left + 100, MdiSect.Top + 1280
   Else
      MdiSect.Sidebar.Visible = True
      MdiSect.TopBar.Visible = False
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorBar").Caption = "Bar On Top"
      Move MdiSect.Left + 2000, MdiSect.Top + 650
   End If
   
End Sub


Private Sub optVert_Click()
   If optVert.Value = True Then iShowVertical = 1
   
End Sub

Private Sub txtRept_GotFocus()
   SelectFormat Me
   
End Sub

Private Sub txtRept_LostFocus()
   txtRept = Trim(txtRept)
   If Len(txtRept) > 0 Then
      If Right(txtRept, 1) <> "\" Then txtRept = txtRept & "\"
   End If
   sReportPath = txtRept
   
End Sub
