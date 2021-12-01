VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form MatlMMe06a
   BorderStyle = 3 'Fixed Dialog
   Caption = "Workstation Lot Locations"
   ClientHeight = 3084
   ClientLeft = 48
   ClientTop = 336
   ClientWidth = 6000
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H8000000F&
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 3084
   ScaleWidth = 6000
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdHlp
      Appearance = 0 'Flat
      Height = 250
      Left = 0
      Picture = "MatlMMe06a.frx":0000
      Style = 1 'Graphical
      TabIndex = 13
      TabStop = 0 'False
      ToolTipText = "Subject Help"
      Top = 0
      UseMaskColor = -1 'True
      Width = 250
   End
   Begin VB.CommandButton cmdUpd
      Cancel = -1 'True
      Caption = "Apply"
      Height = 315
      Left = 5040
      TabIndex = 12
      TabStop = 0 'False
      ToolTipText = "Update The Settings"
      Top = 1080
      Width = 875
   End
   Begin VB.TextBox txtLoc4
      Height = 285
      Left = 3480
      TabIndex = 11
      Tag = "3"
      ToolTipText = "Location Alllowed From This WorkStation"
      Top = 2520
      Width = 615
   End
   Begin VB.TextBox txtLoc3
      Height = 285
      Left = 3480
      TabIndex = 10
      Tag = "3"
      ToolTipText = "Location Alllowed From This WorkStation"
      Top = 2160
      Width = 615
   End
   Begin VB.TextBox txtLoc2
      Height = 285
      Left = 3480
      TabIndex = 9
      Tag = "3"
      ToolTipText = "Location Alllowed From This WorkStation"
      Top = 1800
      Width = 615
   End
   Begin VB.CheckBox optOn
      Alignment = 1 'Right Justify
      Caption = "Turn On Lot Location Checking"
      Height = 255
      Left = 550
      TabIndex = 0
      ToolTipText = "Initialized (Turns On Or Off) Lot Location Verification For This Workstation"
      Top = 1080
      Width = 3550
   End
   Begin VB.TextBox txtLoc1
      Height = 285
      Left = 3480
      TabIndex = 1
      Tag = "3"
      ToolTipText = "Location Alllowed From This WorkStation"
      Top = 1440
      Width = 615
   End
   Begin VB.CommandButton cmdCan
      Caption = "Close"
      Height = 435
      Left = 5040
      TabIndex = 2
      TabStop = 0 'False
      Top = 0
      Width = 875
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 6600
      Top = 4200
      _Version = 196615
      _ExtentX = 593
      _ExtentY = 593
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3084
      FormDesignWidth = 6000
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Location 4"
      ForeColor = &H00400000&
      Height = 255
      Index = 4
      Left = 600
      TabIndex = 8
      Top = 2520
      Width = 4215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Location 3"
      ForeColor = &H00400000&
      Height = 255
      Index = 3
      Left = 600
      TabIndex = 7
      Top = 2160
      Width = 4215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Location 2"
      ForeColor = &H00400000&
      Height = 255
      Index = 1
      Left = 600
      TabIndex = 6
      Top = 1800
      Width = 4215
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Location 1"
      ForeColor = &H00400000&
      Height = 255
      Index = 0
      Left = 600
      TabIndex = 5
      Top = 1440
      Width = 4215
   End
   Begin VB.Label Workstation
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 600
      TabIndex = 4
      Top = 720
      Width = 2415
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Sets Lot Locations Allowed For This Workstation:"
      ForeColor = &H00400000&
      Height = 255
      Index = 2
      Left = 600
      TabIndex = 3
      Top = 360
      Width = 4215
   End
End
Attribute VB_Name = "MatlMMe06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'6/1/05 New - Primary customer; INTCOA to control locations
Option Explicit
Dim bOnLoad As Byte
Dim bChanged As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Dim bResponse As Byte
   If txtLoc1 = "" And txtLoc2 = "" And txtLoc3 = "" And txtLoc4 = "" Then
      optOn.Value = vbUnchecked
   End If
   If bChanged = 0 Then
      bResponse = vbYes
   Else
      bResponse = MsgBox("The Data Has Changed. Exit Without Saving?", _
                  ES_NOQUESTION, Caption)
   End If
   If bResponse = vbYes Then Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 1601
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdUpd_Click()
   SaveSettings
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      Workstation = GetWorkStation()
      GetSettings
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set MatlMMe06a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   bChanged = 0
   
End Sub

Private Sub GetSettings()
   optOn.Value = GetSetting("Esi2000", "WSLotLocs", "On", optOn.Value)
   txtLoc1.Text = GetSetting("Esi2000", "WSLotLocs", "Loc1", Trim(txtLoc1.Text))
   txtLoc2.Text = GetSetting("Esi2000", "WSLotLocs", "Loc2", Trim(txtLoc2.Text))
   txtLoc3.Text = GetSetting("Esi2000", "WSLotLocs", "Loc3", Trim(txtLoc3.Text))
   txtLoc4.Text = GetSetting("Esi2000", "WSLotLocs", "Loc4", Trim(txtLoc4.Text))
   
End Sub




Private Sub SaveSettings()
   If txtLoc1 = "" And txtLoc2 = "" And txtLoc3 = "" And txtLoc4 = "" Then
      MsgBox "Requires At Least One Location.", _
         vbInformation, Caption
      optOn.Value = vbUnchecked
   End If
   SaveSetting "Esi2000", "WSLotLocs", "On", optOn.Value
   If optOn.Value = vbUnchecked Then
      txtLoc1 = ""
      txtLoc2 = ""
      txtLoc3 = ""
      txtLoc4 = ""
   End If
   SaveSetting "Esi2000", "WSLotLocs", "Loc1", txtLoc1.Text
   SaveSetting "Esi2000", "WSLotLocs", "Loc2", txtLoc2.Text
   SaveSetting "Esi2000", "WSLotLocs", "Loc3", txtLoc3.Text
   SaveSetting "Esi2000", "WSLotLocs", "Loc4", txtLoc4.Text
   SysMsg "Selections Where Saved.", True
   bChanged = 0
   
End Sub

Private Sub optOn_Click()
   If Not bOnLoad Then SaveSetting "Esi2000", "WSLotLocs", "On", optOn.Value
   
End Sub


Private Sub txtLoc1_Change()
   If bOnLoad = 0 Then bChanged = 1
   
End Sub


Private Sub txtLoc1_LostFocus()
   txtLoc1 = CheckLen(txtLoc1, 4)
   
End Sub


Private Sub txtLoc2_Change()
   If bOnLoad = 0 Then bChanged = 1
   
End Sub


Private Sub txtLoc2_LostFocus()
   txtLoc2 = CheckLen(txtLoc2, 4)
   
End Sub


Private Sub txtLoc3_Change()
   If bOnLoad = 0 Then bChanged = 1
   
End Sub


Private Sub txtLoc3_LostFocus()
   txtLoc3 = CheckLen(txtLoc3, 4)
   
End Sub


Private Sub txtLoc4_Change()
   If bOnLoad = 0 Then bChanged = 1
   
End Sub


Private Sub txtLoc4_LostFocus()
   txtLoc4 = CheckLen(txtLoc4, 4)
   
End Sub
