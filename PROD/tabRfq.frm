VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form tabRfq
   BorderStyle = 3 'Fixed Dialog
   Caption = "Request For Quote"
   ClientHeight = 4548
   ClientLeft = 1848
   ClientTop = 1608
   ClientWidth = 4944
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 4548
   ScaleWidth = 4944
   ShowInTaskbar = 0 'False
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 360
      Top = 4080
      _Version = 196615
      _ExtentX = 593
      _ExtentY = 593
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 4548
      FormDesignWidth = 4944
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 375
      Left = 3780
      TabIndex = 1
      TabStop = 0 'False
      Top = 4080
      Width = 1095
   End
   Begin TabDlg.SSTab tab1
      Height = 3975
      Left = 0
      TabIndex = 0
      Top = 0
      Width = 4935
      _ExtentX = 8700
      _ExtentY = 7006
      _Version = 393216
      Style = 1
      TabHeight = 476
      TabCaption(0) = "&Edit      "
      TabPicture(0) = "tabRfq.frx":0000
      Tab(0).ControlEnabled = -1 'True
      Tab(0).Control(0) = "zm(0)"
      Tab(0).Control(0).Enabled = 0 'False
      Tab(0).Control(1) = "lstEdt"
      Tab(0).Control(1).Enabled = 0 'False
      Tab(0).ControlCount = 2
      TabCaption(1) = "&View    "
      Tab(1).ControlEnabled = 0 'False
      Tab(1).Control(0) = "zm(1)"
      Tab(1).Control(1) = "lstVew"
      Tab(1).ControlCount = 2
      TabCaption(2) = "&Functions"
      Tab(2).ControlEnabled = 0 'False
      Tab(2).Control(0) = "lstFun"
      Tab(2).Control(1) = "zm(2)"
      Tab(2).ControlCount = 2
      Begin VB.ListBox lstEdt
         Height = 2928
         Left = 480
         TabIndex = 4
         Top = 600
         Width = 4065
      End
      Begin VB.ListBox lstVew
         Height = 2928
         Left = -74520
         TabIndex = 3
         Top = 600
         Width = 4065
      End
      Begin VB.ListBox lstFun
         Height = 2928
         Left = -74520
         TabIndex = 2
         Top = 600
         Width = 4065
      End
      Begin VB.Label zm
         Alignment = 2 'Center
         BackStyle = 0 'Transparent
         Caption = "No User Permissions"
         ForeColor = &H000000C0&
         Height = 255
         Index = 2
         Left = -74520
         TabIndex = 7
         Top = 3600
         Visible = 0 'False
         Width = 4095
      End
      Begin VB.Label zm
         Alignment = 2 'Center
         BackStyle = 0 'Transparent
         Caption = "No User Permissions"
         ForeColor = &H000000C0&
         Height = 255
         Index = 1
         Left = -74520
         TabIndex = 6
         Top = 3600
         Visible = 0 'False
         Width = 4095
      End
      Begin VB.Label zm
         Alignment = 2 'Center
         BackStyle = 0 'Transparent
         Caption = "No User Permissions"
         ForeColor = &H000000C0&
         Height = 255
         Index = 0
         Left = 480
         TabIndex = 5
         Top = 3600
         Visible = 0 'False
         Width = 4095
      End
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 9
      ToolTipText = "Subject Help"
      Top = 4080
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "tabRfq.frx":001C
      PictureDn = "tabRfq.frx":0162
   End
   Begin VB.Label zm
      Alignment = 2 'Center
      BackStyle = 0 'Transparent
      Caption = "No User Permissions"
      ForeColor = &H000000C0&
      Height = 255
      Index = 3
      Left = 0
      TabIndex = 8
      Top = 0
      Visible = 0 'False
      Width = 4095
   End
End
Attribute VB_Name = "tabRfq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'************************************************************************************
'   tabRfq - Request For quote Menu
'
'   Notes:
'
'
'   Created: (JcW)
'   Revisions:
'
'************************************************************************************

'************************************************************************************

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 923
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   cUR.CurrentGroup = "RFQS"
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
End Sub


Private Sub Form_Load()
   SetDiaPos Me, ES_DONTLIST, ES_RESIZE
   tab1.Tab = sActiveTab(6)
   If bSecSet = 1 Then
      User.Group6 = True
      If Secure.UserProdG6E = 1 Then
         lstEdt.Enabled = True
      Else
         lstEdt.Enabled = False
         zm(0).Visible = True
      End If
      If Secure.UserProdG6V = 1 Then
         lstVew.Enabled = True
      Else
         lstVew.Enabled = False
         zm(1).Visible = True
      End If
      If Secure.UserProdG6F = 1 Then
         lstFun.Enabled = True
      Else
         lstFun.Enabled = False
         zm(2).Visible = True
      End If
   End If
   If User.Group6 Then
      lstEdt.AddItem "Request For Quote(RFQ)"
      lstEdt.AddItem "Vendor RFQ Response"
      
      lstVew.AddItem "Request For Quote(RFQ)"
      lstVew.AddItem "RFQ's By Vendor"
      lstVew.AddItem "Lowest RFQ Item Cost"
      lstVew.AddItem "Lowest RFQ Total Cost"
      
      lstFun.AddItem "Cancel An RFQ"
      lstFun.AddItem "Create PO's (RFQ's)"
      lstFun.AddItem "Delete An RFQ"
      lstFun.AddItem "Uncancel An RFQ"
      
   Else
      lstEdt.AddItem "No Group Permissions"
      tab1.Enabled = False
   End If
   
End Sub

Private Sub Form_Resize()
   'Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set tabRfq = Nothing
End Sub

Private Sub lstEdt_KeyPress(KeyAscii As Integer)
   Dim b As Byte
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   Select Case lstEdt.ListIndex
      Case 0
         diaRFe01a.Show
      Case 1
         diaRFe02a.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
End Sub

Private Sub lstEdt_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   Dim b As Byte
   MouseCursor 13
   b = 1
   Select Case lstEdt.ListIndex
      Case 0
         diaRFe01a.Show
      Case 1
         diaRFe02a.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub

Private Sub lstFun_KeyPress(KeyAscii As Integer)
   Dim b As Byte
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   Select Case lstFun.ListIndex
      Case 0
         diaRFf01a.Show
      Case 1
         diaRFf02a.Show
      Case 2
         diaRFf03a.Show
      Case 3
         diaRFf04a.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
End Sub


Private Sub lstFun_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   Dim b As Byte
   MouseCursor 13
   b = 1
   Select Case lstFun.ListIndex
      Case 0
         diaRFf01a.Show
      Case 1
         diaRFf02a.Show
      Case 2
         diaRFf03a.Show
      Case 3
         diaRFf04a.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
End Sub


Private Sub lstVew_KeyPress(KeyAscii As Integer)
   Dim b As Byte
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   Select Case lstVew.ListIndex
      Case 0
         diaRFp01a.Show
      Case 1
         diaRFp02a.Show
      Case 2
         diaRFp03a.Show
      Case 3
         diaRFp04a.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
End Sub

Private Sub lstVew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim b As Byte
   MouseCursor 13
   b = 1
   Select Case lstVew.ListIndex
      Case 0
         diaRFp01a.Show
      Case 1
         diaRFp02a.Show
      Case 2
         diaRFp03a.Show
      Case 3
         diaRFp04a.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
End Sub


Private Sub tab1_Click(PreviousTab As Integer)
   sActiveTab(4) = tab1.Tab
   
End Sub

Private Sub Tab1_GotFocus()
   Select Case tab1.Tab
      Case 0
         lstEdt.TabIndex = 1
      Case 1
         lstVew.TabIndex = 1
      Case 2
         lstFun.TabIndex = 1
   End Select
   
End Sub


Private Sub tab1_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub
