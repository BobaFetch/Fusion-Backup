VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form tabTime
   BorderStyle = 3 'Fixed Dialog
   Caption = "Time Charges"
   ClientHeight = 4545
   ClientLeft = 2535
   ClientTop = 1605
   ClientWidth = 4980
   Icon = "tabTime.frx":0000
   LinkTopic = "Form1"
   LockControls = -1 'True
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 4545
   ScaleWidth = 4980
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 375
      Left = 3780
      TabIndex = 1
      TabStop = 0 'False
      Top = 4050
      Width = 1095
   End
   Begin TabDlg.SSTab tab1
      Height = 3975
      Left = 0
      TabIndex = 0
      Top = 0
      Width = 4935
      _ExtentX = 8705
      _ExtentY = 7011
      _Version = 393216
      Style = 1
      TabHeight = 476
      TabCaption(0) = "&Edit      "
      TabPicture(0) = "tabTime.frx":030A
      Tab(0).ControlEnabled = -1 'True
      Tab(0).Control(0) = "zm(0)"
      Tab(0).Control(0).Enabled = 0 'False
      Tab(0).Control(1) = "lstEdt"
      Tab(0).Control(1).Enabled = 0 'False
      Tab(0).ControlCount = 2
      TabCaption(1) = "&View    "
      TabPicture(1) = "tabTime.frx":0326
      Tab(1).ControlEnabled = 0 'False
      Tab(1).Control(0) = "lstVew"
      Tab(1).Control(1) = "zm(1)"
      Tab(1).ControlCount = 2
      TabCaption(2) = "&Functions"
      TabPicture(2) = "tabTime.frx":0342
      Tab(2).ControlEnabled = 0 'False
      Tab(2).Control(0) = "lstFun"
      Tab(2).Control(1) = "zm(2)"
      Tab(2).ControlCount = 2
      Begin VB.ListBox lstEdt
         Height = 2985
         ItemData = "tabTime.frx":035E
         Left = 480
         List = "tabTime.frx":0360
         TabIndex = 4
         Top = 600
         Width = 4065
      End
      Begin VB.ListBox lstVew
         Height = 2985
         ItemData = "tabTime.frx":0362
         Left = -74520
         List = "tabTime.frx":0364
         TabIndex = 3
         Top = 600
         Width = 4065
      End
      Begin VB.ListBox lstFun
         Height = 2985
         ItemData = "tabTime.frx":0366
         Left = -74520
         List = "tabTime.frx":0368
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
         TabIndex = 8
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
         Index = 0
         Left = 480
         TabIndex = 6
         Top = 3600
         Visible = 0 'False
         Width = 4095
      End
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 5
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
      PictureUp = "tabTime.frx":036A
      PictureDn = "tabTime.frx":04B0
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 480
      Top = 4005
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 4545
      FormDesignWidth = 4980
   End
   Begin VB.Label lblCustomer
      Alignment = 2 'Center
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Caption = "This Feature Is Not Available"
      ForeColor = &H80000008&
      Height = 255
      Left = 840
      TabIndex = 9
      Top = 4080
      Visible = 0 'False
      Width = 2895
   End
End
Attribute VB_Name = "tabTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions

Dim b As Byte

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs923.htm"
      MouseCursor False
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MouseCursor 0
   cur.CurrentGroup = "Time"
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
   
End Sub


Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   lstEdt.AddItem "Enter/Revise Daily Time Charges "
   lstEdt.AddItem "Employees "
   lstEdt.AddItem "Time Type Codes "
   
   lstVew.AddItem "Employees By Name "
   lstVew.AddItem "Employees By Number "
   lstVew.AddItem "Daily Employee Time Charges "
   lstVew.AddItem "Weekly Time Charges "
   lstVew.AddItem "Time Type Codes "
   lstVew.AddItem "Open Point Of Manufacturing Logins "
   
   lstFun.AddItem "Delete A Daily Time Charge"
   lstFun.AddItem "Revise Time Charge Pay Rates"
End Sub


Private Sub Form_Resize()
   'Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set tabTime = Nothing
   
End Sub


Private Sub lstEdt_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   Select Case lstEdt.ListIndex
      Case 0
         diaHrtme.Show
      Case 1
         diaHempl.Show
      Case 2
         diaHcode.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub


Private Sub lstEdt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '    MouseCursor 13
   '    b = 1
   '    Select Case lstEdt.ListIndex
   '        Case 0
   '            diaHrtme.Show
   '        Case 1
   '            diaHempl.Show
   '        Case 2
   '            diaHcode.Show
   '        Case Else
   '            b = 0
   '    End Select
   '    If b = 1 Then Hide Else MouseCursor 0
   
   'treat same as carriage return from selected item
   lstEdt_KeyPress Asc(vbCr)
End Sub

Private Sub lstFun_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   Select Case lstFun.ListIndex
      Case 0
         diaHdlch.Show 'Delete A Daily Time Charge
      Case 1
         ReviseRates.Show 'Revise Time Charge Pay Rates
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub


Private Sub lstFun_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '    MouseCursor 13
   '    b = 1
   '    Select Case lstFun.ListIndex
   '        Case 0
   '            diaHdlch.Show       'delete a daily time charge
   '        Case 1
   '            ReviseRates.Show    'revise time charge pay rates
   '        Case Else
   '            b = 0
   '    End Select
   '    If b = 1 Then Hide Else MouseCursor 0
   
   'treat same as carriage return from selected item
   lstFun_KeyPress Asc(vbCr)
   
End Sub


Private Sub lstVew_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   Select Case lstVew.ListIndex
      Case 0
         diaPhu01.Show
      Case 1
         diaPhu02.Show
      Case 2
         diaPhu03.Show
      Case 3
         diaPhu05.Show
      Case 4
         diaPhu15.Show
      Case 5
         diaPhu16.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub


Private Sub lstVew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '    Dim b As Byte
   '    MouseCursor 13
   '    b = 1
   '    Select Case lstVew.ListIndex
   '        Case 0
   '            diaPhu01.Show
   '        Case 1
   '            diaPhu02.Show
   '        Case 2
   '            diaPhu03.Show
   '        Case 3
   '            diaPhu05.Show
   '        Case 4
   '            diaPhu15.Show
   '        Case Else
   '            b = 0
   '    End Select
   '    If b = 1 Then Hide Else MouseCursor 0
   
   'treat same as carriage return from selected item
   lstVew_KeyPress Asc(vbCr)
End Sub


Private Sub tab1_Click(PreviousTab As Integer)
   'sActiveTab(4) = tab1.Tab
   
End Sub

Private Sub tab1_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub
