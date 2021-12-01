VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form tabLcst 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lot Costing"
   ClientHeight    =   4650
   ClientLeft      =   1845
   ClientTop       =   1605
   ClientWidth     =   5040
   Icon            =   "tabLcst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4650
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   4080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4650
      FormDesignWidth =   5040
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3780
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4050
      Width           =   1095
   End
   Begin TabDlg.SSTab tab1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   476
      TabCaption(0)   =   "&Edit      "
      TabPicture(0)   =   "tabLcst.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstEdt"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "zm(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&View    "
      TabPicture(1)   =   "tabLcst.frx":0440
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstVew"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "zm(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Functions"
      TabPicture(2)   =   "tabLcst.frx":0870
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "zm(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lstFun"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.ListBox lstEdt 
         Height          =   2985
         Left            =   -74520
         TabIndex        =   4
         Top             =   600
         Width           =   4065
      End
      Begin VB.ListBox lstVew 
         Height          =   2985
         Left            =   -74520
         TabIndex        =   3
         Top             =   600
         Width           =   4065
      End
      Begin VB.ListBox lstFun 
         Height          =   2985
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   4065
      End
      Begin VB.Label zm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No User Permissions"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   7
         Top             =   3600
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label zm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No User Permissions"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   6
         Top             =   3600
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label zm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No User Permissions"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   -74520
         TabIndex        =   5
         Top             =   3600
         Visible         =   0   'False
         Width           =   4095
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "Subject Help"
      Top             =   4080
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "tabLcst.frx":0CA5
      PictureDn       =   "tabLcst.frx":0DEB
   End
   Begin VB.Label lblcustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This Feature Is Not Available"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label zm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No User Permissions"
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "tabLcst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions


'************************************************************************************
'   tabLcst - Actual/Lot costing tab display menu.
'
'   Notes:
'
'
'   Created: (nth)
'   Revisions:
'
'************************************************************************************

'************************************************************************************

Option Explicit

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Selection Tab"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   cUR.CurrentGroup = "Lcst"
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
   
End Sub


Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   tab1.Tab = sActiveTab(7)
   If bSecSet = 1 Then
      '        User.Group8 = True
      If Secure.UserFinaG8E = 1 Then
         lstEdt.enabled = True
      Else
         lstEdt.enabled = False
         zm(0).Visible = True
      End If
      If Secure.UserFinaG8V = 1 Then
         lstVew.enabled = True
      Else
         lstVew.enabled = False
         zm(1).Visible = True
      End If
      If Secure.UserFinaG8F = 1 Then
         lstFun.enabled = True
      Else
         lstFun.enabled = False
         zm(2).Visible = True
      End If
   End If
   'New Code here *****:
   If bCustomerGroups(8) = 0 Then
      ' they didn't sign up, but we want them to see but not use
      lstEdt.enabled = False
      lstVew.enabled = False
      lstFun.enabled = False
      lblcustomer.ForeColor = ES_BLUE
      lblcustomer.Visible = True
      '            User.Group8 = 1
   End If
   'End New Code ******
   '    If User.Group8 Then
   'lstEdt.AddItem "No Installed Functions"
   lstVew.AddItem "Lots With QOH With No Cost"
   lstVew.AddItem "Lots Invoiced Not Costed"
   lstVew.AddItem "Cost Of Goods Sold - Lots"
   '
   '
   '        'lstFun.AddItem "No Installed Functions"
   '
   '    Else
   '        lstEdt.AddItem "No Group Permissions"
   '        tab1.enabled = False
   '    End If
   
End Sub


Private Sub Form_Resize()
   'Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set tabClos = Nothing
   
End Sub

Private Sub lstEdt_KeyPress(KeyAscii As Integer)
   Dim b As Byte
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   Select Case lstEdt.ListIndex
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
End Sub

Private Sub lstEdt_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, Y As Single)
   Dim b As Byte
   MouseCursor 13
   b = 1
   Select Case lstEdt.ListIndex
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
      
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub


Private Sub lstFun_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, Y As Single)
   Dim b As Byte
   MouseCursor 13
   b = 1
   Select Case lstFun.ListIndex
      
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
         diaWip.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub

Private Sub lstVew_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim b As Byte
   MouseCursor 13
   b = 1
   Select Case lstVew.ListIndex
      Case 0
         diaWip.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub


Private Sub tab1_Click(PreviousTab As Integer)
   sActiveTab(4) = tab1.Tab
   
End Sub

Private Sub tab1_GotFocus()
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
