VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form tabArec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accounts Receivable"
   ClientHeight    =   5520
   ClientLeft      =   1950
   ClientTop       =   750
   ClientWidth     =   4995
   Icon            =   "tabArec.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5520
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   5100
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5520
      FormDesignWidth =   4995
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3780
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5070
      Width           =   1095
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Subject Help"
      Top             =   5220
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
      PictureUp       =   "tabArec.frx":000C
      PictureDn       =   "tabArec.frx":0152
   End
   Begin TabDlg.SSTab tab1 
      Height          =   5000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8811
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   476
      TabCaption(0)   =   "&Edit      "
      TabPicture(0)   =   "tabArec.frx":0298
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "zm(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstEdt"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&View    "
      TabPicture(1)   =   "tabArec.frx":06CC
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "zm(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstVew"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Functions"
      TabPicture(2)   =   "tabArec.frx":0AFC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "zm(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lstFun"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.ListBox lstEdt 
         Height          =   4155
         ItemData        =   "tabArec.frx":0F31
         Left            =   -74520
         List            =   "tabArec.frx":0F33
         TabIndex        =   5
         Top             =   600
         Width           =   4065
      End
      Begin VB.ListBox lstVew 
         Height          =   4155
         ItemData        =   "tabArec.frx":0F35
         Left            =   480
         List            =   "tabArec.frx":0F37
         TabIndex        =   4
         Top             =   600
         Width           =   4065
      End
      Begin VB.ListBox lstFun 
         Height          =   4155
         ItemData        =   "tabArec.frx":0F39
         Left            =   -74520
         List            =   "tabArec.frx":0F3B
         TabIndex        =   3
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
         Left            =   -74520
         TabIndex        =   8
         Top             =   4740
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
         Left            =   480
         TabIndex        =   7
         Top             =   4740
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
         TabIndex        =   6
         Top             =   4740
         Visible         =   0   'False
         Width           =   4095
      End
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
      TabIndex        =   9
      Top             =   5100
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "tabArec"
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
   cUR.CurrentGroup = "Arec"
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   tab1.Tab = sActiveTab(0)
   If bSecSet = 1 Then
      '        User.Group1 = True
      If Secure.UserFinaG1E = 1 Then
         lstEdt.enabled = True
      Else
         lstEdt.enabled = False
         zm(0).Visible = True
      End If
      If Secure.UserFinaG1V = 1 Then
         lstVew.enabled = True
      Else
         lstVew.enabled = False
         zm(1).Visible = True
      End If
      If Secure.UserFinaG1F = 1 Then
         lstFun.enabled = True
      Else
         lstFun.enabled = False
         zm(2).Visible = True
      End If
      'New Code here *****:
      If bCustomerGroups(1) = 0 Then
         ' they didn't sign up, but we want them to see but not use
         lstEdt.enabled = False
         lstVew.enabled = False
         lstFun.enabled = False
         lblcustomer.ForeColor = ES_BLUE
         lblcustomer.Visible = True
         '            User.Group1 = 1
      End If
      'End New Code ******
   End If
   '    If User.Group1 Then
   lstEdt.AddItem "Customer Invoice (Sales Order)" '0
   lstEdt.AddItem "Customer Invoice (Packing Slip)" '1
   lstEdt.AddItem "Credit or Debit Memo" '2
   lstEdt.AddItem "Credit or Debit Memo Against Invoice" '3
   lstEdt.AddItem "Cash Receipts" '4
   lstEdt.AddItem "Customer Invoice Comments" '5
   lstEdt.AddItem "Customers" '6
   lstEdt.AddItem "Map ES/2000 Customers To QuickBooks ®" '7
   lstEdt.AddItem "Tax Codes" '8
   lstEdt.AddItem "Part B & O Tax Info" '9
   lstEdt.AddItem "Assign Customer Payers" '10
   
   lstVew.AddItem "Customer Invoices" '0
   lstVew.AddItem "Customer Statements" '1
   lstVew.AddItem "Customer Statements by PO" '2
   lstVew.AddItem "Customer Invoice Register" '3
   lstVew.AddItem "Cash Receipts Register" '4
   lstVew.AddItem "Accounts Receivable Aging (New)" '5
   lstVew.AddItem "Sales By GL Account" '6
   lstVew.AddItem "Unprinted Invoices" '7
   lstVew.AddItem "View A Cash Receipt" '8
   lstVew.AddItem "View Invoiced & Non-Invoiced Cash Receipts" '9
   lstVew.AddItem "QuickBooks ® Customer Equivalents" '10
   lstVew.AddItem "Tax Codes" '11
   lstVew.AddItem "Sales Tax Liability" '12
   lstVew.AddItem "B & O Tax Liability" '13
   lstVew.AddItem "Advance Payment Status" '14
   lstVew.AddItem "Customer Delivery Performance" '15
   lstVew.AddItem "Sources of Cash" '16
   lstVew.AddItem "Accounts Receivable Aging (Old)" '5
   
   lstFun.AddItem "Cancel An Invoice, Credit Or Debit Memo" '0
   lstFun.AddItem "Cancel A Cash Receipt" '1
   lstFun.AddItem "Update Sales Order Account Distributions" '2
   lstFun.AddItem "Auto Invoice Packing Slips" '3
   lstFun.AddItem "Update B & O Records" '4
   lstFun.AddItem "Export AR Invoice Activity To QuickBooks ® IIF" '5
   lstFun.AddItem "Create Invoice EDI File" '5
   lstFun.AddItem "Import Cash Receipt from Excel" '5
      '    Else
   '        lstEdt.AddItem "No Group Permissions"
   '        tab1.enabled = False
   '    End If
   
End Sub


Private Sub Form_Resize()
   'Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set tabArec = Nothing
   
End Sub


Private Sub lstEdt_KeyPress(KeyAscii As Integer)
   Dim b As Byte
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   Select Case lstEdt.ListIndex
      Case 0
         diaARe01a.Show
      Case 1
         diaARe02a.Show
      Case 2
         diaARe03a.Show
      Case 3
         diaARe06a.Show
      Case 4
         diaARe04a.Show
      Case 5
         diaARe05a.Show
      Case 6
         diaCcust.Show
      Case 7
         diaARe09a.Show
      Case 8
         diaARe10a.Show
      Case 9
         diaARe11a.Show
      Case 10
         diaARe12a.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub


Private Sub lstEdt_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   
   lstEdt_KeyPress 13
   
End Sub

Private Sub lstFun_KeyPress(KeyAscii As Integer)
   Dim b As Byte
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   Select Case lstFun.ListIndex
      Case 0
         diaARf01a.Show
      Case 1
         diaARf02a.Show
      Case 2
         diaARf04a.Show
      Case 3
         diaARf05a.Show
      Case 4
         diaARf07a.Show
      Case 5
         diaARf08a.Show
      Case 6
         PackPSf09a.Show
      Case 7
         diaARf09a.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub

Private Sub lstFun_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   
   lstFun_KeyPress 13
   
End Sub

Private Sub lstVew_KeyPress(KeyAscii As Integer)
   Dim b As Byte
   If KeyAscii <> 13 Then Exit Sub
   '    MouseCursor 13
   b = 1
   Select Case lstVew.ListIndex
      Case 0
         diaARp01a.Show
      Case 1
         diaARp02a.cbByPO.Value = vbUnchecked
         diaARp02a.Show
      Case 2
         diaARp02a.cbByPO.Value = vbChecked
         diaARp02a.Show
      Case 3
         diaARp03a.Show
      Case 4
         diaARp04a.Show
      Case 5
         ArAging.Show    ' AR AGING (NEW)
      Case 6
         diaARp07a.Show
      Case 7
         diaARp08a.Show
      Case 8
         diaARp09a.Show
      Case 9
         diaARp09b.Show
      Case 10
         diaARp12a.Show
      Case 11
         diaARp13a.Show
      Case 12
         diaARp11a.Show
      Case 13
         diaARp15a.Show
      Case 14 'Advance payment status - what is it?
         diaARp14a.Show
      Case 15
         diaARp16a.Show
         
         'Sources of Cash
      Case 16
         diaARp17a.Show
      Case 17
         diaARp05a.Show    ' ar aging old
         
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub

Private Sub lstVew_MouseDown(Button As Integer, Shift As Integer _
                             , X As Single, Y As Single)
   
   lstVew_KeyPress 13
   
   '    Dim b As Byte
   '    MouseCursor 13
   '    b = 1
   '    Select Case lstVew.ListIndex
   '        Case 0
   '            diaARp01a.Show
   '        Case 1
   '            diaARp02a.Show
   '        Case 2
   '            diaARp03a.Show
   '        Case 3
   '            diaARp04a.Show
   '        Case 4
   '            diaARp05a.Show
   '        Case 5
   '            diaARp07a.Show
   '        Case 6
   '            diaARp08a.Show
   '        Case 7
   '            diaARp09a.Show
   '        Case 8
   '            diaARp12a.Show
   '        Case 9
   '            diaARp13a.Show
   '        Case 10
   '            diaARp11a.Show
   '        Case 11
   '            diaARp15a.Show
   '        Case 12
   '            diaARp14a.Show
   '        Case Else
   '            b = 0
   '    End Select
   '    If b = 1 Then Hide Else MouseCursor 0
End Sub

Private Sub tab1_Click(PreviousTab As Integer)
   sActiveTab(0) = tab1.Tab
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
