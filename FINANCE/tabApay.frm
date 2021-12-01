VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form tabApay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accounts Payable"
   ClientHeight    =   5520
   ClientLeft      =   1845
   ClientTop       =   1605
   ClientWidth     =   4995
   Icon            =   "tabApay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5520
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   180
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
      Left            =   3720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5070
      Width           =   1095
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   -60
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
      PictureUp       =   "tabApay.frx":000C
      PictureDn       =   "tabApay.frx":0152
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
      TabPicture(0)   =   "tabApay.frx":0298
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstEdt"
      Tab(0).Control(1)=   "zm(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&View    "
      TabPicture(1)   =   "tabApay.frx":06CC
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "zm(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstVew"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Functions"
      TabPicture(2)   =   "tabApay.frx":0AFC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "zm(2)"
      Tab(2).Control(1)=   "lstFun"
      Tab(2).ControlCount=   2
      Begin VB.ListBox lstEdt 
         Height          =   4155
         ItemData        =   "tabApay.frx":0F31
         Left            =   -74520
         List            =   "tabApay.frx":0F33
         TabIndex        =   5
         Top             =   600
         Width           =   4065
      End
      Begin VB.ListBox lstVew 
         Height          =   4155
         ItemData        =   "tabApay.frx":0F35
         Left            =   480
         List            =   "tabApay.frx":0F37
         TabIndex        =   4
         Top             =   600
         Width           =   4065
      End
      Begin VB.ListBox lstFun 
         Height          =   4155
         ItemData        =   "tabApay.frx":0F39
         Left            =   -74520
         List            =   "tabApay.frx":0F3B
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
         Left            =   -74580
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
      Left            =   780
      TabIndex        =   10
      Top             =   5100
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
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "tabApay"
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
   cUR.CurrentGroup = "Apay"
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   tab1.Tab = sActiveTab(1)
   
   If bSecSet = 1 Then
      '        User.Group2 = True
      If Secure.UserFinaG2E = 1 Then
         lstEdt.enabled = True
      Else
         lstEdt.enabled = False
         zm(0).Visible = True
      End If
      If Secure.UserFinaG2V = 1 Then
         lstVew.enabled = True
      Else
         lstVew.enabled = False
         zm(1).Visible = True
      End If
      If Secure.UserFinaG2F = 1 Then
         lstFun.enabled = True
      Else
         lstFun.enabled = False
         zm(2).Visible = True
      End If
      'New Code here *****:
      If bCustomerGroups(2) = 0 Then
         ' they didn't sign up, but we want them to see but not use
         lstEdt.enabled = False
         lstVew.enabled = False
         lstFun.enabled = False
         lblcustomer.ForeColor = ES_BLUE
         lblcustomer.Visible = True
         '            User.Group2 = 1
      End If
      'End New Code ******
   End If
   '    If User.Group2 Then
   lstEdt.AddItem "Vendor Invoice" '0
   lstEdt.AddItem "Vendor Credit Or Debit Memo" '1
   lstEdt.AddItem "Cash Disbursements (Pay Invoices)" '2
   lstEdt.AddItem "External Check (No Invoice)" '3
   lstEdt.AddItem "Revise Invoice Due Dates/Comments" '4
   lstEdt.AddItem "Revise Invoice GL Distribution" '5
   lstEdt.AddItem "Vendors" '6
   lstEdt.AddItem "Computer Check Setup" '7
   lstEdt.AddItem "Edit Check Memos" '8
   
   lstVew.AddItem "Vendor Invoices" '0
   lstVew.AddItem "Vendor Statements" '1
   lstVew.AddItem "Average Age Of Paid Invoices" '2
   lstVew.AddItem "Vendor Invoice Register" '3
   lstVew.AddItem "Accounts Payable Aging" '4
   lstVew.AddItem "Purchases By GL Account" '5
   lstVew.AddItem "Material Purchase Price Variance" '6
   lstVew.AddItem "Received And Not Invoiced By Vendor" '7
'   lstVew.AddItem "Cash Requirements" '8            'never implemented
'   lstVew.AddItem "Lost Discounts" '9      'never implemented
   lstVew.AddItem "Check Setup" '10
   lstVew.AddItem "Computer Check Summary" '11
   lstVew.AddItem "Cleared Checks" '12
   lstVew.AddItem "Check Analysis" '13
   lstVew.AddItem "Vendor Delivery Performance" '14
   lstVew.AddItem "View a Cash Disbursement" '15
   lstVew.AddItem "Uses of Cash" '16
   lstVew.AddItem "Vendor Invoice Profile" '17
   lstVew.AddItem "Accounts Payable Aging (Old)" '18
   lstVew.AddItem "Vendor Statements (Old)" '19
   
   lstFun.AddItem "Cancel An AP Invoice" '0
   lstFun.AddItem "Print Computer Checks" '1
   lstFun.AddItem "Void Checks" '2
   lstFun.AddItem "Reprint Checks" '3
   lstFun.AddItem "Export AP Invoice Activity To QuickBooks ® IIF" '4
   lstFun.AddItem "1099's" '5
   lstFun.AddItem "Rename CheckNumber" '6
   'lstFun.AddItem "Clear AP Aging Invoices" '4 ' per IMAINC, no longer allowed
   '    Else
   '        lstEdt.AddItem "No Group Permissions"
   '        tab1.enabled = False
   '    End If
   
End Sub


Private Sub Form_Resize()
   'Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set tabApay = Nothing
   
End Sub


Private Sub lstEdt_KeyPress(KeyAscii As Integer)
   Dim b As Byte
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   Select Case lstEdt.ListIndex
      Case 0
         diaAPe01a.Show
      Case 1
         diaAPe02a.Show
      Case 2
         diaAPe03a.Show
      Case 3
         diaAPe11a.Show
      Case 4
         diaAPe04a.Show
      Case 5
         diaAPe05a.Show
      Case 6
         VendorEdit01.Tag = 2   'Set the calling module tag
         VendorEdit01.Show      'now show the new form
      Case 7
         diaAPe08a.Show
      Case 8
         diaAPe10a.Show
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
         diaAPf01a.Show
      Case 1
         diaAPf03a.Show
      Case 2
         diaAPf04a.Show
      Case 3
         diaAPf05a.Show
      Case 4
         diaAPf06a.Show
      Case 5
         diaAPf07a.Show
      Case 6
         diaAPf08a.Show
'      Case 7
'         diaAPf09a.Show
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
   'MouseCursor 13
   b = 1
   Select Case lstVew.ListIndex
      Case 0
         diaAPp01a.Show
      Case 1
         diaAPp02b.Show
      Case 2
         diaAPp06a.Show
      Case 3
         diaAPp07a.Show
      Case 4
         diaAPp08b.Show
      Case 5
         diaAPp10a.Show
      Case 6
         diaCLp01a.Show
      Case 7
         diaAPp11a.Show
      Case 8        'old 10
         diaAPp14a.Show
      Case 9       'old 11
         diaAPp15a.Show
      Case 10       'old 12
         diaAPp16a.Show
      Case 11       'old 13
         diaAPp19a.Show
         
         'vendor delivery performance
      Case 12       'old 14
         diaAPp23a.Show
         
         'view a cash disbursement
      Case 13       'old 15
         diaAPp24a.Show
         
         'Uses of Cash
      Case 14       'old 16
         diaAPp25a.Show
      Case 15       'old 17
         diaAPp26a.Show
         
      Case 16
         diaAPp08a.Show    ' old ap aging reports
         
      Case 17
         diaAPp02a.Show    ' old vendor statement
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub

Private Sub lstVew_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   
   lstVew_KeyPress 13
   
   '    Dim b As Byte
   '    MouseCursor 13
   '    b = 1
   '    Select Case lstVew.ListIndex
   '        Case 0
   '            diaAPp01a.Show
   '        Case 1
   '            diaAPp02a.Show
   '        Case 2
   '            diaAPp06a.Show
   '        Case 3
   '            diaAPp07a.Show
   '        Case 4
   '            diaAPp08a.Show
   '        Case 5
   '            diaAPp10a.Show
   '        Case 6
   '            diaCLp01a.Show
   '        Case 7
   '            diaAPp11a.Show
   '        Case 10
   '            diaAPp14a.Show
   '        Case 11
   '            diaAPp15a.Show
   '        Case 12
   '            diaAPp16a.Show
   '        Case 13
   '            diaAPp19a.Show
   '        Case Else
   '            b = 0
   '    End Select
   '    If b = 1 Then Hide Else MouseCursor 0
   '
End Sub


Private Sub tab1_Click(PreviousTab As Integer)
   sActiveTab(1) = tab1.Tab
   
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
