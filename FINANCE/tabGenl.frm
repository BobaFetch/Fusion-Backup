VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form tabGenl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "General Ledger"
   ClientHeight    =   5520
   ClientLeft      =   1845
   ClientTop       =   1605
   ClientWidth     =   4995
   Icon            =   "tabGenl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5520
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   420
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
      Left            =   3840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5070
      Width           =   1095
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   60
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
      PictureUp       =   "tabGenl.frx":000C
      PictureDn       =   "tabGenl.frx":0152
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
      Tab             =   2
      TabHeight       =   476
      TabCaption(0)   =   "&Edit      "
      TabPicture(0)   =   "tabGenl.frx":0298
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "zm(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstEdt"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&View    "
      TabPicture(1)   =   "tabGenl.frx":06CC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "zm(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstVew"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Functions"
      TabPicture(2)   =   "tabGenl.frx":0AFC
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "zm(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lstFun"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.ListBox lstEdt 
         Height          =   3960
         ItemData        =   "tabGenl.frx":0F31
         Left            =   -74520
         List            =   "tabGenl.frx":0F33
         TabIndex        =   5
         Top             =   600
         Width           =   4065
      End
      Begin VB.ListBox lstVew 
         Height          =   3960
         ItemData        =   "tabGenl.frx":0F35
         Left            =   -74520
         List            =   "tabGenl.frx":0F37
         TabIndex        =   4
         Top             =   600
         Width           =   4065
      End
      Begin VB.ListBox lstFun 
         Height          =   3960
         ItemData        =   "tabGenl.frx":0F39
         Left            =   480
         List            =   "tabGenl.frx":0F3B
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
         Left            =   480
         TabIndex        =   8
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
         Left            =   -74580
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
         Top             =   3600
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
      Left            =   900
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
Attribute VB_Name = "tabGenl"
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
   cUR.CurrentGroup = "Genl"
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
   
End Sub


Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   tab1.Tab = sActiveTab(3)
   If bSecSet = 1 Then
      '        User.Group5 = True
      If Secure.UserFinaG3E = 1 Then
         lstEdt.enabled = True
      Else
         lstEdt.enabled = False
         zm(0).Visible = True
      End If
      If Secure.UserFinaG3V = 1 Then
         lstVew.enabled = True
      Else
         lstVew.enabled = False
         zm(1).Visible = True
      End If
      If Secure.UserFinaG3F = 1 Then
         lstFun.enabled = True
      Else
         lstFun.enabled = False
         zm(2).Visible = True
      End If
      'New Code here *****:
      If bCustomerGroups(3) = 0 Then
         ' they didn't sign up, but we want them to see but not use
         lstEdt.enabled = False
         lstVew.enabled = False
         lstFun.enabled = False
         lblcustomer.ForeColor = ES_BLUE
         lblcustomer.Visible = True
         '            User.Group5 = 1
      End If
      'End New Code ******
   End If
   '    If User.Group5 Then
   lstEdt.AddItem "Chart Of Accounts" '0
   lstEdt.AddItem "Journal Entry" '1
'   lstEdt.AddItem "Financial Statement Structure" '2
   lstEdt.AddItem "Fiscal Years" '3
   'lstEdt.AddItem "Divisions"                          '4
   lstEdt.AddItem "Account Budgets" '5
   lstEdt.AddItem "Account Numbers For Parts" '6
   lstEdt.AddItem "Revise Division Information" '7
   lstEdt.AddItem "Product Codes" '8
   lstEdt.AddItem "Cash Account Reconciliation" '9
   
   lstVew.AddItem "Chart Of Accounts" '0
   lstVew.AddItem "Journal Entry List" '1
   lstVew.AddItem "General Journal Entry" '2
   'lstVew.AddItem "Old Detailed General Ledger" '3
   lstVew.AddItem "Trial Balance" '4 Old Trial Balance
   lstVew.AddItem "Income Statement" '5
   lstVew.AddItem "Balance Sheet" '6
   lstVew.AddItem "Divisions" '7
   lstVew.AddItem "Budgets" '8
   lstVew.AddItem "Accounts By Part Number" '9
   lstVew.AddItem "Proforma Income" '10
   lstVew.AddItem "Income/Expense Comparison" '11
   lstVew.AddItem "Account Balance" '12
   lstVew.AddItem "Cash Account Reconciliation" '13
   'lstVew.AddItem "Detailed General Ledger" '14
   'lstVew.AddItem "Trial Balance" '15
   lstVew.AddItem "Financial Statement Structure" '2
   lstVew.AddItem "Rolling Income Statement" '17
   
   If (GetTopSumAcctFlag = 1) Then
      lstVew.AddItem "Trial Balance - Top Summary Account" '18
   End If
   
   lstFun.AddItem "Post A Journal Entry" '0
   lstFun.AddItem "Copy A Journal Entry" '1
   lstFun.AddItem "Copy A Multiple Journal Entry" '2
   lstFun.AddItem "Cancel A Journal Entry" '3
   lstFun.AddItem "Cancel A Posted Journal Entry" '4
   lstFun.AddItem "Change GL Account" '5
   lstFun.AddItem "Open/Close GL Accounting Periods" '6
   lstFun.AddItem "Import Payroll Journal from Excel" '7
   lstFun.AddItem "Import General Journal from Excel" '8
   '    Else
   '
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
      Case 0
         diaGLe01a.Show
      Case 1
         diaGLe02a.Show
'      Case 2
'         diaGLe03a.Show
      Case 2
         diaGLe04a.Show
      Case 3
         diaGLe06a.Show
      Case 4
         diaGLe07a.Show
      Case 5
         diaCdivs.Show
      Case 6
         diaPcode.Show
      Case 7
         diaGLe10a.Show
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
         diaGLf01a.Show
      Case 1
         diaGLf02a.Show
      Case 2
         diaGLf07a.Show
      Case 3
         diaGLf03a.Show
      Case 4
         diaGLf03a.Caption = "Cancel A Posted Journal Entry"
         diaGLf03a.Show
      Case 5
         diaGLf05a.Show
      Case 6
         diaGLf06a.Show
      Case 7
         diaGLf08.Show
      Case 8
         diaGLf09.Show
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
   'Dim b As Byte
   If KeyAscii <> 13 Then Exit Sub
   'MouseCursor 13
   'b = 1
   Select Case lstVew.ListIndex
      Case 0
         diaGLp01a.Show
      Case 1
         diaGLp02a.Show
      Case 2
         diaGLp03a.Show
'      Case 3
'         diaGLp04a.Show
      Case 3
         'We are going to reuse the same form
         'the was used on GL Detail just change to form caption
         diaGLp04a.Caption = "Trial Balance (Report)"
         diaGLp04a.Show
      Case 4
         'diaGLp06a.Show
         diaGLp06a_new.Show
      Case 5
         'diaGLp08a.Show
         diaGLp08a_new.Show
      Case 6
         diaPco01.Show
      Case 7
         diaGLp09a.Show
      Case 8
         diaGLp11a.Show
      Case 9
         'diaGLp12a.Show
         diaGLp12a_New.Show
      Case 10
         diaGLp13a.Show
      Case 11
         diaGLp16a.Show
      Case 12
         diaGLp15a.Show
'      Case 14
'         diaGLp04b.Show
'      Case 15
'         'We are going to reuse the same form
'         'the was used on GL Detail just change to form caption
'         diaGLp04b.Caption = "Trial Balance (Report)"
'         diaGLp04b.Show
      Case 13
         diaGLe03a.Show
      Case 14
         diaGLp17a.Show
      Case 15
         diaGLp18a.Show
      Case Else
         'b = 0
   End Select
   'If b = 1 Then Hide Else MouseCursor 0
End Sub

Private Sub lstVew_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   
   lstVew_KeyPress 13
   
   '    Dim b As Byte
   '    MouseCursor 13
   '    b = 1
   '    Select Case lstVew.ListIndex
   '        Case 0
   '            diaGLp01a.Show
   '        Case 1
   '            diaGLp02a.Show
   '        Case 2
   '            diaGLp03a.Show
   '        Case 3
   '            diaGLp04a.Show
   '        Case 4
   '            'We are going to reuse the same form
   '            'the was used on GL Detail just change to form caption
   '            diaGLp04a.Caption = "Trial Balance (Report)"
   '            diaGLp04a.Show
   '        Case 5
   '            diaGLp06a.Show
   '        Case 6
   '            diaGLp08a.Show
   '        Case 7
   '            diaPco01.Show
   '        Case 8
   '            diaGLp09a.Show
   '        Case 9
   '            diaGLp11a.Show
   '        Case 10
   '            diaGLp12a.Show
   '        Case 11
   '            diaGLp13a.Show
   '            'diaGARMAN.Show
   '        Case 12
   '            diaGLp14a.Show
   '        Case 13
   '            diaGLp15a.Show
   '        Case Else
   '            b = 0
   '    End Select
   '    If b = 1 Then Hide Else MouseCursor 0
End Sub

Private Sub tab1_Click(PreviousTab As Integer)
   sActiveTab(3) = tab1.Tab
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
