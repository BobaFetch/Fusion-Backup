VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form tabJorn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Journals"
   ClientHeight    =   4545
   ClientLeft      =   1845
   ClientTop       =   1605
   ClientWidth     =   4995
   Icon            =   "tabJorn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4545
   ScaleWidth      =   4995
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
      FormDesignHeight=   4545
      FormDesignWidth =   4995
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
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Subject Help"
      Top             =   4200
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
      PictureUp       =   "tabJorn.frx":000C
      PictureDn       =   "tabJorn.frx":0152
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
      Tabs            =   2
      Tab             =   1
      TabHeight       =   476
      TabCaption(0)   =   "&View       "
      TabPicture(0)   =   "tabJorn.frx":0298
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstVew"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "zm(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Functions"
      TabPicture(1)   =   "tabJorn.frx":06C8
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "zm(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstFun"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.ListBox lstFun 
         Height          =   2985
         ItemData        =   "tabJorn.frx":0AFD
         Left            =   480
         List            =   "tabJorn.frx":0AFF
         TabIndex        =   4
         Top             =   600
         Width           =   4065
      End
      Begin VB.ListBox lstVew 
         Height          =   2985
         ItemData        =   "tabJorn.frx":0B01
         Left            =   -74520
         List            =   "tabJorn.frx":0B03
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
         Index           =   1
         Left            =   480
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
   Begin VB.Label lblcustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This Feature Is Not Available"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   8
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
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "tabJorn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
   cUR.CurrentGroup = "Jorn"
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   If sActiveTab(5) > 1 Then sActiveTab(5) = 1
   tab1.Tab = sActiveTab(5)
   If bSecSet = 1 Then
      '        User.Group4 = True
      If Secure.UserFinaG4E = 1 Then
         '  lstEdt.Enabled = True
      Else
         '  lstEdt.Enabled = False
         '  zm(0).Visible = True
      End If
      If Secure.UserFinaG4V = 1 Then
         lstVew.enabled = True
      Else
         lstVew.enabled = False
         zm(1).Visible = True
      End If
      If Secure.UserFinaG4F = 1 Then
         lstFun.enabled = True
      Else
         lstFun.enabled = False
         zm(2).Visible = True
      End If
      'New Code here *****:
      If bCustomerGroups(4) = 0 Then
         ' they didn't sign up, but we want them to see but not use
         'lstEdt.Enabled = False
         lstVew.enabled = False
         lstFun.enabled = False
         lblcustomer.ForeColor = ES_BLUE
         lblcustomer.Visible = True
         '            User.Group4 = 1
      End If
      'End New Code ******
   End If
   '    If User.Group4 Then
   lstVew.AddItem "Journals (Report)"
   lstVew.AddItem "Journal Status (Report)"
   'lstVew.AddItem " 2. Purchases Journal"
   'lstVew.AddItem " 3. Manual Check Journal"
   'lstVew.AddItem " 3. Computer Check Journal"
   'lstVew.AddItem " 4. Sales Journal"
   'lstVew.AddItem " 5. Cash Receipts Journal"
   'lstVew.AddItem " 6. Time Journal"
   'lstVew.AddItem " 7. Payroll Labor Journal"
   'lstVew.AddItem " 8. Payroll Disbursements Journal"
   'lstVew.AddItem " 9. Cash Transfer Journal"
   'lstVew.AddItem "10. Material/Services Overhead Journal"
   
   lstFun.AddItem "Open Journals"
   lstFun.AddItem "Close Journals"
   lstFun.AddItem "Post Journals"
   lstFun.AddItem "Open A Closed Journal"
   lstFun.AddItem "Copy Journals To Summary Account"
   lstFun.AddItem "Post Year End Journals"
   '    Else
   '        lstVew.AddItem "No Group Permissions"
   '        tab1.enabled = False
   '    End If
End Sub

Private Sub Form_Resize()
   'Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set tabClos = Nothing
End Sub

Private Sub lstFun_KeyPress(KeyAscii As Integer)
   Dim b As Byte
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   Select Case lstFun.ListIndex
      Case 0
         diaJRf01a.Show
      Case 1
         diaJRf02a.Show
      Case 2
         diaJRf03a.Show
      Case 3
         diaJRf05a.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
End Sub

Private Sub lstFun_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim b As Byte
   MouseCursor 13
   b = 1
   Select Case lstFun.ListIndex
      Case 0
         diaJRf01a.Show
      Case 1
         diaJRf02a.Show
      Case 2
         diaJRf03a.Show
      Case 3
         diaJRf05a.Show
      Case 4
         diaJRf06a.Show
      Case 5
         diaJRf07a.Show
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
         diaJRp01a.Show
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
         diaJRp01a.Show
      Case 1
         diaJRp02a.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
End Sub

Private Sub tab1_Click(PreviousTab As Integer)
   sActiveTab(5) = tab1.Tab
End Sub

Private Sub Tab1_GotFocus()
   Select Case tab1.Tab
      Case 0
         lstVew.TabIndex = 1
      Case 1
         lstFun.TabIndex = 1
   End Select
End Sub

Private Sub tab1_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
End Sub
