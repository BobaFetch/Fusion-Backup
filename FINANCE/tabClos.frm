VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form tabClos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Closing"
   ClientHeight    =   5100
   ClientLeft      =   1845
   ClientTop       =   1605
   ClientWidth     =   5130
   Icon            =   "tabClos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5100
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   4560
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5100
      FormDesignWidth =   5130
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3780
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4530
      Width           =   1095
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Subject Help"
      Top             =   4680
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
      PictureUp       =   "tabClos.frx":000C
      PictureDn       =   "tabClos.frx":0152
   End
   Begin TabDlg.SSTab tab1 
      Height          =   4455
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   476
      TabCaption(0)   =   "&View       "
      TabPicture(0)   =   "tabClos.frx":0298
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "zm(0)"
      Tab(0).Control(1)=   "lstVew"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Functions"
      TabPicture(1)   =   "tabClos.frx":06C8
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "zm(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstFun"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.ListBox lstVew 
         Height          =   3570
         ItemData        =   "tabClos.frx":0AFD
         Left            =   -74520
         List            =   "tabClos.frx":0AFF
         TabIndex        =   5
         Top             =   600
         Width           =   4065
      End
      Begin VB.ListBox lstFun 
         Height          =   3570
         ItemData        =   "tabClos.frx":0B01
         Left            =   480
         List            =   "tabClos.frx":0B03
         TabIndex        =   4
         Top             =   600
         Width           =   4065
      End
      Begin VB.Label zm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No User Permissions"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   -74520
         TabIndex        =   7
         Top             =   4200
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
         TabIndex        =   6
         Top             =   4200
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
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "tabClos"
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
   cUR.CurrentGroup = "Cred"
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   tab1.Tab = sActiveTab(2)
   If bSecSet = 1 Then
      '        User.Group3 = True
      'If Secure.UserFinaG3E = 1 Then
      'lstEdt.Enabled = True
      'Else
      'lstEdt.Enabled = False
      '    zm(0).Visible = True
      'End If
      If Secure.UserFinaG5V = 1 Then
         lstVew.enabled = True
      Else
         lstVew.enabled = False
         zm(0).Visible = True
      End If
      If Secure.UserFinaG5F = 1 Then
         lstFun.enabled = True
      Else
         lstFun.enabled = False
         zm(1).Visible = True
      End If
   End If
   
   
   
   'New Code here *****:
   If bCustomerGroups(5) = 0 Then
      ' they didn't sign up, but we want them to see but not use
      'lstEdt.Enabled = False
      lstVew.enabled = False
      lstFun.enabled = False
      lblcustomer.ForeColor = ES_BLUE
      lblcustomer.Visible = True
      '            User.Group3 = 1
   End If
   'End New Code ******
   '    If User.Group3 Then
   lstVew.AddItem "Received But Not Invoiced"                     '0
   lstVew.AddItem "Material Purchase Price Variance"              '1
   lstVew.AddItem "Raw Material/Finished Goods Inventory"         '2
   lstVew.AddItem "Inventory Movement To WIP"                '3
   lstVew.AddItem "Work In Process"                               '4
   lstVew.AddItem "Inventory Movement From WIP"              '5
   lstVew.AddItem "Material Movement To Project MO"               '6
   lstVew.AddItem "Scrapped MO's"                            '7
   lstVew.AddItem "Inventory Adjustments"                   '8
   lstVew.AddItem "*Print Standard Product Variance "         '9
   'lstVew.AddItem "Overhead Applied"                              '10
   'lstVew.AddItem "*Shipped But Not Invoiced"                '11
   lstVew.AddItem "Cost Of Goods Sold"                            '12
   lstVew.AddItem "Inventory Movement from WIP - Restock"   '13
   lstVew.AddItem "*Picked Raw Materials & Subassemblies"    '14
   lstVew.AddItem "Raw Material/Finished Goods Inventory Cost Detail"         '2
   lstVew.AddItem "Work In Process - Calculated OH"         '2
   lstVew.AddItem "Detail MO Cost Of Goods Sold"         '2
	  
   lstFun.AddItem "Update Sales Activity Standard Cost"        '0
   lstFun.AddItem "Update Sales Order Account Distributions"   '1
   lstFun.AddItem "Close A Manufacturing Order"                '2
   lstFun.AddItem "Open A Closed Manufacturing Order"          '3
   lstFun.AddItem "Close All Completed Manufacturing Orders"   '4
   'lstFun.AddItem "Recost Purchase Orders"                     '5
   'lstFun.AddItem "Recost Manufacturing Orders"                '6
   'lstFun.AddItem "Recost Shipments"                           '7
   
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
         diaCLf01a.Show
      Case 1
         diaARf04a.Show
      Case 2
         ShopSHf04a.Show
      Case 3
         ShopSHf05a.Show
      Case 4
         ShopSHf07a.Show
    '  Case 5
    '     Close07.Show
    '  Case 6
    '     Close06.Show
    '  Case 7
    '     Close08.Show
      Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
End Sub


Private Sub lstFun_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
                             
   lstFun_KeyPress 13
   
'   Dim b As Byte
'   MouseCursor 13
'   b = 1
'   Select Case lstFun.ListIndex
'      Case 0
'         diaCLf01a.Show
'      Case 1
'         diaARf04a.Show
'      Case Else
'         b = 0
'   End Select
'   If b = 1 Then Hide Else MouseCursor 0
End Sub


Private Sub lstVew_KeyPress(KeyAscii As Integer)
   Dim b As Byte
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   Select Case lstVew.ListIndex
      Case 0
         diaAPp11a.Show
      Case 1
         diaCLp01a.Show
      Case 2
         diaRMFG_new.Show
         'diaRMFG.Show
      Case 3
         diaCLp07a.Show
      Case 4
         diaWip.Show
      Case 5
         diaCLp08a.Show
      Case 6
         diaCLp06a.Show
      Case 7
         diaCLp09a.Show
      Case 8
         diaCLp14a.Show
      Case 9
         diaCLp11a.Show
      'Case 10
      '   diaCLp05a.Show
      'Case 10
      '  '"Shipped But Not Invoiced"
      '   diaCLp13a.Show
      Case 10
         diaCGS.Show
      Case 11
        diaCLp12a.Show
       Case 12
         diaCLp10a.Show
       Case 13
         diaRMFG_Cost.Show
       Case 14
         diaWipOH.Show
       Case 15
         diaDetailCGS.Show
	   Case Else
         b = 0
   End Select
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub

Private Sub lstVew_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   
   lstVew_KeyPress 13
   
   '    Dim b As Byte
   '    MouseCursor 13
   '    b = 1
   '    Select Case lstVew.ListIndex
   '        Case 0
   '            diaAPp11a.Show
   '        Case 1
   '            diaCLp01a.Show
   '        Case 2
   '            diaRMFG.Show
   '        Case 4
   '            diaWip.Show
   '        Case 6
   '            diaCLp06a.Show
   '        Case 10
   '            diaCLp05a.Show
   '        Case 12
   '            diaCGS.Show
   ''        Case 13
   ''            WipTest.Show
   ''        Case 14
   ''            diaWip2.Show
   '        Case Else
   '            b = 0
   '    End Select
   '    If b = 1 Then Hide Else MouseCursor 0
End Sub


Private Sub tab1_Click(PreviousTab As Integer)
   sActiveTab(2) = tab1.Tab
End Sub

Private Sub tab1_GotFocus()
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
