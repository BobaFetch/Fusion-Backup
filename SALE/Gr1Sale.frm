VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form zGr1Sale 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Order Processing"
   ClientHeight    =   4875
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4920
   ControlBox      =   0   'False
   Icon            =   "Gr1Sale.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Gr1Sale.frx":08CA
   ScaleHeight     =   4875
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   4440
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   23
      MaskColor       =   -2147483633
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Sale.frx":1194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Sale.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Sale.frx":1B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Sale.frx":1CB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstSelect 
      Height          =   3570
      Left            =   432
      TabIndex        =   3
      Top             =   600
      Width           =   4068
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3780
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "    Close (Escape)    "
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "Gr1Sale.frx":1DE0
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7646
      MultiRow        =   -1  'True
      TabFixedHeight  =   473
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit             "
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&View           "
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Functions"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   480
      Top             =   4092
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4875
      FormDesignWidth =   4920
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1800
      Top             =   4200
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483628
      ImageWidth      =   22
      ImageHeight     =   20
      MaskColor       =   -2147483636
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Sale.frx":258E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Sale.frx":2A3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Sale.frx":2EDE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCustomer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This Feature Is Not Available"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "zGr1Sale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim b As Byte
Dim ActiveTab As String

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 923
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   cUR.CurrentGroup = "Ordr"
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
   
End Sub

Private Sub Form_Load()
   ActiveTab = "E"
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   'If bSecSet = 1 Then User.Group1 = 1
   FillEdit
   If bCustomerGroups(1) = 0 Then
      GroupTabPermissions Me
      'User.Group1 = 1
   End If
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set zGr1Sale = Nothing
   
End Sub




Private Sub lstSelect_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   If ActiveTab = "E" Then
      Select Case lstSelect.ListIndex
         Case 0
            SaleSLe01a.Show
         Case 1
            SaleSLe02a.Show
         Case 2
            SaleSLe03a.Show
         Case 3
            SaleSLe04a.Show
         Case 4
            SaleSLe05a.Show
         Case Else
            b = 0
      End Select
   ElseIf ActiveTab = "V" Then
      Select Case lstSelect.ListIndex
         Case 0
            SaleSLp01a.Show
         Case 1
            SaleSLp02a.Show
         Case 2
            SaleSLp03a.Show
         Case 3
            SaleSLp04a.Show
         Case 4
            SaleSLp05a.Show
         Case 5
            SaleSLp06a.Show
         Case 6
            SaleSLp07a.Show
         Case 7
            ShopSHp11a.Show
         Case 8
            ShopSHp12a.Show
         Case 9
            SaleSLp08a.Show
         Case 10
            SaleSLp09a.Show
         Case 11
            SaleSLp10a.Show
         Case 12
            SaleSLp42a.Show
         Case 13
            BookBKp18a.Show
         Case 14
            EstiESp02a.Show
         Case 15
            SaleSLp11a.Show
         Case 16
            SaleSLp12a.Show
         Case 17
            SaleSLp13a.Show
         Case 18  'Stays at the bottom shows only for INTCOA
            Intavl03.Show
         Case Else
            b = 0
      End Select
   Else 'Functions
      Select Case lstSelect.ListIndex
         Case 0
            SaleSLf01a.Show
         Case 1
            SaleSLf02a.Show
         Case 2
            SaleSLf03a.Show
         Case 3
            SaleSLf04a.Show
         Case 4
            SaleSLf05a.Show
         Case 5
            SaleSLf06a.Show
         Case 6
            SaleSLf07a.Show
         Case 7
            SaleSLf08a.Show
         Case 8
            SaleSLf09a.Show
         Case 9
            SaleSLf10a.Show
         Case 10
            SaleSLf11a.Show
         Case 11
            SaleSLf12a.Show
         Case 12
            SaleSLf13a.Show
         Case 13
            SaleSLf14a.Show
         Case 14
            SaleSLf15a.Show
         Case 15
            SaleSLf15b.Show
         Case 16
            SaleSLf15c.Show
         Case Else
            b = 0
      End Select
   End If
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub


Private Sub lstSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   lstSelect_KeyPress 13

End Sub


Private Sub tab1_Click()
   bActiveTab(1) = tab1.SelectedItem.Index
   Select Case bActiveTab(1)
      Case 1
         FillEdit
      Case 2
         FillView
      Case 3
         FillFunctions
   End Select
   
End Sub

Private Sub FillEdit()
   ActiveTab = "E"
   lstSelect.Clear
   If Secure.UserSaleG1E = 1 Then
      'If User.Group1 Then
      lstSelect.AddItem "New Sales Order "
      lstSelect.AddItem "Revise A Sales Order "
      lstSelect.AddItem "Customers (Enter/Revise)"
      lstSelect.AddItem "Price Books"
      lstSelect.AddItem "Reschedule Sales Order Item"
      lstSelect.Enabled = True
      '        Else
      '            lstSelect.Enabled = False
      '            lstSelect.AddItem "No Group Permissions"
      '        End If
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub

Private Sub FillView()
   ActiveTab = "V"
   lstSelect.Clear
   If Secure.UserSaleG1V = 1 Then
      'If User.Group1 Then
      lstSelect.AddItem "Sales Orders "                           '0
      lstSelect.AddItem "Customer List "                          '1
      lstSelect.AddItem "Customer Directory "                     '2
      lstSelect.AddItem "Customer Sales Order List "              '3
      lstSelect.AddItem "Sales Order Register "                   '4
      lstSelect.AddItem "Customer List By Zip Code "              '5
      lstSelect.AddItem "Customer Sales Order List By PO "        '6
      lstSelect.AddItem "Sales Order Allocations By Customer"     '7
      lstSelect.AddItem "Sales Order Allocations By MO"           '8
      lstSelect.AddItem "Price Books Report"                      '9
      lstSelect.AddItem "Price Books By Customer"                 '10
      lstSelect.AddItem "Price Books By Part Number"              '11
      lstSelect.AddItem "Sales Order Acknowledgements "           '12
      lstSelect.AddItem "Part Availability Report"                '13
      lstSelect.AddItem "Estimating Summary By Customer"          '14
      lstSelect.AddItem "Customer Sales Analysis"                 '15
      lstSelect.AddItem "Customer Shipment Analysis"              '16
      lstSelect.AddItem "Price Book Discrepancy Report"
      'V INTCOA only Leave at the bottom
      If ES_CUSTOM = "INTCOA" Then _
                     lstSelect.AddItem "Part Availability For A Customer"
      lstSelect.Enabled = True
      '        Else
      '            lstSelect.Enabled = False
      '            lstSelect.AddItem "No Group Permissions"
      '        End If
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub

Public Sub FillFunctions()
   ActiveTab = "F"
   lstSelect.Clear
   If Secure.UserSaleG1F = 1 Then
      'If User.Group1 Then
      lstSelect.AddItem "Cancel A Sales Order "
      lstSelect.AddItem "Copy A Sales Order"
      lstSelect.AddItem "Delete A Customer "
      lstSelect.AddItem "Change A Customer Nickname "
      lstSelect.AddItem "Change A Sales Order Number "
      lstSelect.AddItem "Revise Booked Dates"
      lstSelect.AddItem "Change A Price Book ID"
      lstSelect.AddItem "Copy A Price Book"
      lstSelect.AddItem "Delete A Price Book"
      lstSelect.AddItem "Lock/Unlock Sales Orders"
      lstSelect.AddItem "Create Sales Orders From Estimates"
      lstSelect.AddItem "Import Sales Order Data From Exostar"
      lstSelect.AddItem "Export Sales Orders to Excel"
      lstSelect.AddItem "Create Sales Orders From EDI (EDI)"
      lstSelect.AddItem "Import VOI data from external data source"
      lstSelect.AddItem "Add Packslip and Print Packslip From VOI"
      lstSelect.AddItem "Post Invoice For VOI Checks"
      
      lstSelect.Enabled = True
      '        Else
      '            lstSelect.Enabled = False
      '            lstSelect.AddItem "No Group Permissions"
      '        End If
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub
