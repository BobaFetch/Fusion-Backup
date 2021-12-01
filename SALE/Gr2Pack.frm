VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form zGr2Pack 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pack Slips And Shipping"
   ClientHeight    =   5355
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5025
   ControlBox      =   0   'False
   Icon            =   "Gr2Pack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Gr2Pack.frx":08CA
   ScaleHeight     =   5355
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   5040
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
            Picture         =   "Gr2Pack.frx":1194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr2Pack.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr2Pack.frx":1B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr2Pack.frx":1CB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstSelect 
      Height          =   3765
      Left            =   435
      TabIndex        =   3
      Top             =   600
      Width           =   4065
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "    Close (Escape)    "
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "Gr2Pack.frx":1DE0
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8281
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
      Top             =   4695
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5355
      FormDesignWidth =   5025
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1800
      Top             =   4920
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
            Picture         =   "Gr2Pack.frx":258E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr2Pack.frx":2A3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr2Pack.frx":2EDE
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
      Height          =   252
      Left            =   840
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   2892
   End
End
Attribute VB_Name = "zGr2Pack"
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
Dim bPack As Byte

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
   cUR.CurrentGroup = "Pack"
   bPack = AllowPsPrepackaging()
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
   
End Sub

Private Sub Form_Load()
   ActiveTab = "E"
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   'If bSecSet = 1 Then User.Group2 = 1
   FillEdit
   If bCustomerGroups(2) = 0 Then
      GroupTabPermissions Me
      'User.Group2 = 1
   End If
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set zGr2Pack = Nothing
   
End Sub




Private Sub lstSelect_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   If ActiveTab = "E" Then
      Select Case lstSelect.ListIndex
         Case 0
            PackPSe01a.Show
         Case 1
            PackPSe02a.Show
         Case 2
            If bPack = 1 Then
               PackPSe03a.Show
            Else
               MouseCursor 0
               MsgBox "You Must Allow Prepackaging To Use This Function.", _
                  vbInformation, Caption
               b = 0
            End If
         Case 3
            If bPack = 1 Then
               PackPSe04a.Show
            Else
               b = 0
               MsgBox "This Feature Is Only Available When Pre-Packing Is Set.", _
                  vbInformation
            End If
         Case 4
            PackPSe05a.Show
         Case 5
            PackPSe06a.Show
         Case 6
            PackPSe07a.Show
         Case Else
            b = 0
      End Select
   ElseIf ActiveTab = "V" Then
      Dim ssrs As SSRSViewer
      Select Case lstSelect.ListIndex
         Case 0
            PackPSp01a.cbShowRange.Value = vbChecked    'Added for Ticket #80687
            PackPSp01a.Show
         Case 1
            PackPSp02a.Show
         Case 2
            PackPSp03a.Show
         Case 3
            PackPSp04a.Show
         Case 4
            PackPSp05a.Show
         Case 5
            PackPSp06a.Show
         Case 6
            PackPSp07a.Show
         Case 7
            PackPSp08a.Show
         Case 8
            PackPSp09a.Show
         Case 9
            PackPSp10a.Show
         Case 10
            PackPSp11a.Show
         Case 11
            PackPSp12a.Show
         Case 12
            PackPSp13a.Show
         Case 13
            PackPSp14a.Show
         Case 14
             PackPSp15a.Show
        Case 15
            diaARp16a.Show
        Case 16
            PackPSp16a.Show
        Case 17
            PackPSp17a.Show
        Case 18
            PackPSp18a.Show
        Case 19
            Set ssrs = New SSRSViewer
            ssrs.report = "Packslip with BOM Lot Detail"
            ssrs.Show
        Case 20
            Set ssrs = New SSRSViewer
            ssrs.report = "Inventory Available to Ship with Lots"
            ssrs.Show
         Case Else
            b = 0
      End Select
   Else 'Functions
      Select Case lstSelect.ListIndex
         Case 0
            PackPSf01a.Show
         Case 1
            PackPSf02a.Show
         Case 2
            PackPSf03a.Show
         Case 3
            PackPSf04a.Show
         Case 4
            PackPSf05a.Show
         Case 5
            PackPSf06a.Show
         Case 6
            PackPSf07a.Show
         Case 7
            PackPSf14a.Show
         Case 8
            PackPSf08a.Show
         Case 9
            PackPSf11a.Show
         Case 10
            PackPSf12.Show
         Case 11
            PackPSf13a.Show
         'Case 8
         '   PackPSf09a.Show
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
   bActiveTab(2) = tab1.SelectedItem.Index
   Select Case bActiveTab(2)
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
   If Secure.UserSaleG2E = 1 Then
      'If User.Group2 Then
      lstSelect.AddItem "New Packing Slip "
      lstSelect.AddItem "Revise A Packing Slip "
      lstSelect.AddItem "Ship Packaged Goods (Or Mark Not Shipped)"
      lstSelect.AddItem "Add An Item To A Printed Packing Slip"
      lstSelect.AddItem "Revise A Packing Slip - Not Shipped"
      lstSelect.AddItem "Split Packing Slip Item - Not Shipped"
      lstSelect.AddItem "Revise Request Date For Unshipped PS Items"
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
   If Secure.UserSaleG2V = 1 Then
      'If User.Group2 Then
      lstSelect.AddItem "Packing Slips "
      lstSelect.AddItem "Packing Slip Log "
      lstSelect.AddItem "Packing Slip Edit "
      lstSelect.AddItem "Shipped Items By Part Number"
      lstSelect.AddItem "Shipped Items By Shipping Date"
      lstSelect.AddItem "Packing Slips Not Printed" '6
      lstSelect.AddItem "Printed Packing Slips Not Invoiced "
      lstSelect.AddItem "Packing Slip Pick List"
      lstSelect.AddItem "Packing Slips Printed Not Shipped" '9
      lstSelect.AddItem "Shipped Items By Sales Person"
      lstSelect.AddItem "Shipped Items By Customer"
      lstSelect.AddItem "Unshipped Packing Slips By SO Request Date"
      lstSelect.AddItem "Company Transfers"
      lstSelect.AddItem "Inventory Available To Ship"
      lstSelect.AddItem "Reprint Inventory Labels"
      lstSelect.AddItem "Customer Delivery Performance"
      lstSelect.AddItem "Packing Slip Label"
      lstSelect.AddItem "Shipment Manifest"
      lstSelect.AddItem "Special Pickup Job Request"
      lstSelect.AddItem "Packing Slip with BOM Detail (SSRS)"
      lstSelect.AddItem "Inventory Available to Ship with Lots (SSRS)"
      'Lots Available by Part
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
   If Secure.UserSaleG2F = 1 Then
      'If User.Group2 Then
      lstSelect.AddItem "Cancel A Packing Slip "
      lstSelect.AddItem "Cancel Packing Slip Printing"
      lstSelect.AddItem "Cancel A Packing Slip Shipped Flag"
      lstSelect.AddItem "Cancel A Packing Slip Item (Not Printed)"
      lstSelect.AddItem "Cancel A Packing Slip Item (Printed)"
      lstSelect.AddItem "Cancel A Company Transfer"
      lstSelect.AddItem "Export EDI ASN Shipping"
      lstSelect.AddItem "Get EDI ASN Information"
      lstSelect.AddItem "Create ASN (SID) Labels"
      lstSelect.AddItem "Create Advance Shipping Manifest"
      lstSelect.AddItem "Update Customer List for Shipping Manifest"
      lstSelect.AddItem "Export ASN from Manifest"
      'lstSelect.AddItem "Create Invoice EDI File"
      'lstSelect.AddItem "Reprint Inventory Labels"
      
      lstSelect.Enabled = True
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub
