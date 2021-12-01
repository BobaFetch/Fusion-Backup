VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form zGr3Purc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchasing"
   ClientHeight    =   5010
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4920
   ControlBox      =   0   'False
   Icon            =   "Gr3Purc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Gr3Purc.frx":08CA
   ScaleHeight     =   5010
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
            Picture         =   "Gr3Purc.frx":1194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr3Purc.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr3Purc.frx":1B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr3Purc.frx":1CB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstSelect 
      Height          =   3375
      Left            =   435
      TabIndex        =   3
      Top             =   600
      Width           =   4065
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3780
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "    Close (Escape)    "
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "Gr3Purc.frx":1DE0
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   4320
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
      Top             =   4320
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5010
      FormDesignWidth =   4920
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1440
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
            Picture         =   "Gr3Purc.frx":258E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr3Purc.frx":2A3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr3Purc.frx":2EDE
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
      Top             =   4560
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "zGr3Purc"
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
   cUR.CurrentGroup = "Purc"
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
   
End Sub

Private Sub Form_Load()
   ActiveTab = "E"
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   'If bSecSet = 1 Then User.Group3 = 1
   FillEdit
   If bCustomerGroups(3) = 0 Then
      GroupTabPermissions Me
      'User.Group3 = 1
   End If
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set zGr3Purc = Nothing
   
End Sub




Private Sub lstSelect_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   If ActiveTab = "E" Then
      Select Case lstSelect.ListIndex
         Case 0
            PurcPRe02a.optNew.Value = vbChecked
            PurcPRe02a.Show
            PurcPRe01a.Show
         Case 1
            PurcPRe02a.optNew.Value = vbChecked
            PurcPRe02a.Show
            bPOCaption = 1
            PurcPRe01a.Caption = "New Services Purchase Order"
            PurcPRe01a.optSrv.Value = vbChecked
            PurcPRe01a.Show
         Case 2
            PurcPRe02a.Show
         
        Case 3
            PurcPRe09a.Show
         
         Case 4
            VendorEdit01.Tag = 1
            VendorEdit01.Show
         Case 5
            PurcPRe04a.Show
         Case 6
            PurcPRe05a.Show
         Case 7
            PurcPRe06a.Show
         Case 8
            PurcPRe07a.Show
         Case 9
            PurcPRe08a.Show
         Case 10
            MrplMRe02.Show
         Case Else
            b = 0
      End Select
   ElseIf ActiveTab = "V" Then
      Select Case lstSelect.ListIndex
         Case 0
            PurcPRp01a.Show
         Case 1
            PurcPRp02a.Show
         Case 2
            PurcPRp03a.Show
         Case 3
            PurcPRp04a.Show
         Case 4
            PurcPRp05a.Show
         Case 5
            PurcPRp06a.Show
         Case 6
            PurcPRp07a.Show
         Case 7
            PurcPRp08a.Show
         Case 8
            PurcPRp09a.Show
         Case 9
            PurcPRp10a.Show
         Case 10
            PurcPRp11a.Show
         Case 11
            PurcPRp12a.Show
         Case 12
            PurcPRp13a.Show
         Case 13
            PurcPRp14a.Show
         Case 14
            PurcPRp15a.Show
         Case 15
            PurcPRp16a.Show
         Case 16
            diaAPp23a.Show
         Case 17
            PurcPRp17a.Show
         Case 18
            Dim ssrs As New SSRSViewer
            ssrs.report = "Service POs"
            ssrs.Show
         Case 19
            PurcPRp21a.Show
         Case Else
            b = 0
      End Select
   Else 'Functions
      Select Case lstSelect.ListIndex
         Case 0
            PurcPRf01a.Show
         Case 1
            PurcPRf02a.Show
         Case 2
            PurcPRf03a.Show
         Case 3
            PurcPRf04a.Show
         Case 4
            PurcPRf05a.Show
         Case 5
            PurcPRf06a.Show
         Case 6
            PurcPRf07a.Show
         Case 7
            PurcPRf08a.Show
         Case 8
            PurcPRf09a.Show
         Case 9
            PurcPRf10a.Show
         Case 10
            PurcPRf11a.Show
         Case 11
            PurcPRf12a.Show
         Case 12
            PurcPRf13a.Show
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
   bActiveTab(3) = tab1.SelectedItem.Index
   Select Case bActiveTab(3)
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
   If Secure.UserProdG3E = 1 Then
      'If User.Group3 Then
      lstSelect.AddItem "New Purchase Order"
      lstSelect.AddItem "New Services Purchase Order"
      lstSelect.AddItem "Revise A Purchase Order"
      lstSelect.AddItem "Re-Schedule Purchase Order Item"
      lstSelect.AddItem "Vendors  "
      lstSelect.AddItem "Part Purchasing Information "
      lstSelect.AddItem "Buyers "
      lstSelect.AddItem "Assign Parts To Buyers"
      lstSelect.AddItem "Approved Supplier by Part Number"
      lstSelect.AddItem "Manufacturers"
          lstSelect.AddItem "Auto Create PO from MRP's"
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
   If Secure.UserProdG3V = 1 Then
      'If User.Group3 Then
      lstSelect.AddItem "Purchase Orders " '0
      lstSelect.AddItem "Vendor List " '1
      lstSelect.AddItem "Vendor Directory " '2
      lstSelect.AddItem "Purchase Order Log " '3
      lstSelect.AddItem "Purchase Order Log By Date " '4
      lstSelect.AddItem "Purchase Order Expediting " '5
      lstSelect.AddItem "Purchase Order Expediting By Buyer " '6
      lstSelect.AddItem "Purchasing History By Vendor " '7
      lstSelect.AddItem "Purchasing History By Manufacturing Order" '8
      lstSelect.AddItem "Purchasing History By Part " '9
      lstSelect.AddItem "Outside Services Requirements" '10
      lstSelect.AddItem "Part Numbers By Buyer" '11
      lstSelect.AddItem "Approved Supplier by Part Number List" '12
      lstSelect.AddItem "Purchase Order Requested By" '13
      lstSelect.AddItem "List Of Manufacturers" '14
      lstSelect.AddItem "Manufacturers By Part Number" '15
      lstSelect.AddItem "Vendor Delivery Performance" '16
      lstSelect.AddItem "Vendor Delivery Performance and Inspection Report" '16
      lstSelect.AddItem "Service Purchase Orders (SSRS)" '18
      lstSelect.AddItem "Purchasing History To Date Costing" '19
      
      'lstSelect.AddItem "Purchasing Information By Part"
      'lstSelect.AddItem "Purchasing Lead Time Discrepancy"
      'lstSelect.AddItem "Buy List Report"
      'lstSelect.AddItem "PO Items With Zero Cost"
      'lstSelect.AddItem "Manufacturer Directory"
      'lstSelect.AddItem "Purchasing Expediting By MO"
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
   If Secure.UserProdG3F = 1 Then
      'If User.Group3 Then
      lstSelect.AddItem "Cancel A Purchase Order "
      lstSelect.AddItem "Delete A Vendor"
      lstSelect.AddItem "Change A Vendor Nickname"
      lstSelect.AddItem "Change A Buyer ID"
      lstSelect.AddItem "Delete A Buyer ID"
      lstSelect.AddItem "Revise A PO Line Item Price (Invoiced Item)"
      lstSelect.AddItem "Split A Purchase Order Item"
      lstSelect.AddItem "Open A Canceled Purchase Order"
      lstSelect.AddItem "Change Purchase Order Requested By"
      lstSelect.AddItem "Change A Manufacturer's Nickname"
      lstSelect.AddItem "Delete A Manufacturer"
      lstSelect.AddItem "Export PO Activity to Excel"
      lstSelect.AddItem "Copy Purchase Order"
      'lstSelect.AddItem "Update Open PO's With Part Conversion Information"
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
