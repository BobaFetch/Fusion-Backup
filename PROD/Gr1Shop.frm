VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form zGr1Shop 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Orders"
   ClientHeight    =   5235
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4920
   ControlBox      =   0   'False
   Icon            =   "Gr1Shop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Gr1Shop.frx":08CA
   ScaleHeight     =   5235
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   5160
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
            Picture         =   "Gr1Shop.frx":1194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Shop.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Shop.frx":1B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Shop.frx":1CB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstSelect 
      Height          =   3960
      Left            =   430
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
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "Gr1Shop.frx":1DE0
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   4800
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
      Top             =   4815
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5235
      FormDesignWidth =   4920
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
            Picture         =   "Gr1Shop.frx":258E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Shop.frx":2A3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Shop.frx":2EDE
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
      Top             =   4800
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "zGr1Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'8/14/06 Added Pictures By Routing Operation (Engr)
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
   cUR.CurrentGroup = "Shop"
   
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
   Set zGr1Shop = Nothing
   
End Sub






Private Sub lstSelect_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   If ActiveTab = "E" Then
      Select Case lstSelect.ListIndex
         Case 0
            ShopSHe01a.Show
         Case 1
            ShopSHe02a.Show
         Case 2
            ShopSHe05a.Show
         Case 3
            ShopSHe06a.Show
         Case 4
            MrplMRe01.Show
         Case 5
            MrplMRe03.Show
         Case 6
            MrplMRe04a.Show
         Case 7
            MrplMRe04.Show
         Case Else
            b = 0
      End Select
   ElseIf ActiveTab = "V" Then
      Dim ssrs As New SSRSViewer
      Select Case lstSelect.ListIndex
         Case 0
            If ES_CUSTOM = "WATERJET" Then
               awiShopSHp01a.Show
'            ElseIf ES_CUSTOM = "JEVCO" Then
'               jevShopSHp01a.Show
            Else
               ShopSHp01a.Show
            End If
         Case 1
            ShopSHp02a.Show
         Case 2
            ShopSHp03a.Show
         Case 3
            ShopSHp04a.Show
         Case 4
            ShopSHp05a.Show
         Case 5
            ShopSHp06a.Show
            '            Case 6
            '                ShopSHp07a.Show
            '            Case 7
            '                ShopSHp08a.Show
            '            Case 8
            '                ShopSHp09a.Show
            '            Case 9
            '                ShopSHp10a.Show
         Case 6
            ShopSHp11a.Show
         Case 7
            ShopSHp12a.Show
         Case 8
            PickMCp01a.Show
            '            Case 13
            '                ShopSHp13a.Show
         Case 9
            ShopSHp14a.Show
         Case 10
            RoutRTp07a.Show
         Case 11
            ShopSHp17a.Show
         Case 12
            MrplMRp09a.Show
         Case 13
            MrplMRp10a.Show
         Case 14
            ShopSHp22a.Show
'         Case 15
'            ShopSHp24.Show
         
         Case 15
            ssrs.report = "Multilevel Cost Analysis"
            ssrs.Show
         
         Case Else
            b = 0
      End Select
   Else 'Functions
      Select Case lstSelect.ListIndex
         Case 0
            ShopSHf01a.Show
            '            Case 1
            '                ShopSHf02a.Show
         Case 1
            ShopSHf03a.Show
            '            Case 3
            '                ShopSHf04a.Show
            '            Case 4
            '                ShopSHf05a.Show
         Case 2
            PickMCe05a.Show
         Case 3
            ShopSHf06a.Show
         Case 4
            CustSh01.Show
         Case 5
            ShopSHf07.Show
        Case 6
            ShopSHF08a.Show
         Case 7
            diaAPe12a.Show
         Case 8
            ShopSHF09a.Show
         Case 9
            ShopSHF10a.Show
         Case 10
            ShopSHF11a.Show
         Case 11
            WCSchList.Show
         Case 12
            WCSchedSheet.Show
         Case 13
            SelLotAvailable.Show
         Case 14
            ShopSHf13.Show
         Case Else
            b = 0
      End Select
   End If
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub


Private Sub lstSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   
   lstSelect_KeyPress 13
   
End Sub



Private Sub tab1_Click()
   bActiveTab(4) = tab1.SelectedItem.Index
   Select Case bActiveTab(4)
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
   If Secure.UserProdG1E = 1 Then
      lstSelect.AddItem "New Manufacturing Order"
      lstSelect.AddItem "Revise A Manufacturing Order"
      lstSelect.AddItem "Part Manufacturing Parameters "
      lstSelect.AddItem "Release Manufacturing Orders "
      lstSelect.AddItem "Auto Create MO from MRP's"
      lstSelect.AddItem "Revise MO Dates, Quantity and Status"
      lstSelect.AddItem "Auto Release SC Status MO's"
      lstSelect.AddItem "Auto Release SC Status MO's (Old)"
      lstSelect.Enabled = True
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub

Private Sub FillView()
   ActiveTab = "V"
   lstSelect.Clear
   If Secure.UserProdG1V = 1 Then
      lstSelect.AddItem "Manufacturing Orders "
      lstSelect.AddItem "MO History"
      lstSelect.AddItem "Manufacturing Orders By Date "
      lstSelect.AddItem "Manufacturing Orders By Part "
      lstSelect.AddItem "Manufacturing Order Status "
      lstSelect.AddItem "MO Status By Customer "
      
      '2: 6 & 7
      '        lstSelect.AddItem "Late MO's By Operation " '8
      lstSelect.AddItem "Sales Order Allocations By Customer" '10
      lstSelect.AddItem "Sales Order Allocations By MO" '11
      lstSelect.AddItem "Individual Pick List " '12
      lstSelect.AddItem "Manufacturing Orders Splits" '14
      lstSelect.AddItem "Pictures By Routing Operation" '15
      lstSelect.AddItem "View Work Center MO Schedules" '16
      lstSelect.AddItem "MO Early/Late dates" '16
      lstSelect.AddItem "MO Early/Late dates for Purchase Parts" '17
      lstSelect.AddItem "Part Information and Status" '18
'      lstSelect.AddItem "Priority Dispatch Report" '15
      lstSelect.AddItem "Multilevel MO Cost (SSRS)" '15
      
      lstSelect.Enabled = True
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub

Public Sub FillFunctions()
   ActiveTab = "F"
   lstSelect.Clear
   If Secure.UserProdG1F = 1 Then
      lstSelect.AddItem "Cancel A Manufacturing Order" '0
      lstSelect.AddItem "Update MO Routings" '1
      lstSelect.AddItem "Add A Pick List Item" '2
      lstSelect.AddItem "Split A Manufacturing Order" '3
      lstSelect.AddItem "MO Priority/Work Center Schedules" '4
      lstSelect.AddItem "Re-Schedule MO's" '5
      lstSelect.AddItem "Export MO's to Excel"  '6
      lstSelect.AddItem "Add Part Prefix to SINC MfgRel report" '7
      lstSelect.AddItem "Create SINC Manufacturing Releases report" '8
      lstSelect.AddItem "Create SINC Deliveries from Supply Base" '9
      lstSelect.AddItem "Create SINC Quality from Supply Base" '10
      
      lstSelect.AddItem "Work Center List for Schedules View" '11
      lstSelect.AddItem "All Work Center Schedules" '12
      lstSelect.AddItem "Pick LotNumbers for POM pick" '13
      lstSelect.AddItem "Import MO Priorities from Excel" '14
      
      lstSelect.Enabled = True
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub
