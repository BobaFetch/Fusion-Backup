VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form zGr4Matl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventory Management"
   ClientHeight    =   4515
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4995
   ControlBox      =   0   'False
   Icon            =   "Gr4Matl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Gr4Matl.frx":08CA
   ScaleHeight     =   4515
   ScaleWidth      =   4995
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
            Picture         =   "Gr4Matl.frx":1194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr4Matl.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr4Matl.frx":1B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr4Matl.frx":1CB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstSelect 
      Height          =   2790
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
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "Gr4Matl.frx":1DE0
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   3972
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4932
      _ExtentX        =   8705
      _ExtentY        =   7011
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
      FormDesignHeight=   4515
      FormDesignWidth =   4995
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1800
      Top             =   4320
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
            Picture         =   "Gr4Matl.frx":258E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr4Matl.frx":2A3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr4Matl.frx":2EDE
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
Attribute VB_Name = "zGr4Matl"
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
   cUR.CurrentGroup = "Invm"
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
   
End Sub

Private Sub Form_Load()
   ActiveTab = "E"
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   FillEdit
   If bCustomerGroups(4) = 0 Then
      GroupTabPermissions Me
   End If
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set zGr4Matl = Nothing
   
End Sub

Private Sub lstSelect_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   If ActiveTab = "E" Then
      Select Case lstSelect.ListIndex
         Case 0
            CyclCYe01a.Show
         Case 1
            MatlMMe01a.Show
         Case 2
            CyclCYe02a.Show
'         Case 3
'            CyclCYe03a.Show      'old
         Case 3
            CyclCYe04.Show
         Case 4
            CyclCYe05.Show
         Case 5
            MatlMMe02a.Show
         Case Else
            b = 0
      End Select
   
   ElseIf ActiveTab = "V" Then
      Select Case lstSelect.ListIndex
         Case 0
            MatlMMp01a.Show
         Case 1
            MatlMMp02a.Show
         Case 2
            InvcINp01a.Show
         Case 3
            InvcINp02a.Show
         Case 4
            MatlMMp03a.Show
         Case 5
            MatlMMp04a.Show
         Case 6
            MatlMMp05a.Show
         Case 7
            InvcINp07a.Show
         Case 8
            CyclCYp01a.Show
         Case 9
            CyclCYp02a.Show
         Case 10
            CyclCYp03a.Show
         Case 11
            CyclCYp04a.Show
         Case Else
            b = 0
      End Select
   
   Else 'Functions
      Select Case lstSelect.ListIndex
         Case 0
            MatlMMf01a.Show
         Case 1
            CyclCYf01a.Show
         Case 2
            CyclCYf02a.Show
         Case 3
            CyclCYf03a.Show
         Case 4
            CyclCYf04a.Show
         'Case 5
         '   CyclCYf05a.Show     'old
         Case 5
            CyclCYf07.Show
         Case 6
            CyclCYf06a.Show
         Case 7
            MatlMMf02a.Show
         Case 8
            CyclCYf08.Show
         Case 9
            MatlMMf03a.Show
         Case 10
            MatlMMf04a.Show
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
   bActiveTab(4) = Tab1.SelectedItem.Index
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
   If Secure.UserInvcG4E = 1 Then
      lstSelect.AddItem "Create ABC Classes"             '"Inventory ABC Class Setup"
      lstSelect.AddItem "Assign ABC Classes"             '"Inventory ABC Classes"
      lstSelect.AddItem "Initialize a Cycle Count"       '"New Cycle Count"
      'lstSelect.AddItem "Assign Parts to a Cycle Count (Old)"  '"Revise A Cycle Count"
      lstSelect.AddItem "Assign Parts to a Cycle Count"  '"Revise A Cycle Count"
      lstSelect.AddItem "Add Parts/Lots to a Cycle Count"
      lstSelect.AddItem "Standard Cost"
      lstSelect.Enabled = True
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub

Private Sub FillView()
   ActiveTab = "V"
   lstSelect.Clear
   If Secure.UserInvcG4V = 1 Then
      lstSelect.AddItem "Part Activity With Cost "
      lstSelect.AddItem "Part Quantity Activity "
      lstSelect.AddItem "Parts By Part Number"
      lstSelect.AddItem "Parts By Description "
      lstSelect.AddItem "Part Numbers Without An ABC Class"
      lstSelect.AddItem "ABC Class By Part Number"
      lstSelect.AddItem "Inventory Due By ABC Class And Date"
      lstSelect.AddItem "Part Numbers Without An Inventory Location"
      lstSelect.AddItem "Preview A Cycle Count"
      lstSelect.AddItem "Count Sheets"    '"Print A Cycle Count For Inventory"
      lstSelect.AddItem "Inventory Variance Reconciliation"
      lstSelect.AddItem "Cycle Count Status"
      lstSelect.Enabled = True
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub

Public Sub FillFunctions()
   ActiveTab = "F"
   lstSelect.Clear
   If Secure.UserInvcG4F = 1 Then
      lstSelect.AddItem "Adjust Part Quantity"
      lstSelect.AddItem "Update ABC Class Codes By Standard Cost"
      lstSelect.AddItem "Update Inventory Dates For ABC Classes"
      lstSelect.AddItem "Unlock A Locked Cycle Count"
      lstSelect.AddItem "Delete A Cycle Count"
      'lstSelect.AddItem "ABC Inventory Reconciliation (Old)"
      lstSelect.AddItem "ABC Inventory Reconciliation (Physical Cycle Count)"
      lstSelect.AddItem "Mark ABC Cycle Count Audited"
      lstSelect.AddItem "Update Inventory Activity Standard Costs"
      lstSelect.AddItem "Update based on physical counts"
      lstSelect.AddItem "Adjust Part QOH"
      lstSelect.AddItem "Adjust LotQty Item"
      lstSelect.Enabled = True
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub
