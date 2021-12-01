VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form zGr5EstiPPI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estimating"
   ClientHeight    =   4515
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4920
   ControlBox      =   0   'False
   Icon            =   "Gr5EstiPPI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Gr5EstiPPI.frx":08CA
   ScaleHeight     =   4515
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
            Picture         =   "Gr5EstiPPI.frx":1194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr5EstiPPI.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr5EstiPPI.frx":1B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr5EstiPPI.frx":1CB2
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
      Picture         =   "Gr5EstiPPI.frx":1DE0
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
      FormDesignWidth =   4920
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1440
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
            Picture         =   "Gr5EstiPPI.frx":258E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr5EstiPPI.frx":2A3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr5EstiPPI.frx":2EDE
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
Attribute VB_Name = "zGr5EstiPPI"
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
   cUR.CurrentGroup = "Esti"
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
   
End Sub

Private Sub Form_Load()
   ActiveTab = "E"
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   'If bSecSet = 1 Then User.Group5 = 1
   FillEdit
   If bCustomerGroups(5) = 0 Then
      GroupTabPermissions Me
      'User.Group5 = 1
   End If
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set zGr5EstiPPI = Nothing
   
End Sub




Private Sub lstSelect_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   If ActiveTab = "E" Then
      Select Case lstSelect.ListIndex
         Case 0
            ' MM TODO: ppiESe02a.Show
         Case 1
            ' MM TODO: Load ppiESe01a
         Case 2
            EstiESe03a.Show
         Case 3
            EstiESe04a.Show
         Case 4
            EstiESe05a.Show
         Case 5
            ' MM TODO: ppiESe03a.Show
         Case Else
            b = 0
      End Select
   ElseIf ActiveTab = "V" Then
      Select Case lstSelect.ListIndex
         Case 0
            ' MM TODO: ppiESp01a.Show
         Case 1
            ' MM TODO: ppiESp02a.Show
         Case 2
            ' MM TODO: ppiESp03a.Show
         Case 3
            EstiESp04a.Show
         Case 4
            EstiESp05a.Show
         Case 5
            EstiESp06a.Show
         Case 6
            EstiESp07a.Show
         Case 7
            EstiESp08a.Show
         Case 8
            ' MM TODO: ppiESp04a.Show
         Case Else
            b = 0
      End Select
   Else 'Functions
      Select Case lstSelect.ListIndex
         Case 0
            ' MM TODO: ppiESf04a.Show
         Case 1
            EstiESf01a.Show
         Case 2
            EstiESf05a.Show
         Case 3
            EstiESf02a.Show
         Case 4
            EstiESf03a.Show
         Case 5
            ' MM TODO: ppiESf05a.Show
         Case 6
            ' MM TODO: ppiESf06a.Show
         Case 7
            EstiESf06a.Show
         Case Else
            b = 0
      End Select
   End If
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub


Private Sub lstSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   lstSelect_KeyPress 13
   
   '    MouseCursor 13
   '    b = 1
   '    If ActiveTab = "E" Then
   '        Select Case lstSelect.ListIndex
   '            Case 0
   '                ppiESe02a.Show
   '            Case 1
   '                Load ppiESe01a
   '            Case 2
   '                EstiESe03a.Show
   '            Case 3
   '                EstiESe04a.Show
   '            Case 4
   '                EstiESe05a.Show
   '            Case 5
   '                ppiESe03a.Show
   '            Case Else
   '                b = 0
   '        End Select
   '    ElseIf ActiveTab = "V" Then
   '        Select Case lstSelect.ListIndex
   '            Case 0
   '                ppiESp01a.Show
   '            Case 1
   '                ppiESp02a.Show
   '            Case 2
   '                ppiESp03a.Show
   '            Case 3
   '                EstiESp04a.Show
   '            Case 4
   '                EstiESp05a.Show
   '            Case 5
   '                EstiESp06a.Show
   '            Case 6
   '                EstiESp07a.Show
   '            Case 7
   '                EstiESp08a.Show
   '            Case 8
   '                ppiESp04a.Show
   '            Case Else
   '                b = 0
   '        End Select
   '    Else 'Functions
   '        Select Case lstSelect.ListIndex
   '            Case 0
   '                ppiESf04a.Show
   '            Case 1
   '                EstiESf01a.Show
   '            Case 2
   '                EstiESf05a.Show
   '            Case 3
   '                EstiESf02a.Show
   '            Case 4
   '                EstiESf03a.Show
   '            Case 5
   '                ppiESf05a.Show
   '            Case 6
   '                ppiESf06a.Show
   '            Case 7
   '                EstiESf06a.Show
   '            Case Else
   '                b = 0
   '        End Select
   '    End If
   '
   '    If b = 1 Then Hide Else MouseCursor 0
   '
End Sub


Private Sub tab1_Click()
   bActiveTab(5) = Tab1.SelectedItem.Index
   Select Case bActiveTab(5)
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
   If Secure.UserEngrG5E = 1 Then
      'If User.Group5 Then
      lstSelect.AddItem "Full Estimate "
      lstSelect.AddItem "Qwik Bid "
      lstSelect.AddItem "Complete Estimates"
      lstSelect.AddItem "Accept Estimates "
      lstSelect.AddItem "Requests For Quotation (RFQ) "
      lstSelect.AddItem "Formula Calculator"
      'lstSelect.AddItem "Test Formula Parse (Demo)"
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
   If Secure.UserEngrG5V = 1 Then
      'If User.Group5 Then
      lstSelect.AddItem "Estimate "
      lstSelect.AddItem "Estimate Summary By Customer "
      lstSelect.AddItem "Estimate Summary By Part Number"
      lstSelect.AddItem "Estimates Not Completed"
      lstSelect.AddItem "Canceled Estimates "
      lstSelect.AddItem "Requests For Quotation By Customer "
      lstSelect.AddItem "Estimating Parts "
      lstSelect.AddItem "Estimate Status "
      lstSelect.AddItem "List Of Formulae"
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
   If Secure.UserEngrG5F = 1 Then
      'If User.Group5 Then
      lstSelect.AddItem "Add/Edit Formulae"
      lstSelect.AddItem "Mark Estimates Not Accepted "
      lstSelect.AddItem "Copy An Estimate "
      lstSelect.AddItem "Cancel An Estimate "
      lstSelect.AddItem "Cancel An RFQ "
      lstSelect.AddItem "Change An Estimating Formula"
      lstSelect.AddItem "Delete An Estimating Formula"
      lstSelect.AddItem "Copy An Estimate Routing"
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
