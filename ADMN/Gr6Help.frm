VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form zGr6Help 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Help"
   ClientHeight    =   4515
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4920
   ControlBox      =   0   'False
   Icon            =   "Gr6Help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Gr6Help.frx":08CA
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
            Picture         =   "Gr6Help.frx":1194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr6Help.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr6Help.frx":1B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr6Help.frx":1CB2
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
      Height          =   375
      Left            =   0
      Picture         =   "Gr6Help.frx":1DE0
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   375
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
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&View    "
            ImageVarType    =   2
            ImageIndex      =   1
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
      Top             =   4080
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483628
      ImageWidth      =   22
      ImageHeight     =   20
      MaskColor       =   -2147483636
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr6Help.frx":1F13
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
Attribute VB_Name = "zGr6Help"
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
   cUR.CurrentGroup = "Help"
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
   
End Sub

Private Sub Form_Load()
   ActiveTab = "V"
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   'If bSecSet = 1 Then User.Group6 = 1
   FillView
   If bCustomerGroups(6) = 0 Then
      GroupTabPermissions Me
      'User.Group6 = 1
   End If
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set zGr6Help = Nothing
   
End Sub




Private Sub lstSelect_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   If ActiveTab = "E" Then
      Select Case lstSelect.ListIndex
         Case 0
            b = 0
         Case Else
            b = 0
      End Select
   ElseIf ActiveTab = "V" Then
      Select Case lstSelect.ListIndex
         Case 0
            OpenHelpContext 5
         Case 1
            OpenHelpContext 1000
         Case 2
            OpenHelpContext 2000
         Case 3
            OpenHelpContext 3000
         Case 4
            OpenHelpContext 7000
         Case 5
            OpenHelpContext 4000
         Case 6
            OpenHelpContext 5000
         Case 7
            OpenHelpContext 6000
         Case Else
            b = 0
      End Select
   Else 'Functions
      Select Case lstSelect.ListIndex
         Case 0
            b = 0
         Case Else
            b = 0
      End Select
   End If
   MouseCursor 0
   
End Sub


Private Sub lstSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   lstSelect_KeyPress 13
   
   '    MouseCursor 13
   '    b = 1
   '    If ActiveTab = "E" Then
   '        Select Case lstSelect.ListIndex
   '            Case 0
   '                b = 0
   '            Case Else
   '                b = 0
   '        End Select
   '    ElseIf ActiveTab = "V" Then
   '        Select Case lstSelect.ListIndex
   '            Case 0
   '                OpenHelpContext 5
   '            Case 1
   '                OpenHelpContext 1000
   '            Case 2
   '                OpenHelpContext 2000
   '            Case 3
   '                OpenHelpContext 3000
   '            Case 4
   '                OpenHelpContext 7000
   '            Case 5
   '                OpenHelpContext 4000
   '            Case 6
   '                 OpenHelpContext 5000
   '            Case 7
   '                 OpenHelpContext 6000
   '            Case Else
   '                b = 0
   '        End Select
   '    Else 'Functions
   '        Select Case lstSelect.ListIndex
   '            Case 0
   '                b = 0
   '            Case Else
   '                b = 0
   '        End Select
   '    End If
   '    MouseCursor 0
   '
End Sub


Private Sub tab1_Click()
   bActiveTab(6) = Tab1.SelectedItem.Index
   Select Case bActiveTab(6)
      Case 1
         FillView
      Case Else
         FillView
   End Select
   
End Sub

Private Sub FillEdit()
   ActiveTab = "E"
   lstSelect.Clear
   If Secure.UserAdmnG6E = 1 Then
      'If User.Group6 Then
      lstSelect.AddItem "No Edit Functions"
      'Else
      '    lstSelect.Enabled = False
      '    lstSelect.AddItem "No Group Permissions"
      'End If
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub

Private Sub FillView()
   ActiveTab = "V"
   lstSelect.Clear
   If Secure.UserAdmnG6V = 1 Then
      'If User.Group6 Then
      lstSelect.AddItem "Fusion ERP"
      lstSelect.AddItem "Administration Help "
      lstSelect.AddItem "Sales Help"
      lstSelect.AddItem "Engineering Help "
      lstSelect.AddItem "Financial Accounting Help "
      lstSelect.AddItem "Production Control Help "
      lstSelect.AddItem "Inventory Control Help "
      lstSelect.AddItem "Quality Assurance Help "
      'Else
      '    lstSelect.Enabled = False
      '    lstSelect.AddItem "No Group Permissions"
      'End If
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub

Public Sub FillFunctions()
   ActiveTab = "F"
   lstSelect.Clear
   If Secure.UserAdmnG6F = 1 Then
      'If User.Group6 Then
      lstSelect.AddItem "No Functions"
      'Else
      '    lstSelect.Enabled = False
      '    lstSelect.AddItem "No Group Permissions"
      'End If
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub
