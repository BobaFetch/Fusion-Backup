VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form zGr1Rout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Routings"
   ClientHeight    =   4515
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4920
   ControlBox      =   0   'False
   Icon            =   "Gr1Rout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Gr1Rout.frx":08CA
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
            Picture         =   "Gr1Rout.frx":1194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Rout.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Rout.frx":1B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Rout.frx":1CB2
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
      Picture         =   "Gr1Rout.frx":1DE0
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
      Left            =   2160
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
            Picture         =   "Gr1Rout.frx":258E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Rout.frx":2A3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr1Rout.frx":2EDE
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
Attribute VB_Name = "zGr1Rout"
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
   cUR.CurrentGroup = "Rout"
   
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
   Set zGr1Rout = Nothing
   
End Sub




Private Sub lstSelect_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   If ActiveTab = "E" Then
      Select Case lstSelect.ListIndex
         Case 0
            RoutRTe01a.Caption = "Routing"
            RoutRTe01a.Show
         Case 1
            RoutRTe02a.Show
         Case 2
            RoutRTe03a.Show
         Case 3
            RoutRTe04a.Show
         Case 4
            RoutRTe05a.Show
         Case 5
            RoutRTe06a.Show
         Case Else
            b = 0
      End Select
   ElseIf ActiveTab = "V" Then
      Select Case lstSelect.ListIndex
         Case 0
            RoutRTp01a.Show
         Case 1
            RoutRTp02a.Show
         Case 2
            RoutRTp03a.Show
         Case 3
            RoutRTp04a.Show
         Case 4
            RoutRTp05a.Show
         Case 5
            RoutRTp06a.Show
         Case 6
            RoutRTp07a.Show
         Case 7
            RoutRTp08a.Show
         Case Else
            b = 0
      End Select
   Else 'Functions
      Select Case lstSelect.ListIndex
         Case 0
            RoutRTf01a.Show
         Case 1
            RoutRTf02a.Show
         Case 2
            RoutRTf03a.Show
         Case 3
            RoutRTf04a.Show
         Case 4
            RoutRTf05a.Show
         Case 5
            RoutRTf06a.Show
         Case 6
            RoutRTf07a.Show
         Case 7
            RoutRTf08a.Show			
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
   '                RoutRTe01a.Caption = "Routing"
   '                RoutRTe01a.Show
   '            Case 1
   '                RoutRTe02a.Show
   '            Case 2
   '                RoutRTe03a.Show
   '            Case 3
   '                RoutRTe04a.Show
   '            Case 4
   '                RoutRTe05a.Show
   '            Case 5
   '                RoutRTe06a.Show
   '            Case Else
   '                b = 0
   '        End Select
   '    ElseIf ActiveTab = "V" Then
   '        Select Case lstSelect.ListIndex
   '            Case 0
   '                RoutRTp01a.Show
   '            Case 1
   '                RoutRTp02a.Show
   '            Case 2
   '                RoutRTp03a.Show
   '            Case 3
   '                RoutRTp04a.Show
   '            Case 4
   '                RoutRTp05a.Show
   '            Case 5
   '                RoutRTp06a.Show
   '            Case 6
   '                RoutRTp07a.Show
   '            Case Else
   '                b = 0
   '        End Select
   '    Else 'Functions
   '        Select Case lstSelect.ListIndex
   '            Case 0
   '                RoutRTf01a.Show
   '            Case 1
   '                RoutRTf02a.Show
   '            Case 2
   '                RoutRTf03a.Show
   '            Case 3
   '                RoutRTf04a.Show
   '            Case 4
   '                RoutRTf05a.Show
   '            Case 5
   '                RoutRTf06a.Show
   '            Case 6
   '                RoutRTf07a.Show
   '            Case Else
   '                b = 0
   '        End Select
   '    End If
   '
   '    If b = 1 Then Hide Else MouseCursor 0
   '
End Sub


Private Sub tab1_Click()
   bActiveTab(1) = Tab1.SelectedItem.Index
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
   If Secure.UserEngrG1E = 1 Then
      'If User.Group1 Then
      lstSelect.AddItem "Routings "
      lstSelect.AddItem "Routing Assignments "
      lstSelect.AddItem "Default Routing Assignments "
      lstSelect.AddItem "Revise A Routing Number "
      lstSelect.AddItem "Routing Operation Library "
      lstSelect.AddItem "Routing Operation Pictures"
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
   If Secure.UserEngrG1V = 1 Then
      'If User.Group1 Then
      lstSelect.AddItem "Routing"
      lstSelect.AddItem "Routings By Routing Number "
      lstSelect.AddItem "Routings By Part Number "
      lstSelect.AddItem "Routing Used On "
      lstSelect.AddItem "Routings By Service Part Number"
      lstSelect.AddItem "Routings By Engineer"
      lstSelect.AddItem "Pictures By Routing Operation"
      lstSelect.AddItem "Labor Efficiency"
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
   If Secure.UserEngrG1F = 1 Then
      'If User.Group1 Then
      lstSelect.AddItem "Delete A Routing"
      lstSelect.AddItem "Copy A Routing"
      lstSelect.AddItem "Reassign Shops And Work Centers"
      lstSelect.AddItem "Merge Routings"
      lstSelect.AddItem "Reorganize Operations"
      lstSelect.AddItem "Delete A Routing Library Operation"
      lstSelect.AddItem "Copy A Routing From A Manufacturing Order"
      lstSelect.AddItem "Copy A Routing Operation Comments"
	  'lstSelect.AddItem "Update A Routing From An MO"
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
