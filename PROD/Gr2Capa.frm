VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form zGr2Capa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capacity Requirements Planning"
   ClientHeight    =   5010
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4920
   ControlBox      =   0   'False
   Icon            =   "Gr2Capa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Gr2Capa.frx":08CA
   ScaleHeight     =   5010
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   4920
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
            Picture         =   "Gr2Capa.frx":1194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr2Capa.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr2Capa.frx":1B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr2Capa.frx":1CB2
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
      Picture         =   "Gr2Capa.frx":1DE0
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7858
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
      Top             =   4575
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
      Left            =   1680
      Top             =   4800
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
            Picture         =   "Gr2Capa.frx":258E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr2Capa.frx":2A3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr2Capa.frx":2EDE
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
      Top             =   4680
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "zGr2Capa"
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
   cUR.CurrentGroup = "Capa"
   
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
   Set zGr2Capa = Nothing
   
End Sub




Private Sub lstSelect_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   If ActiveTab = "E" Then
      Select Case lstSelect.ListIndex
         Case 0
            CapaCPe01a.Show
         Case 1
            CapaCPe02a.Show
         Case 2
            CapaCPe03a.Show
         Case 3
            CapaCPe04a.Show
         Case 4
            CapaCPe05a.Show
            '11/7/03
            'Case 5
            '   'diaCswcn.Show
            'Case 6
            '    diaCcwcn.Show
         Case Else
            b = 0
      End Select
   ElseIf ActiveTab = "V" Then
      Dim ssrs As New SSRSViewer
      Select Case lstSelect.ListIndex
         Case 0
            CapaCPp01a.Show
         Case 1
            CapaCPp02a.Show
         Case 2
            CapaCPp03a.Show
         Case 3
            CapaCPp04a.Show
         Case 4
            CapaCPp05a.Show
         Case 5
            CapaCPp06a.Show
         Case 6
            CapaCPp07a.Show
         Case 7
            CapaCPp08a.Show
         Case 8
            CapaCPp09a.Show
         Case 9
            CapaCPp10a.Show
         Case 10
            CapaCPp11a.Show
         Case 11
            ssrs.report = "Capacity vs Load"
            ssrs.Show
         Case 12
            ssrs.report = "Load Details"
            ssrs.Show
         Case Else
            b = 0
      End Select
   Else 'Functions
      Select Case lstSelect.ListIndex
         Case 0
            CapaCPf01a.Show
         Case 1
            CapaCPf02a.Show
         Case 2
            CapaCPf03a.Show
         Case 3
            CapaCPf04a.Show
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
   '             Case 0
   '                 CapaCPe01a.Show
   '             Case 1
   '                 CapaCPe02a.Show
   '             Case 2
   '                 CapaCPe03a.Show
   '             Case 3
   '                 CapaCPe04a.Show
   '             Case 4
   '                 CapaCPe05a.Show
   '            '11/7/03
   '            'Case 5
   '            '   'diaCswcn.Show
   '            'Case 6
   '            '    diaCcwcn.Show
   '            Case Else
   '                b = 0
   '        End Select
   '    ElseIf ActiveTab = "V" Then
   '        Select Case lstSelect.ListIndex
   '            Case 0
   '                CapaCPp01a.Show
   '            Case 1
   '                CapaCPp02a.Show
   '            Case 2
   '                CapaCPp03a.Show
   '            Case 3
   '                CapaCPp04a.Show
   '            Case 4
   '                CapaCPp05a.Show
   '            Case 5
   '                CapaCPp06a.Show
   '            Case 6
   '                CapaCPp07a.Show
   '            Case 7
   '                CapaCPp08a.Show
   '            Case 8
   '                CapaCPp09a.Show
   '            Case 9
   '                CapaCPp10a.Show
   '            Case Else
   '                b = 0
   '        End Select
   '    Else 'Functions
   '        Select Case lstSelect.ListIndex
   '            Case 0
   '                CapaCPf01a.Show
   '            Case 1
   '                CapaCPf02a.Show
   '            Case 2
   '                CapaCPf03a.Show
   '            Case Else
   '                b = 0
   '        End Select
   '    End If
   '
   '    If b = 1 Then Hide Else MouseCursor 0
   '
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
   If Secure.UserProdG2E = 1 Then
      'If User.Group2 Then
      lstSelect.AddItem "Work Centers "
      lstSelect.AddItem "Shops"
      lstSelect.AddItem "Work Center Calendar"
      lstSelect.AddItem "Company Calendar"
      lstSelect.AddItem "Company Calendar Template "
      ' 9/22/03 Removed
      ' lstSelect.AddItem "Work Center Schedule "
      ' lstSelect.AddItem "Work Center Chart"
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
   If Secure.UserProdG2V = 1 Then
      'If User.Group2 Then
      lstSelect.AddItem "Work Centers Report"
      lstSelect.AddItem "Shop Information"
      lstSelect.AddItem "Work Center Load (Includes Chart)"
      lstSelect.AddItem "Late Manufacturing Orders "
      lstSelect.AddItem "List Of Work Center Calendars "
      lstSelect.AddItem "Shop Load By Work Center (Includes Chart) "
      lstSelect.AddItem "Work Centers Without Calendars"
      lstSelect.AddItem "Capacity And Load By Shops, Work Centers"
      lstSelect.AddItem "Work Center Used On"
      lstSelect.AddItem "Work Center Load Analysis"
      lstSelect.AddItem "Work Center Performance Report"
      'lstSelect.AddItem "Capacity Plan"
      'lstSelect.AddItem "MO Schedule"
      'lstSelect.AddItem "SO Load By MO"
      'lstSelect.AddItem "SO Load By WC"
      'lstSelect.AddItem "MO Load By WC"
      'lstSelect.AddItem "Milestones Not Met"
      'lstSelect.AddItem "OPs Scheduled To Finite Schedule"
      'lstSelect.AddItem "Load By Shop By Part Class"
      lstSelect.AddItem "Capacity vs Load Report (SSRS)"
      lstSelect.AddItem "Load Details Report (SSRS)"
      
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
   If Secure.UserProdG2F = 1 Then
      'If User.Group2 Then
      lstSelect.AddItem "Reassign Shops and Work Centers "
      lstSelect.AddItem "Delete Shops"
      lstSelect.AddItem "Delete Work Centers"
      lstSelect.AddItem "OP's completed by Work Centers"
      'lstSelect.AddItem "Finite Schedule All MO's"
      'lstSelect.AddItem "Schedule MO's"
      'lstSelect.AddItem "Reschedule MO's"
      'lstSelect.AddItem "Generate Capacity Plan"
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
