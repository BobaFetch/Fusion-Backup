VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form zGr3Book 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bookings And Backlog"
   ClientHeight    =   4770
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4920
   ControlBox      =   0   'False
   Icon            =   "Gr3Book.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Gr3Book.frx":08CA
   ScaleHeight     =   4770
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstSelect 
      Height          =   3375
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   4065
   End
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
            Picture         =   "Gr3Book.frx":1194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr3Book.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr3Book.frx":1B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr3Book.frx":1CB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3780
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "    Close (Escape)    "
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "Gr3Book.frx":1DE0
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   250
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
      FormDesignHeight=   4770
      FormDesignWidth =   4920
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   4215
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7435
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2160
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
            Picture         =   "Gr3Book.frx":258E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr3Book.frx":2A3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gr3Book.frx":2EDE
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
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "zGr3Book"
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
   cUR.CurrentGroup = "Book"
   
End Sub

Private Sub Form_Deactivate()
   On Error Resume Next
   Hide
   
End Sub

Private Sub Form_Load()
   ActiveTab = "E"
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   'If bSecSet = 1 Then User.Group3 = 1
   '9/14/06 allow for old Permissions
   If Secure.UserSaleG3E = 1 Then
      Secure.UserSaleG3V = 1
      Secure.UserSaleG3F = 1
   End If
   FillEdit
   If bCustomerGroups(3) = 0 Then
      GroupTabPermissions Me
      'User.Group3 = 1
   End If
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set zGr3Book = Nothing
   
End Sub




Private Sub lstSelect_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then Exit Sub
   MouseCursor 13
   b = 1
   If ActiveTab = "E" Then
      b = 0
   ElseIf ActiveTab = "V" Then
      Select Case lstSelect.ListIndex
         Case 0
            BookBKp01a.Show
         Case 1
            BookBKp02a.Show
         Case 2
            BookBKp03a.Show
         Case 3
            BookBKp04a.Show
         Case 4
            BookBKp05a.Show
         Case 5
            BookBKp06a.Show
         Case 6
            BookBKp18a.Show
            
            'backlog
            
         Case 7
            BookBLp01a.Show
         Case 8
            BookBLp02a.Show
         Case 9
            BookBLp03a.Show
         Case 10
            BookBLp04a.Show
         Case 11
            BookBLp05a.Show
         Case 12
            BookBLp06a.Show
         Case 13
            BookBLp09a.Show
         Case 14
            BookBLp07a.Show
         Case 15
            BookBLp08a.Show
         Case 16
            PackPSp14a.Show
         Case Else
            b = 0
      End Select
   Else
      b = 0
   End If
   If b = 1 Then Hide Else MouseCursor 0
   
End Sub


Private Sub lstSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   
   lstSelect_KeyPress 13
   
End Sub


Private Sub tab1_Click()
   bActiveTab(3) = Tab1.SelectedItem.Index
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
   If Secure.UserSaleG3E = 1 Then
      lstSelect.AddItem "No Functions"
      lstSelect.Enabled = False
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub

Private Sub FillView()
   ActiveTab = "V"
   lstSelect.Clear
   If Secure.UserSaleG3V = 1 Then
      lstSelect.AddItem "Bookings By Order Date (Chart) "
      lstSelect.AddItem "Bookings By Part Number "
      lstSelect.AddItem "Bookings By Salesperson (Chart) "
      lstSelect.AddItem "Bookings By Division, Class and Code (Chart) "
      lstSelect.AddItem "Bookings By Business Unit, Code (Chart) "
      lstSelect.AddItem "Bookings By Customer (Chart) "
      lstSelect.AddItem "Part Availability Report"
      lstSelect.AddItem "Backlog By Part With Prices "
      lstSelect.AddItem "Backlog By Scheduled Date "
      lstSelect.AddItem "Backlog By Sales Order Date "
      lstSelect.AddItem "Backlog By Salesperson "
      lstSelect.AddItem "Backlog By Customer "
      lstSelect.AddItem "Backlog By Part Number, Current Month"
      lstSelect.AddItem "Backlog By Part Number"
      lstSelect.AddItem "Backlog, 12 Month By Division"
      lstSelect.AddItem "Customer Backlog Analysis"
      lstSelect.AddItem "Inventory Available To Ship"
      lstSelect.Enabled = True
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub

Public Sub FillFunctions()
   ActiveTab = "F"
   lstSelect.Clear
   If Secure.UserSaleG3F = 1 Then
      '        lstSelect.AddItem "Backlog By Part With Prices "
      '        lstSelect.AddItem "Backlog By Scheduled Date "
      '        lstSelect.AddItem "Backlog By Sales Order Date "
      '        lstSelect.AddItem "Backlog By Salesperson "
      '        lstSelect.AddItem "Backlog By Customer "
      '        lstSelect.AddItem "Backlog By Part Number, Current Month"
      '        lstSelect.AddItem "Backlog, 12 Month By Division"
      '        lstSelect.AddItem "Customer Backlog Analysis"
      '        lstSelect.AddItem "Inventory Available To Ship"
      lstSelect.AddItem "No Functions"
      lstSelect.Enabled = False
   Else
      lstSelect.Enabled = False
      lstSelect.AddItem "No User Permissions"
   End If
   
End Sub
