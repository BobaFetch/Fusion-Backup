VERSION 5.00
Begin VB.Form BookBLp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backlog By Sales Order Date"
   ClientHeight    =   5220
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5220
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BookBLp03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   79
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   25
      Left            =   4920
      TabIndex        =   62
      Top             =   3720
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   24
      Left            =   4680
      TabIndex        =   61
      Top             =   3720
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   23
      Left            =   4440
      TabIndex        =   60
      Top             =   3720
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   22
      Left            =   4200
      TabIndex        =   59
      Top             =   3720
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   21
      Left            =   3960
      TabIndex        =   58
      Top             =   3720
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   20
      Left            =   3720
      TabIndex        =   57
      Top             =   3720
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   19
      Left            =   3480
      TabIndex        =   56
      Top             =   3720
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   18
      Left            =   3240
      TabIndex        =   55
      Top             =   3720
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   17
      Left            =   3000
      TabIndex        =   54
      Top             =   3720
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   16
      Left            =   2760
      TabIndex        =   53
      Top             =   3720
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   15
      Left            =   2520
      TabIndex        =   52
      Top             =   3720
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   14
      Left            =   2280
      TabIndex        =   51
      Top             =   3720
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   13
      Left            =   2040
      TabIndex        =   50
      Top             =   3720
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   12
      Left            =   4920
      TabIndex        =   49
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   11
      Left            =   4680
      TabIndex        =   48
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   10
      Left            =   4440
      TabIndex        =   47
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   9
      Left            =   4200
      TabIndex        =   46
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   8
      Left            =   3960
      TabIndex        =   45
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   7
      Left            =   3720
      TabIndex        =   44
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   6
      Left            =   3480
      TabIndex        =   43
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   3240
      TabIndex        =   42
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   3000
      TabIndex        =   41
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   2760
      TabIndex        =   40
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   2520
      TabIndex        =   39
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   38
      Top             =   3240
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      BackColor       =   &H00000000&
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   37
      Top             =   3240
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.ComboBox cmbCde 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox cmbCls 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   27
      Top             =   4440
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   25
      Top             =   4200
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2040
      TabIndex        =   29
      Top             =   4680
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "BookBLp03a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "BookBLp03a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With Sales Orders"
      Top             =   1440
      Width           =   1555
   End
   Begin VB.PictureBox ReSize1 
      Height          =   480
      Left            =   6480
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   80
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "( Blank For All)"
      Height          =   285
      Index           =   13
      Left            =   5640
      TabIndex        =   78
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "( Blank For All)"
      Height          =   285
      Index           =   12
      Left            =   5640
      TabIndex        =   77
      Top             =   2520
      Width           =   2505
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "( Blank For All)"
      Height          =   285
      Index           =   11
      Left            =   5640
      TabIndex        =   76
      Top             =   2160
      Width           =   2505
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      Height          =   255
      Index           =   25
      Left            =   4920
      TabIndex        =   75
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      Height          =   255
      Index           =   24
      Left            =   4680
      TabIndex        =   74
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Index           =   23
      Left            =   4440
      TabIndex        =   73
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      Height          =   255
      Index           =   22
      Left            =   4200
      TabIndex        =   72
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      Height          =   255
      Index           =   21
      Left            =   3960
      TabIndex        =   71
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      Height          =   255
      Index           =   20
      Left            =   3720
      TabIndex        =   70
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      Height          =   255
      Index           =   19
      Left            =   3480
      TabIndex        =   69
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      Height          =   255
      Index           =   18
      Left            =   3240
      TabIndex        =   68
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      Height          =   255
      Index           =   17
      Left            =   3000
      TabIndex        =   67
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      Height          =   255
      Index           =   16
      Left            =   2760
      TabIndex        =   66
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      Height          =   255
      Index           =   15
      Left            =   2520
      TabIndex        =   65
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      Height          =   255
      Index           =   14
      Left            =   2280
      TabIndex        =   64
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      Height          =   255
      Index           =   13
      Left            =   2040
      TabIndex        =   63
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Types"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   36
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      Height          =   255
      Index           =   12
      Left            =   4920
      TabIndex        =   34
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   32
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      Height          =   255
      Index           =   10
      Left            =   4440
      TabIndex        =   30
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   28
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      Height          =   255
      Index           =   8
      Left            =   3960
      TabIndex        =   26
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   24
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   23
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   22
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   21
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   20
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   19
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   18
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   17
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Class"
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   15
      Top             =   2160
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   1
      Left            =   3480
      TabIndex        =   14
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Part Descriptions"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   13
      Tag             =   " "
      Top             =   4200
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "( Blank For All)"
      Height          =   288
      Index           =   4
      Left            =   5640
      TabIndex        =   12
      Top             =   1440
      Width           =   1548
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Desc"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   10
      Tag             =   " "
      Top             =   4440
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Comments"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Tag             =   " "
      Top             =   4680
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Dates"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1425
   End
End
Attribute VB_Name = "BookBLp03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/25/05 Changed dates and Options
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbcde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If Len(cmbCde) = 0 Then cmbCde = "ALL"
   
End Sub

Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 4)
   If Len(cmbCls) = 0 Then cmbCls = "ALL"
   
End Sub

Private Sub cmbCst_Click()
   GetCustomer
   If Len(cmbCde) = 0 Then cmbCst = "ALL"
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Len(cmbCst) = 0 Then cmbCst = "ALL"
   GetCustomer
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_GetCustomerSalesOrder"
   LoadComboBox cmbCst
   If Not bSqlRows Then
      lblNme = "*** No Customers With SO's Found ***"
   Else
      cmbCst = "ALL"
      GetCustomer
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductCodes
      FillProductClasses
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set BookBLp03a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim a As Byte
   Dim b As Byte
   Dim C As Byte
   Dim sCust As String
   Dim sBeg As String
   Dim sEnd As String
   Dim sCode As String
   Dim sClass As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   
   For b = 0 To 25
      If optTyp(b).Value = vbChecked Then C = 1
   Next
   If C = 0 Then
      MsgBox "No SO Types Selected.", vbInformation, Caption
      Exit Sub
   End If
   MouseCursor 13
   If Trim(cmbCst) = "" Then cmbCst = "ALL"
   If cmbCst <> "ALL" Then sCust = Compress(cmbCst)
   If Not IsDate(txtBeg) Then
      sBeg = "1995,01,01"
   Else
      sBeg = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEnd = "2024,12,31"
   Else
      sEnd = Format(txtEnd, "yyyy,mm,dd")
   End If
   
   On Error GoTo DiaErr1
'   SetMdiReportsize MdiSect
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MdiSect.Crw.Formulas(1) = "Includes='Customer(s) " & cmbCst & "." _
'                        & " From " & txtBeg & " Through " & txtEnd & "'"
'   MdiSect.Crw.Formulas(2) = "RequestBy = 'Requested By: " & sInitials & "'"
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDesc"
   aFormulaName.Add "ShowExDesc"
   aFormulaName.Add "ShowComments"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Customer(s) " & CStr(cmbCst & "." _
                        & " From " & txtBeg & " Through " & txtEnd) & "...'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optDsc.Value
   aFormulaValue.Add optExt.Value
   aFormulaValue.Add optCmt.Value
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("slebl03")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
'   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{CustTable.CUREF} LIKE '" & sCust & "*' " _
          & "AND {SohdTable.SODATE} In Date(" & sBeg & ") " _
          & "To Date(" & sEnd & ") " _
          & "AND {PartTable.PAPRODCODE} LIKE '" & sCode & "*' " _
          & "AND {PartTable.PACLASS} LIKE '" & sClass & "*' AND ("
   a = 65
   C = 0
   For b = 0 To 25
      If optTyp(b).Value = vbChecked Then
         If C = 1 Then
            sSql = sSql & "OR {SohdTable.SOTYPE}='" & Chr$(a) & "' "
         Else
            sSql = sSql & "{SohdTable.SOTYPE}='" & Chr$(a) & "' "
         End If
         C = 1
      End If
      a = a + 1
   Next
   sSql = sSql & ")"
   
  sSql = sSql & " and {SoitTable.ITCANCELED} = .000 and" _
              & " {SoitTable.ITPSNUMBER} = '' and " _
              & "{SoitTable.ITINVOICE} = .000 and " _
              & "{SoitTable.ITPSSHIPPED} = .000"

'   If optDsc.value = vbUnchecked Then
'      MdiSect.Crw.SectionFormat(0) = "DETAIL.0.0;F;;;"
'   Else
'      MdiSect.Crw.SectionFormat(0) = "DETAIL.0.0;T;;;"
'   End If
'   If optExt.value = vbUnchecked Then
'      MdiSect.Crw.SectionFormat(1) = "DETAIL.0.1;F;;;"
'   Else
'      MdiSect.Crw.SectionFormat(1) = "DETAIL.0.1;T;;;"
'   End If
'   If optCmt.value = vbUnchecked Then
'      MdiSect.Crw.SectionFormat(2) = "GROUPFTR.2.0;F;;;"
'      MdiSect.Crw.SectionFormat(3) = "GROUPFTR.2.1;F;;;"
'   Else
'      MdiSect.Crw.SectionFormat(2) = "GROUPFTR.2.0;T;;;"
'      MdiSect.Crw.SectionFormat(3) = "GROUPFTR.2.1;T;;;"
'   End If
'   MdiSect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim a As Byte
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   cmbCls = "ALL"
   cmbCde = "ALL"
   a = 5
   For b = 0 To 24
      optTyp(b).TabIndex = a
      a = a + 1
   Next
   optTyp(b).TabIndex = a
   
End Sub

Private Sub SaveOptions()
   Dim b As Byte
   Dim sOptions As String
   sOptions = Trim(txtEnd) & Trim(str(optExt.Value)) _
              & Trim(str(optDsc.Value)) & Trim(str(optCmt.Value))
   SaveSetting "Esi2000", "EsiSale", "bL03", Trim(sOptions)
   
   sOptions = ""
   For b = 0 To 25
      sOptions = sOptions & Trim$(optTyp(b).Value)
   Next
   SaveSetting "Esi2000", "EsiSale", "bl03b", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim b As Byte
   Dim sOptions As String
   Dim sOption2 As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "bL03", Trim(sOptions))
   If Len(sOptions) Then
      optExt = Mid(sOptions, 9, 1)
      optDsc = Mid(sOptions, 10, 1)
      optCmt = Mid(sOptions, 11, 1)
   Else
      optExt.Value = vbChecked
      optDsc.Value = vbChecked
      optCmt.Value = vbChecked
   End If
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 5)
   
   sOption2 = GetSetting("Esi2000", "EsiSale", "bl03b", Trim(sOption2))
   On Error GoTo DiaErr1
   If Len(Trim(sOption2)) > 0 Then
      For b = 0 To 24
         optTyp(b).Value = Val(Mid$(sOption2, b + 1, 1))
      Next
      optTyp(b).Value = Val(Mid$(sOption2, b + 1, 1))
   End If
   Exit Sub
   
DiaErr1:
   For b = 0 To 24
      optTyp(b).Value = vbChecked
   Next
   optTyp(b).Value = vbChecked
   
End Sub

Private Sub lblNme_Change()
   If Left(lblNme, 9) = "*** No Cu" Then
      lblNme.ForeColor = ES_RED
   Else
      lblNme.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub optTyp_GotFocus(Index As Integer)
   lblAlp(Index).BorderStyle = vbFixedSingle
   
End Sub


Private Sub optTyp_LostFocus(Index As Integer)
   Dim b As Byte
   For b = 0 To 24
      lblAlp(b).BorderStyle = 0
   Next
   lblAlp(b).BorderStyle = 0
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub



Private Sub GetCustomer()
   Dim RdoCst As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetCustomerBasics '" & Compress(cmbCst) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
   If bSqlRows Then
      With RdoCst
         cmbCst = "" & Trim(.Fields(1))
         If Len(cmbCst) > 3 Then
            lblNme = "" & Trim(.Fields(2))
         Else
            lblNme = "*** Range Of Customers Selected ***"
         End If
         ClearResultSet RdoCst
      End With
   Else
      lblNme = "*** Range Of Customers Selected ***"
   End If
   Set RdoCst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcustomer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Trim(txtEnd) <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub


Private Sub z1_Click(Index As Integer)
   optDsc.SetFocus
   
End Sub
