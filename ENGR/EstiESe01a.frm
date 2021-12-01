VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESe01a 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Qwik Bid"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ControlBox      =   0   'False
   HelpContextID   =   3501
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   87
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "N&ew"
      Height          =   285
      Left            =   5200
      TabIndex        =   83
      TabStop         =   0   'False
      ToolTipText     =   "Add A New Estimate"
      Top             =   480
      Width           =   1000
   End
   Begin VB.CheckBox optFrom 
      Caption         =   "from"
      Height          =   255
      Left            =   3840
      TabIndex        =   81
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox optPart 
      Caption         =   "Parts"
      Height          =   195
      Left            =   360
      TabIndex        =   80
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox optCom 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4080
      TabIndex        =   77
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdCom 
      Caption         =   "Unc&omplete"
      Height          =   285
      Left            =   5200
      TabIndex        =   76
      TabStop         =   0   'False
      ToolTipText     =   "Mark This Bid As Complete"
      Top             =   820
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Height          =   5400
      Index           =   1
      Left            =   6480
      TabIndex        =   64
      Top             =   1080
      Width           =   6200
      Begin VB.CommandButton cmdComments 
         DisabledPicture =   "EstiESe01a.frx":07AE
         DownPicture     =   "EstiESe01a.frx":1120
         Height          =   350
         Left            =   1920
         Picture         =   "EstiESe01a.frx":1A92
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Standard Comments"
         Top             =   2520
         Width           =   350
      End
      Begin VB.CommandButton cmdDis 
         Caption         =   "&Discounts"
         Height          =   285
         Left            =   5160
         TabIndex        =   82
         TabStop         =   0   'False
         ToolTipText     =   "Volume Discounts"
         Top             =   240
         Width           =   875
      End
      Begin VB.TextBox txtByr 
         Height          =   285
         Left            =   2400
         TabIndex        =   25
         Tag             =   "2"
         ToolTipText     =   "Profit Calculated On The Subtotal Of All (Including G&A)"
         Top             =   2040
         Width           =   2595
      End
      Begin VB.TextBox txtEst 
         Height          =   285
         Left            =   2400
         TabIndex        =   24
         Tag             =   "2"
         ToolTipText     =   "Profit Calculated On The Subtotal Of All (Including G&A)"
         Top             =   1680
         Width           =   2595
      End
      Begin VB.ComboBox txtFst 
         Height          =   315
         Left            =   2400
         TabIndex        =   23
         Tag             =   "4"
         Top             =   1320
         Width           =   1250
      End
      Begin VB.CommandButton cmdLst 
         Caption         =   "<<< &Back"
         Height          =   255
         Left            =   5200
         TabIndex        =   65
         Top             =   4440
         Width           =   875
      End
      Begin VB.TextBox txtCmt 
         Height          =   885
         Left            =   2400
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Tag             =   "2"
         Top             =   2520
         Width           =   3435
      End
      Begin VB.TextBox txtTlg 
         Height          =   285
         Left            =   2400
         TabIndex        =   22
         Tag             =   "1"
         Top             =   960
         Width           =   915
      End
      Begin VB.TextBox txtEpd 
         Height          =   285
         Left            =   4800
         TabIndex        =   27
         Tag             =   "1"
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox txtSet 
         Height          =   285
         Left            =   2400
         TabIndex        =   21
         Tag             =   "1"
         Top             =   600
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Buyer"
         Height          =   255
         Index           =   33
         Left            =   240
         TabIndex        =   75
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Estimator"
         Height          =   255
         Index           =   32
         Left            =   240
         TabIndex        =   74
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bid Total"
         Height          =   255
         Index           =   30
         Left            =   2160
         TabIndex        =   73
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label lblBid 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   72
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments:"
         Height          =   255
         Index           =   24
         Left            =   240
         TabIndex        =   71
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "First Delivery Date"
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   70
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Additional Bid Costs (One Time Charges):"
         Height          =   252
         Index           =   22
         Left            =   240
         TabIndex        =   69
         Top             =   360
         Width           =   4092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Extra Setup Charges"
         Height          =   255
         Index           =   21
         Left            =   480
         TabIndex        =   68
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Expediting"
         Height          =   255
         Index           =   20
         Left            =   3480
         TabIndex        =   67
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tooling Charges"
         Height          =   255
         Index           =   19
         Left            =   480
         TabIndex        =   66
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   6240
      Top             =   120
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Tag             =   "4"
      Top             =   720
      Width           =   1250
   End
   Begin VB.CheckBox optSle 
      Caption         =   "View Sales"
      Height          =   255
      Left            =   2600
      TabIndex        =   53
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbBid 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "Select Or Enter A Bid Number (Contains Qwik Bids)"
      Top             =   680
      Width           =   975
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      ToolTipText     =   "Class A-Z"
      Top             =   680
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   5388
      Index           =   0
      Left            =   0
      TabIndex        =   31
      Top             =   1080
      Width           =   6200
      Begin VB.ComboBox cmbPrt 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   120
         Width           =   3075
      End
      Begin VB.ComboBox txtDue 
         Height          =   315
         Left            =   3960
         TabIndex        =   10
         Tag             =   "4"
         ToolTipText     =   "Due Date"
         Top             =   1200
         Width           =   1250
      End
      Begin VB.CommandButton cmdPrt 
         Height          =   315
         Left            =   5600
         Picture         =   "EstiESe01a.frx":2094
         Style           =   1  'Graphical
         TabIndex        =   79
         TabStop         =   0   'False
         ToolTipText     =   "New Part Numbers Find A Part Number (F2 At Part Number)"
         Top             =   135
         Width           =   350
      End
      Begin VB.TextBox txtScr 
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Tag             =   "1"
         ToolTipText     =   "Scrap And Reduction"
         Top             =   3360
         Width           =   912
      End
      Begin VB.TextBox txtGna 
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Tag             =   "1"
         ToolTipText     =   "General And Administrative Expense"
         Top             =   3720
         Width           =   915
      End
      Begin VB.ComboBox cmbRfq 
         Height          =   315
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   8
         Tag             =   "3"
         ToolTipText     =   "Select RFQ From List"
         Top             =   480
         Width           =   2040
      End
      Begin VB.CommandButton cmdVew 
         DownPicture     =   "EstiESe01a.frx":252F
         Height          =   315
         Left            =   5200
         Picture         =   "EstiESe01a.frx":2A09
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Show Sales Orders For Parts"
         Top             =   135
         Width           =   350
      End
      Begin VB.CommandButton cmdNxt 
         Caption         =   "&Next >>>"
         Height          =   255
         Left            =   5200
         TabIndex        =   28
         Top             =   4440
         Width           =   875
      End
      Begin VB.ComboBox cmbCst 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   7
         Tag             =   "3"
         ToolTipText     =   "Select A Customer"
         Top             =   480
         Width           =   1555
      End
      Begin VB.TextBox txtPrt 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Tag             =   "3"
         ToolTipText     =   "Enter A Part Or Click The ""Look Up"""
         Top             =   135
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.CommandButton cmdFnd 
         DownPicture     =   "EstiESe01a.frx":2EE3
         Height          =   315
         Left            =   4800
         Picture         =   "EstiESe01a.frx":3225
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Find A Part Number Find A Part Number (F4 At Part Number)"
         Top             =   135
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Tag             =   "1"
         ToolTipText     =   "Unit Costs"
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox txtHrs 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Tag             =   "1"
         ToolTipText     =   "Unit Costs"
         Top             =   1608
         Width           =   915
      End
      Begin VB.TextBox txtMtl 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Tag             =   "1"
         Top             =   2280
         Width           =   915
      End
      Begin VB.TextBox txtOsp 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Tag             =   "1"
         ToolTipText     =   "Unit Costs"
         Top             =   3000
         Width           =   915
      End
      Begin VB.TextBox txtPrf 
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Tag             =   "1"
         ToolTipText     =   "Profit Of Sale"
         Top             =   4080
         Width           =   915
      End
      Begin VB.TextBox txtPrc 
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "1"
         ToolTipText     =   "Locked TextBoxes"
         Top             =   4440
         Width           =   915
      End
      Begin VB.TextBox txtRte 
         Height          =   285
         Left            =   3960
         TabIndex        =   12
         Tag             =   "1"
         ToolTipText     =   "Unit Costs"
         Top             =   1608
         Width           =   915
      End
      Begin VB.TextBox txtCm1 
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Tag             =   "2"
         ToolTipText     =   "Description"
         Top             =   2280
         Width           =   3075
      End
      Begin VB.TextBox txtCm2 
         Height          =   285
         Left            =   2760
         TabIndex        =   16
         Tag             =   "2"
         ToolTipText     =   "Description"
         Top             =   3000
         Width           =   3075
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Estimator"
         Height          =   252
         Index           =   36
         Left            =   240
         TabIndex        =   89
         Top             =   4920
         Width           =   1332
      End
      Begin VB.Label Estimator 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   1680
         TabIndex        =   88
         Top             =   4920
         Width           =   2292
      End
      Begin VB.Label lblNme 
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   1680
         TabIndex        =   86
         Top             =   840
         Width           =   3852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Due"
         Height          =   252
         Index           =   34
         Left            =   3300
         TabIndex        =   85
         Top             =   1200
         Width           =   996
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   252
         Index           =   18
         Left            =   2640
         TabIndex        =   63
         Top             =   3360
         Width           =   252
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reduction (Scrap)"
         Height          =   252
         Index           =   17
         Left            =   240
         TabIndex        =   62
         Top             =   3360
         Width           =   2004
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bid Total"
         Height          =   252
         Index           =   29
         Left            =   2760
         TabIndex        =   58
         Top             =   4440
         Width           =   1332
      End
      Begin VB.Label lblBid 
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Index           =   0
         Left            =   3720
         TabIndex        =   57
         Top             =   4440
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "G&A"
         Height          =   252
         Index           =   28
         Left            =   240
         TabIndex        =   56
         Top             =   3720
         UseMnemonic     =   0   'False
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   252
         Index           =   27
         Left            =   2640
         TabIndex        =   55
         Top             =   3720
         Width           =   252
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "RFQ"
         Height          =   255
         Index           =   25
         Left            =   3300
         TabIndex        =   52
         Top             =   480
         Width           =   1000
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   252
         Index           =   15
         Left            =   2640
         TabIndex        =   50
         Top             =   4080
         Width           =   252
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   49
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   48
         Top             =   135
         Width           =   1575
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Labor:"
         Height          =   252
         Index           =   3
         Left            =   240
         TabIndex        =   47
         Top             =   1440
         Width           =   1332
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Material:"
         Height          =   252
         Index           =   4
         Left            =   240
         TabIndex        =   46
         Top             =   2280
         Width           =   1332
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Outside Services"
         Height          =   252
         Index           =   5
         Left            =   240
         TabIndex        =   45
         Top             =   3000
         Width           =   1692
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Profit"
         Height          =   252
         Index           =   6
         Left            =   240
         TabIndex        =   44
         Top             =   4080
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         Height          =   252
         Index           =   7
         Left            =   240
         TabIndex        =   43
         Top             =   4440
         Width           =   1332
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   252
         Index           =   8
         Left            =   240
         TabIndex        =   42
         Top             =   1200
         Width           =   1332
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hours"
         Height          =   252
         Index           =   9
         Left            =   600
         TabIndex        =   41
         Top             =   1632
         Width           =   1332
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   252
         Index           =   10
         Left            =   3300
         TabIndex        =   40
         Top             =   1632
         Width           =   996
      End
      Begin VB.Label lblLbr 
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   39
         Top             =   1920
         Width           =   912
      End
      Begin VB.Label lblFoh 
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   1680
         TabIndex        =   38
         Top             =   1920
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Overhead"
         Height          =   252
         Index           =   11
         Left            =   600
         TabIndex        =   37
         Top             =   1920
         Width           =   1332
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   252
         Index           =   12
         Left            =   3300
         TabIndex        =   36
         Top             =   1920
         Width           =   996
      End
      Begin VB.Label lblBrd 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   35
         Top             =   2280
         Width           =   915
      End
      Begin VB.Label lblMat 
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   3960
         TabIndex        =   34
         Top             =   2640
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Burden"
         Height          =   252
         Index           =   13
         Left            =   600
         TabIndex        =   33
         Top             =   2640
         Width           =   1332
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   252
         Index           =   14
         Left            =   3000
         TabIndex        =   32
         Top             =   2640
         Width           =   912
      End
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View Parts"
      Height          =   255
      Left            =   1400
      TabIndex        =   30
      Top             =   100
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5200
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   0
      Width           =   1000
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   1080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6645
      FormDesignWidth =   6255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Complete"
      Height          =   255
      Index           =   35
      Left            =   3240
      TabIndex        =   78
      ToolTipText     =   "Bid Is Marked Complete"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Form Width 6340 closed, 12000 open"
      Height          =   252
      Index           =   16
      Left            =   6840
      TabIndex        =   61
      Top             =   120
      Visible         =   0   'False
      Width           =   4152
   End
   Begin VB.Label lblNxt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   60
      Top             =   360
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Next Estimate"
      Height          =   255
      Index           =   31
      Left            =   240
      TabIndex        =   59
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bid Date"
      Height          =   255
      Index           =   26
      Left            =   3240
      TabIndex        =   54
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate Number"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   51
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "EstiESe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'2/14/06 Changed GetRates firing (Activate)
'2/14/06 Added permissions to allow overwriting
'5/23/06 Changed default BIDDATE to ES_SYSDATE
'8/30/06 Added sCurrEstimator and label Estimator
Option Explicit
Dim bOnLoad As Byte
Dim bCanceled As Byte

Dim bFromNew As Byte
Dim bGoodCust As Byte
Dim bGoodPart As Byte

Dim iBids As Integer

'General defaults
Dim cFoh As Currency
Dim cGna As Currency
Dim cMBurden As Currency
Dim cProfit As Currency
Dim cRate As Currency
Dim cScrap As Currency

'Bid stuff
Dim cBFoh As Currency
Dim cBGna As Currency
Dim cBprofit As Currency
Dim cBScrap As Currency

Dim cBMbrdRate As Currency
Dim cBFohRate As Currency
Dim cBGnaRate As Currency
Dim cBProfitRate As Currency
Dim cBScrapRate As Currency

Dim sEstimator As String

Dim lBids(300) As Long

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub UpdateDiscounts()
   Dim cPer As Currency
   Dim cPrc As Currency
   On Error GoTo DiaErr1
   With RdoBid
      '.Edit
      cPer = (!BIDQTYDISC1 / 100)
      cPrc = (1 - cPer) * Val(txtPrc)
      !BIDQTYPRICE1 = cPrc
      
      cPer = (!BIDQTYDISC2 / 100)
      cPrc = (1 - cPer) * Val(txtPrc)
      !BIDQTYPRICE2 = cPrc
      
      cPer = (!BIDQTYDISC3 / 100)
      cPrc = (1 - cPer) * Val(txtPrc)
      !BIDQTYPRICE3 = cPrc
      
      cPer = (!BIDQTYDISC4 / 100)
      cPrc = (1 - cPer) * Val(txtPrc)
      !BIDQTYPRICE4 = cPrc
      
      cPer = (!BIDQTYDISC5 / 100)
      cPrc = (1 - cPer) * Val(txtPrc)
      !BIDQTYPRICE5 = cPrc
      
      cPer = (!BIDQTYDISC6 / 100)
      cPrc = (1 - cPer) * Val(txtPrc)
      !BIDQTYPRICE6 = cPrc
      .Update
   End With
   Exit Sub
   
DiaErr1:
   sProcName = "updatedisc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbBid_Click()
   bGoodBid = GetThisQBid(True)
   
End Sub

Private Sub cmbBid_LostFocus()
   cmbBid = CheckLen(cmbBid, 6)
   cmbBid = Format(Abs(Val(cmbBid)), "000000")
   If bCanceled Then Exit Sub
   bGoodBid = GetThisQBid(False)
   If bGoodBid = 0 Then
      MsgBox "That Bid Has Not Been Recorded. Select New To Add.", _
         vbInformation, Caption
      cmbBid = cmbBid.List(0)
   End If
   
End Sub


Private Sub cmbCls_Change()
   If Len(cmbCls) > 1 Then cmbCls = Left(cmbCls, 1)
   
End Sub

Private Sub cmbCls_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbCls = CheckLen(cmbCls, 1)
   For iList = 0 To cmbCls.ListCount - 1
      If cmbCls.List(iList) = cmbCls Then b = 1
   Next
   If b = 0 Then cmbCls = "Q"
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDPRE = Trim(cmbCls)
      RdoBid.Update
   End If
   
End Sub


Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
   FillCustomerRFQs Me, cmbCst, True
   
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   bGoodCust = GetBidCustomer(Me, cmbCst)
   If bGoodCust Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDCUST = Compress(cmbCst)
      If bGoodPart Then
         RdoBid!BIDLOCKED = 0
      Else
         RdoBid!BIDLOCKED = 1
      End If
      RdoBid.Update
   End If
   '        If bGoodCust = 0 Or bGoodPart = 0 Then
   '            MsgBox "Bids Without A Valid Customer And Valid " _
   '                & "Part Number Will Not Be Saved.", _
   '                    vbInformation, Caption
   '        End If
   FillCustomerRFQs Me, cmbCst, True
   
End Sub


Private Sub cmbPrt_Change()
      txtPrt = cmbPrt
      
End Sub


Private Sub cmbPrt_Click()
      txtPrt = cmbPrt
      
End Sub
Private Sub cmbRfq_Click()
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDRFQ = cmbRfq
      RdoFull.Update
   End If
   
End Sub

Private Sub cmbRfq_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbRfq = CheckLen(cmbRfq, 14)
   If Trim(cmbRfq) = "" Then cmbRfq = "NONE"
   If cmbRfq <> "NONE" Then
      For iList = 0 To cmbRfq.ListCount - 1
         If cmbRfq = cmbRfq.List(iList) Then bByte = 1
      Next
      If bByte = 0 Then
         MsgBox "RFQ " & cmbRfq & " Has Not Been Recorded.", _
            vbInformation, Caption
         cmbRfq = "NONE"
      End If
   End If
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDRFQ = "" & cmbRfq
      RdoBid.Update
   End If
   
End Sub


Private Sub cmdCan_Click()
   Dim bByte As Byte
   bByte = CheckBidEntries(bGoodPart, bGoodCust)
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = 1
   
End Sub


Private Sub cmdCom_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If cmdCom.Caption = "C&omplete" Then
      sMsg = "Do You Wish To Mark This Bid As Completed?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         On Error Resume Next
         Err = 0
         clsADOCon.ADOErrNum = 0
         sSql = "UPDATE EstiTable SET BIDCOMPLETE=1, " _
                & "BIDCOMPLETED='" & Format(ES_SYSDATE, "mm/dd/yyyy") & "' " _
                & "WHERE BIDREF=" & Val(cmbBid) & " "
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         If clsADOCon.ADOErrNum = 0 Then
            MsgBox "Bid Was Successfully Marked Complete.", _
               vbInformation, Caption
            bGoodBid = GetThisQBid()
         Else
            MsgBox "Bid Not Was Successfully Marked Complete.", _
               vbExclamation, Caption
         End If
      Else
         CancelTrans
      End If
   Else
      sMsg = "Do You Wish To Mark This Bid As Incomplete?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         Err = 0
         sSql = "UPDATE EstiTable SET BIDCOMPLETE=0, " _
                & "BIDCOMPLETED=Null " _
                & "WHERE BIDREF=" & Val(cmbBid) & " "
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         If clsADOCon.ADOErrNum = 0 Then
            MsgBox "Bid Was Successfully Marked Incomplete.", _
               vbInformation, Caption
            bGoodBid = GetThisQBid()
         Else
            MsgBox "Bid Not Was Successfully Marked Incomplete.", _
               vbExclamation, Caption
         End If
      Else
         CancelTrans
      End If
   End If
   
End Sub

Private Sub cmdComments_Click()
   If cmdComments Then
      'See List For Index
      txtCmt.SetFocus
      SysComments.lblListIndex = 7
      SysComments.Show
      cmdComments = False
   End If
   
End Sub

Private Sub cmdDis_Click()
   optFrom.Value = vbChecked
   EstiESe01c.Show
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   cmbBid.Enabled = False
   optVew.Value = vbChecked
   ViewParts.Show
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3501
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub



Private Sub cmdLst_Click()
   On Error Resume Next
   Frame1(1).Visible = False
   Frame1(0).Visible = True
   If optCom.Value = vbUnchecked Then cmdCom.Enabled = True
   cmbCls.Enabled = True
   cmbBid.Enabled = True
   txtDte.Enabled = True
   txtQty.SetFocus
   
End Sub

Private Sub cmdLst_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub cmdNew_Click()
   EstiESe01h.Show
   
End Sub

Private Sub cmdNxt_Click()
   Dim bByte As Byte
   bByte = CheckBidEntries(bGoodPart, bGoodCust)
   If bByte = 1 Then Exit Sub
   
   On Error Resume Next
   cmdCom.Enabled = False
   cmbCls.Enabled = False
   cmbBid.Enabled = False
   txtDte.Enabled = False
   Frame1(0).Visible = False
   Frame1(1).Visible = True
   txtSet.SetFocus
   
End Sub

Private Sub cmdNxt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub cmdPrt_Click()
   optPart.Value = vbChecked
   cmbBid.Enabled = False
   EstiESe01p.optQwik.Value = vbChecked
   EstiESe01p.txtPrt = cmbPrt
   EstiESe01p.Show
   
End Sub

Private Sub cmdVew_Click()
   optSle.Value = vbChecked
   ViewSales.lblPrt = cmbPrt
   ViewSales.Show
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   On Error Resume Next
   sCurrEstimator = GetSetting("Esi2000", "EsiEngr", "Estimator", sCurrEstimator)
   If bOnLoad Then
      GetRates
      cmdComments.Enabled = True
      GetNextBid Me
      txtPrt.ToolTipText = "No Part Number Has Been Entered."
      FillCustomers
      FillCombo
      
      Dim bPartSearch As Boolean
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillPartCombo cmbPrt
      
      cmbPrt = ""
      If cmbCst.ListCount > 0 Then bGoodCust = GetBidCustomer(Me, cmbCst)
      bOnLoad = 0
   End If
   If optVew.Value = vbChecked Then
      optVew.Value = vbUnchecked
      Unload ViewParts
   End If
   If optSle.Value = vbChecked Then
      optSle.Value = vbUnchecked
      Unload ViewSales
   End If
   If optFrom.Value = vbChecked Then
      Unload EstiESe01c
      optFrom.Value = vbUnchecked
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Width = 6340
   Frame1(0).Left = 10
   Frame1(1).Left = 10
   Frame1(1).Visible = False
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bResponse As Byte
   If Trim(cmbCst) = "" Or Trim(cmbPrt) = "" Then
      bResponse = MsgBox("The Current Estimate Is Missing a Part Number And/Or " & vbCrLf _
                  & "Customer And Will Be Removed. Do You Still Wish To Quit?", _
                  ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         Cancel = True
      Else
         sSql = "DELETE FROM EstiTable WHERE BIDREF=" & Val(cmbBid) & " "
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
      End If
   Else
      If optPart.Value = vbChecked Then Unload EstiESe01p
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   RdoBid.Close
   Set EstiESe01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   Dim bScrap As Byte
   Dim bGna As Byte
   Dim bProfit As Byte
   
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   For b = 65 To 88
      cmbCls.AddItem Chr$(b)
   Next
   cmbCls = "Q"
   cmbCls.AddItem Chr$(b)
   If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
 '  txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   txtDte = ""
   txtDte.ToolTipText = "The Date Of The Current Bid. Default Today For New Bids."
'   txtFst = Format(Now + 60, "mm/dd/yy")
   txtFst = ""
   GetEstimatingPermissions bScrap, bGna, bProfit
   If bScrap = 0 Then
      txtScr.BackColor = Es_TextDisabled
      txtScr.Locked = True
   End If
   If bGna = 0 Then
      txtGna.BackColor = Es_TextDisabled
      txtGna.Locked = True
   End If
   
   If bProfit = 0 Then
      txtPrf.BackColor = Es_TextDisabled
      txtPrf.Locked = True
   End If
   txtPrc.BackColor = Es_TextDisabled
   OpenBoxes (0)
   
End Sub

Private Sub FillCombo()
   cmbBid.Clear
   On Error GoTo DiaErr1
   FillEstimateCombo Me, "QWIK"
   If cmbBid.ListCount > 0 Then
      cmbBid = cmbBid.List(0)
      bGoodBid = GetThisQBid(True)
   Else
      cmbBid = lblNxt
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub lblBid_Change(Index As Integer)
   '
End Sub

Private Sub lblBrd_Change()
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDBURDEN = Val(lblBrd)
      RdoBid!BIDBURDENRATE = cBMbrdRate
      RdoBid.Update
   End If
   
End Sub

Private Sub lblFoh_Change()
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDFOH = Val(lblFoh)
      RdoBid!BIDFOHRATE = cBFohRate
      RdoBid.Update
   End If
   
End Sub

Private Sub lblLbr_Change()
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDTOTLABOR = Val(lblLbr)
      RdoBid.Update
   End If
   
End Sub

Private Sub lblMat_Change()
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDTOTMATL = Val(lblMat)
      RdoBid.Update
   End If
   
End Sub

Private Sub optPart_Click()
   'never visible - EstiESe01p is loaded
   
End Sub

Private Sub optSle_Click()
   On Error Resume Next
   If optSle.Value = vbUnchecked Then cmbCst.SetFocus
   
End Sub

Private Sub optVew_Click()
   On Error Resume Next
   If optVew.Value = vbUnchecked Then cmbPrt.SetFocus
   
End Sub


Private Sub Timer1_Timer()
   GetNextBid Me
   
End Sub

Private Sub txtByr_LostFocus()
   txtByr = CheckLen(txtByr, 30)
   txtByr = StrCase(txtByr)
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDBUYER = Trim(txtByr)
      RdoBid.Update
   End If
   
End Sub


Private Sub txtCm1_LostFocus()
   txtCm1 = CheckLen(txtCm1, 40)
   txtCm1 = StrCase(txtCm1)
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDMATLDESC = Trim(txtCm1)
      RdoBid.Update
   End If
   
End Sub


Private Sub txtCm2_LostFocus()
   txtCm2 = CheckLen(txtCm2, 40)
   txtCm2 = StrCase(txtCm2)
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDOSPDESC = Trim(txtCm2)
      RdoBid.Update
   End If
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 1024)
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDCOMMENT = Trim(txtCmt)
      RdoBid.Update
   End If
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDateEx(txtDte)
   cmbBid.Enabled = True
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDDATE = Format(txtDte, "mm/dd/yyyy")
      RdoBid.Update
   End If
   bCanceled = 0
   
End Sub


Private Sub txtDue_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtDue_LostFocus()
   txtDue = CheckDateEx(txtDue)
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDDUE = Format(txtDue, "mm/dd/yyyy")
      RdoBid.Update
   End If
   
End Sub


Private Sub txtEpd_LostFocus()
   txtEpd = CheckLen(txtEpd, 9)
   txtEpd = Format(Abs(Val(txtEpd)), "####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDEXPEDITE = Val(txtEpd)
      RdoBid.Update
   End If
   UnitPrice
   
End Sub


Private Sub txtEst_LostFocus()
   txtEst = CheckLen(txtEst, 30)
   txtEst = StrCase(txtEst)
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDESTIMATOR = Trim(txtEst)
      RdoBid.Update
   End If
   If Trim(txtEst) <> "" Then
      SaveSetting "Esi2000", "EsiEngr", "Estimator", txtEst
      sCurrEstimator = txtEst
      sEstimator = txtEst
   End If
   Estimator = sCurrEstimator
   
End Sub


Private Sub txtFst_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtFst_LostFocus()
   If Trim(txtFst) = "" Then
      If bGoodBid Then
         On Error Resume Next
         'RdoBid.Edit
         RdoBid!BIDFIRSTDELIVERY = Null
         RdoBid.Update
      End If
   Else
      txtFst = CheckDateEx(txtFst)
      If bGoodBid Then
         On Error Resume Next
         'RdoBid.Edit
         RdoBid!BIDFIRSTDELIVERY = Format(txtFst, "mm/dd/yyyy")
         RdoBid.Update
      End If
   End If
   
End Sub


Private Sub txtGna_LostFocus()
   txtGna = CheckLen(txtGna, 7)
   txtGna = Format(Abs(Val(txtGna)), ES_QuantityDataFormat)
   UnitPrice
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDGNARATE = (Val(txtGna) / 100)
      RdoBid!BIDGNA = cBGna
      RdoBid.Update
   End If
   
End Sub


Private Sub txtHrs_LostFocus()
   txtHrs = CheckLen(txtHrs, 7)
   txtHrs = Format(Abs(Val(txtHrs)), ES_QuantityDataFormat)
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDHOURS = Val(txtHrs)
      RdoBid.Update
   End If
   UnitPrice
   
End Sub


Private Sub txtMtl_LostFocus()
   txtMtl = CheckLen(txtMtl, 7)
   txtMtl = Format(Abs(Val(txtMtl)), ES_QuantityDataFormat)
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDMATL = Val(txtMtl)
      RdoBid.Update
   End If
   UnitPrice
   
End Sub


Private Sub txtOsp_LostFocus()
   txtOsp = CheckLen(txtOsp, 7)
   txtOsp = Format(Abs(Val(txtOsp)), ES_QuantityDataFormat)
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDOSP = Val(txtOsp)
      RdoBid.Update
   End If
   UnitPrice
   
End Sub


Private Sub txtPrc_Change()
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDUNITPRICE = Val(txtPrc)
      RdoBid.Update
      UpdateDiscounts
   End If
   
End Sub

Private Sub txtPrc_LostFocus()
   txtPrc = CheckLen(txtPrc, 9)
   txtPrc = Format(Abs(Val(txtPrc)), "####0.00")
   
End Sub


Private Sub txtPrf_LostFocus()
   txtPrf = CheckLen(txtPrf, 7)
   txtPrf = Format(Abs(Val(txtPrf)), "#0.000")
   UnitPrice
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDPROFITRATE = Val(txtPrf) / 100
      RdoBid!BIDPROFIT = cBprofit
      RdoBid.Update
   End If
   
End Sub


Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      optVew.Value = vbChecked
      ViewParts.Show
   ElseIf KeyCode = vbKeyF2 Then
      optPart.Value = vbChecked
      cmbBid.Enabled = False
      EstiESe01p.optQwik.Value = vbChecked
      EstiESe01p.txtPrt = txtPrt
      EstiESe01p.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   bGoodPart = GetBidPart(Me)
   If bGoodPart Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BidPart = Compress(txtPrt)
      If bGoodCust Then
         RdoBid!BIDLOCKED = 0
      Else
         RdoBid!BIDLOCKED = 1
      End If
      RdoBid.Update
   End If
   cmbBid.Enabled = True
   
End Sub



Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   bGoodPart = GetBidPart(Me)
   If bGoodPart Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BidPart = Compress(cmbPrt)
      If bGoodCust Then
         RdoBid!BIDLOCKED = 0
      Else
         RdoBid!BIDLOCKED = 1
      End If
      RdoBid.Update
   End If
   cmbBid.Enabled = True
   
End Sub

Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 10)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BidQuantity = Val(txtQty)
      RdoBid.Update
   End If
   UnitPrice
   
End Sub


Private Sub txtRte_LostFocus()
   txtRte = CheckLen(txtRte, 7)
   txtRte = Format(Abs(Val(txtRte)), ES_QuantityDataFormat)
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDRATE = Val(txtRte)
      RdoBid.Update
   End If
   UnitPrice
   
End Sub



Private Sub GetRates()
   On Error GoTo DiaErr1
   GetEstimatingRates cMBurden, cFoh, cRate, cGna, cProfit, cScrap
   
   'Defaults
   cBFohRate = cFoh
   cBMbrdRate = cMBurden
   cBGnaRate = cGna
   cBProfitRate = cProfit
   cBScrapRate = cScrap
   sEstimator = GetSetting("Esi2000", "EsiEngr", "Estimator", sEstimator)
   Exit Sub
   
DiaErr1:
   sProcName = "getrates"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtScr_LostFocus()
   txtScr = CheckLen(txtScr, 7)
   txtScr = Format(Abs(Val(txtScr)), ES_QuantityDataFormat)
   UnitPrice
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDSCRAPRATE = (Val(txtScr) / 100)
      RdoBid!BIDSCRAP = cBScrap
      RdoBid.Update
   End If
   
End Sub


Private Sub txtSet_LostFocus()
   txtSet = CheckLen(txtSet, 9)
   txtSet = Format(Abs(Val(txtSet)), "####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDSETUP = Val(txtSet)
      RdoBid.Update
   End If
   UnitPrice
   
End Sub


Private Sub txtTlg_LostFocus()
   txtTlg = CheckLen(txtTlg, 7)
   txtTlg = Format(Abs(Val(txtTlg)), "####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoBid.Edit
      RdoBid!BIDTOOLING = Val(txtTlg)
      RdoBid.Update
   End If
   UnitPrice
   
End Sub





'Calculate the bottom line

Private Sub UnitPrice()
   Dim cULabor As Currency
   Dim cUMatl As Currency
   Dim cUMtlb As Currency
   Dim cUPrice As Currency
   Dim cUFoh As Currency
   Dim cUOsp As Currency
   Dim cTotalBid As Currency
   
   On Error GoTo DiaErr1
   'Labor
   lblLbr = "0"
   If Val(txtPrf) > 0 Then cBProfitRate = Val(txtPrf) / 100
   cBGnaRate = Val(txtGna) / 100
   cRate = Val(txtRte)
   cBFoh = (Val(txtHrs * cRate) * cBFohRate)
   cULabor = (Val(txtHrs * cRate))
   lblFoh = Format(cBFoh, ES_QuantityDataFormat)
   lblLbr = Format(cBFoh + cULabor, ES_QuantityDataFormat)
   cUPrice = cULabor + cBFoh
   
   'Material
   lblMat = "0"
   cUMatl = Val(txtMtl)
   cUMtlb = (cUMatl * cBMbrdRate)
   lblBrd = Format(cUMtlb, ES_QuantityDataFormat)
   cUMatl = cUMtlb + cUMatl
   lblMat = Format(cUMatl, ES_QuantityDataFormat)
   cUPrice = cUPrice + cUMatl
   
   'Scrap
   cBScrapRate = (Val(txtScr) / 100)
   cBScrap = cBScrapRate * (Val(lblMat) + Val(lblLbr))
   cUPrice = cUPrice + cBScrap
   
   'OSP
   cUOsp = Val(txtOsp)
   cUPrice = cUPrice + cUOsp
   
   'Gna
   cBGna = (Val(lblLbr) + Val(lblMat) + Val(txtOsp) + cBScrap) * cBGnaRate
   
   'force change
   txtPrc = "0"
   
   'Price
   cUPrice = cUPrice + (cUPrice * cBGnaRate)
   cBprofit = (cUPrice * cBProfitRate)
   cUPrice = cUPrice + (cBprofit)
   cUPrice = Int(cUPrice * 100) / 100
   txtPrc = Format(cUPrice, "####0.000")
   txtPrf = Format(cBProfitRate * 100, "####0.000")
   
   txtGna = Format(cBGnaRate * 100, "#####0.000")
   cTotalBid = (cUPrice * Val(txtQty)) + Val(txtSet) + Val(txtEpd) + Val(txtTlg)
   lblBid(0) = Format(cTotalBid, "######0.000")
   lblBid(1) = Format(cTotalBid, "######0.000")
   
   'Ticket #59986 Update the values that are not being done
   'RdoBid.Edit
   RdoBid!BIDGNA = cBGna
   RdoBid!BIDPROFIT = cBprofit
   RdoBid.Update
   
   
   Exit Sub
   
DiaErr1:
   Err = 0
   clsADOCon.ADOErrNum = 0
   Resume Next
   
End Sub


Private Sub OpenBoxes(bOpen As Byte)
   On Error Resume Next
   Frame1(0).Visible = True
   Frame1(1).Visible = False
   cmbRfq = "NONE"
   txtQty = "1.000"
   txtHrs = "0.000"
   txtRte = "0.000"
   txtMtl = "0.000"
   txtOsp = "0.000"
   txtPrc = "0.00"
   txtEpd = "0.00"
   txtSet = "0.00"
   txtTlg = "0.00"
   lblFoh = "0.000"
   lblLbr = "0.000"
   lblBrd = "0.000"
   lblBid(0) = "0.000"
   lblBid(1) = "0.000"
   If bOpen = 0 Then
      cmdCom.Enabled = False
      cmbCst.ForeColor = vbGrayText
      cmbRfq.ForeColor = vbGrayText
      txtPrt.ForeColor = vbGrayText
      txtQty.ForeColor = vbGrayText
      txtHrs.ForeColor = vbGrayText
      txtRte.ForeColor = vbGrayText
      txtMtl.ForeColor = vbGrayText
      txtQty.ForeColor = vbGrayText
      txtCm1.ForeColor = vbGrayText
      txtOsp.ForeColor = vbGrayText
      txtCm2.ForeColor = vbGrayText
      txtSet.ForeColor = vbGrayText
      txtEpd.ForeColor = vbGrayText
      txtTlg.ForeColor = vbGrayText
      txtCmt.ForeColor = vbGrayText
      txtPrf.ForeColor = vbGrayText
      txtGna.ForeColor = vbGrayText
      txtPrc.ForeColor = vbGrayText
      txtScr.ForeColor = vbGrayText
      Frame1(0).Enabled = False
   Else
      cmbCst.ForeColor = Es_TextForeColor
      txtPrt.ForeColor = Es_TextForeColor
      txtQty.ForeColor = Es_TextForeColor
      txtHrs.ForeColor = Es_TextForeColor
      txtRte.ForeColor = Es_TextForeColor
      txtMtl.ForeColor = Es_TextForeColor
      txtQty.ForeColor = Es_TextForeColor
      txtCm1.ForeColor = Es_TextForeColor
      txtOsp.ForeColor = Es_TextForeColor
      txtCm2.ForeColor = Es_TextForeColor
      txtSet.ForeColor = Es_TextForeColor
      txtEpd.ForeColor = Es_TextForeColor
      txtTlg.ForeColor = Es_TextForeColor
      txtCmt.ForeColor = Es_TextForeColor
      txtScr.ForeColor = Es_TextForeColor
      cmbRfq.ForeColor = ES_BLUE
      txtPrc.ForeColor = ES_BLUE
      txtPrf.ForeColor = ES_BLUE
      txtGna.ForeColor = ES_BLUE
      Frame1(0).Enabled = True
   End If
   Refresh
   
End Sub

Public Function GetThisQBid(Optional bHideMsg As Boolean) As Byte
   Dim lBidNo As Long
   lBidNo = Val(cmbBid)
   bGoodBid = 0
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM EstiTable WHERE BIDREF=" & lBidNo & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBid, ES_KEYSET)
   If bSqlRows Then
      With RdoBid
         If !BIDCANCELED = 1 Then
            cmdCom.Enabled = False
            ClearResultSet RdoBid
            OpenBoxes 0
            bCanceled = 1
            Exit Function
         End If
         If !BIDACCEPTED = 1 Then
            cmdCom.Enabled = False
            If Not bHideMsg Then
               MsgBox "Bid " & cmbBid & " Was Accepted And Cannot Be Edited.", _
                  vbInformation
            End If
            ClearResultSet RdoBid
            OpenBoxes 0
            bCanceled = 1
            Exit Function
         End If
         bCanceled = 0
         OpenBoxes 1
         If !BIDCOMPLETE = 1 Then
            cmdCom.Caption = "Unc&omplete"
            cmdCom.ToolTipText = "Remove Complete Flag From Bid"
            cmbPrt.Enabled = False
            cmbCst.Enabled = False
            optCom.Value = vbChecked
         Else
            cmdCom.Caption = "C&omplete"
            cmdCom.ToolTipText = "Mark Bid As Completed"
            cmdCom.Enabled = True
            cmbPrt.Enabled = True
            cmbCst.Enabled = True
            optCom.Value = vbUnchecked
         End If
         If Trim(!BidClass) = "FULL" Then
            MsgBox "Estimate Is A Full Bid.  Edit From The Full Bid Area.", _
               vbInformation, Caption
            OpenBoxes 0
            bCanceled = 1
            Exit Function
         End If
         cmbCls = "" & Trim(!BIDPRE)
         cmbPrt = "" & Trim(!BidPart)
         bGoodPart = GetBidPart(Me)
         If "" & Trim(!BIDCUST) <> "" Then cmbCst = "" & Trim(!BIDCUST)
         bGoodCust = GetBidCustomer(Me, cmbCst)
         txtDte = Format(!BIDDATE, "mm/dd/yyyy")
         txtQty = Format(!BidQuantity, ES_QuantityDataFormat)
         txtHrs = Format(!BIDHOURS, ES_QuantityDataFormat)
         If !BIDRATE = 0 Then
            txtRte = Format(cRate, ES_QuantityDataFormat)
         Else
            txtRte = Format(!BIDRATE, ES_QuantityDataFormat)
         End If
         If !BIDFOHRATE = 0 Then cBFohRate = cFoh _
                          Else cBFohRate = !BIDFOHRATE
         
         If !BIDBURDENRATE = 0 Then cBMbrdRate = cMBurden _
                             Else cBMbrdRate = !BIDBURDENRATE
         
         If !BIDGNARATE = 0 Then cBGnaRate = cGna _
                          Else cBGnaRate = !BIDGNARATE
         
         If !BIDPROFITRATE = 0 Then cBProfitRate = cProfit _
                             Else cBProfitRate = !BIDPROFITRATE
         
         If !BIDPROFIT = 0 Then cBprofit = cProfit _
                         Else cBprofit = !BIDPROFIT
         
         If !BIDSCRAPRATE = 0 Then cBScrapRate = cScrap _
                            Else cBScrapRate = !BIDSCRAPRATE
         
         cBScrap = Format(!BIDSCRAP, ES_QuantityDataFormat)
         txtScr = Format(cBScrapRate * 100, ES_QuantityDataFormat)
         lblFoh = Format(!BIDFOH, ES_QuantityDataFormat)
         lblLbr = Format(!BIDTOTLABOR, ES_QuantityDataFormat)
         txtMtl = Format(!BIDMATL, ES_QuantityDataFormat)
         txtCm1 = "" & Trim(!BIDMATLDESC)
         lblMat = Format(!BIDTOTMATL, ES_QuantityDataFormat)
         txtOsp = Format(!BIDOSP, ES_QuantityDataFormat)
         txtCm2 = "" & Trim(!BIDOSPDESC)
         txtPrf = Format(cBProfitRate * 100, ES_QuantityDataFormat)
         txtGna = Format(cBGnaRate * 100, ES_QuantityDataFormat)
         txtPrc = Format(cBProfitRate, "#####0.00")
         txtSet = Format(!BIDSETUP, "#####0.00")
         txtEpd = Format(!BIDEXPEDITE, "#####0.00")
         txtTlg = Format(!BIDTOOLING, "#####0.00")
         txtFst = Format(!BIDFIRSTDELIVERY, "mm/dd/yyyy")
         If Not IsNull(!BIDDUE) Then
            txtDue = Format(!BIDDUE, "mm/dd/yyyy")
         Else
            txtDue = Format(!BIDDATE + 10, "mm/dd/yyyy")
         End If
         txtCmt = "" & Trim(!BIDCOMMENT)
         If "" & Trim(!BIDESTIMATOR) = "" Then
            txtEst = sEstimator
         Else
            txtEst = "" & Trim(!BIDESTIMATOR)
         End If
         Estimator = txtEst
         FindCustomer Me, cmbCst
         FillCustomerRFQs Me, cmbCst, True
         cmbRfq = "" & Trim(!BIDRFQ)
         txtByr = "" & Trim(!BIDBUYER)
         cmdCom.Enabled = True
         UnitPrice
      End With
      GetThisQBid = 1
   Else
      'Reset defaults
      cmdCom.Enabled = False
      cBFohRate = cFoh
      cBMbrdRate = cMBurden
      cBGnaRate = cGna
      cBProfitRate = cProfit
      cBScrapRate = cScrap
      cBprofit = 0
      GetThisQBid = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "GetThisQBid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Function

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPrt.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPrt.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function

