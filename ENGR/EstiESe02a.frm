VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Full Estimate"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   3502
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtQuoteDt 
      Height          =   315
      Left            =   8640
      TabIndex        =   102
      Tag             =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtLeadTime 
      Height          =   285
      Left            =   8640
      TabIndex        =   101
      Tag             =   "2"
      ToolTipText     =   "Profit Calculated On The Subtotal Of All (Including G&A)"
      Top             =   2280
      Width           =   2595
   End
   Begin VB.TextBox txtFOB 
      Height          =   285
      Left            =   8640
      TabIndex        =   100
      Tag             =   "2"
      ToolTipText     =   "Profit Calculated On The Subtotal Of All (Including G&A)"
      Top             =   2640
      Width           =   2595
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   1560
      Width           =   3075
   End
   Begin VB.TextBox txtEstimator 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1680
      TabIndex        =   99
      ToolTipText     =   "Double Click to Edit"
      Top             =   6960
      Width           =   2055
   End
   Begin VB.TextBox txtTlg 
      Height          =   285
      Left            =   6420
      TabIndex        =   18
      Tag             =   "1"
      Top             =   6240
      Width           =   1050
   End
   Begin VB.TextBox txtEpd 
      Height          =   285
      Left            =   5280
      TabIndex        =   17
      Tag             =   "1"
      Top             =   6240
      Width           =   1050
   End
   Begin VB.TextBox txtSet 
      Height          =   285
      Left            =   4140
      TabIndex        =   16
      Tag             =   "1"
      Top             =   6240
      Width           =   1050
   End
   Begin VB.TextBox txtByr 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Tag             =   "2"
      ToolTipText     =   "Profit Calculated On The Subtotal Of All (Including G&A)"
      Top             =   2640
      Width           =   2595
   End
   Begin VB.ComboBox txtFst 
      Height          =   315
      Left            =   8640
      TabIndex        =   11
      Tag             =   "4"
      Top             =   1920
      Width           =   1250
   End
   Begin VB.TextBox txtCmt 
      Height          =   1425
      Left            =   6300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "9"
      Top             =   60
      Width           =   3795
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "EstiESe02a.frx":0000
      DownPicture     =   "EstiESe02a.frx":0972
      Height          =   350
      Left            =   5820
      Picture         =   "EstiESe02a.frx":12E4
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "Standard Comments"
      Top             =   60
      Width           =   350
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   4740
      Top             =   60
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESe02a.frx":18E6
      Style           =   1  'Graphical
      TabIndex        =   73
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtDue 
      Height          =   315
      Left            =   4440
      TabIndex        =   2
      Tag             =   "4"
      ToolTipText     =   "Due Date"
      Top             =   720
      Width           =   1250
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "N&ew"
      Height          =   285
      Left            =   10260
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Add A New Estimate"
      Top             =   540
      Width           =   1000
   End
   Begin VB.CheckBox optFrom 
      Caption         =   "from"
      Height          =   255
      Left            =   1560
      TabIndex        =   69
      Top             =   60
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdDis 
      Caption         =   "&Discounts"
      Enabled         =   0   'False
      Height          =   285
      Left            =   10260
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Volume Discounts"
      Top             =   1200
      Width           =   1000
   End
   Begin VB.CommandButton cmdCom 
      Caption         =   "Unc&omplete"
      Enabled         =   0   'False
      Height          =   285
      Left            =   10260
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Mark This Bid As Complete"
      Top             =   855
      Width           =   1000
   End
   Begin VB.CheckBox optSle 
      Caption         =   "View Sales"
      Height          =   255
      Left            =   60
      TabIndex        =   54
      Top             =   7560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View Parts"
      Height          =   255
      Left            =   60
      TabIndex        =   53
      Top             =   7680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox optPart 
      Caption         =   "Parts"
      Height          =   195
      Left            =   1440
      TabIndex        =   52
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Unit Costs"
      Top             =   1080
      Width           =   915
   End
   Begin VB.TextBox txtPrf 
      Height          =   285
      Left            =   2880
      TabIndex        =   15
      Tag             =   "1"
      ToolTipText     =   "Locked TextBoxes"
      Top             =   5520
      Width           =   750
   End
   Begin VB.TextBox txtGna 
      Height          =   285
      Left            =   2880
      TabIndex        =   14
      Tag             =   "1"
      ToolTipText     =   "Locked TextBoxes"
      Top             =   5160
      Width           =   750
   End
   Begin VB.TextBox txtScr 
      Height          =   285
      Left            =   2880
      TabIndex        =   13
      Tag             =   "1"
      ToolTipText     =   "Locked TextBoxes"
      Top             =   4800
      Width           =   750
   End
   Begin VB.CommandButton cmdSrv 
      Caption         =   "S&ervices"
      Height          =   300
      Left            =   2880
      MaskColor       =   &H8000000F&
      TabIndex        =   28
      ToolTipText     =   "Services And Processing"
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdMat 
      Caption         =   "&Material"
      Height          =   300
      Left            =   2880
      MaskColor       =   &H8000000F&
      TabIndex        =   27
      ToolTipText     =   "Material-Parts List"
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdLbr 
      Caption         =   "&Labor"
      Height          =   300
      Left            =   2880
      MaskColor       =   &H8000000F&
      TabIndex        =   26
      ToolTipText     =   "Labor-Routing"
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame z2 
      Height          =   40
      Left            =   120
      TabIndex        =   37
      Top             =   1440
      Width           =   5955
   End
   Begin VB.CheckBox optCom 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4200
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "EstiESe02a.frx":2094
      Height          =   315
      Left            =   4800
      Picture         =   "EstiESe02a.frx":23D6
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number Find A Part Number (F4 At Part Number)"
      Top             =   1560
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Tag             =   "3"
      ToolTipText     =   "Enter A Part Or Click The ""Look Up"""
      Top             =   1560
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.ComboBox cmbCst 
      Height          =   288
      Left            =   1680
      TabIndex        =   9
      Tag             =   "3"
      ToolTipText     =   "Select A Customer"
      Top             =   1920
      Width           =   1555
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "EstiESe02a.frx":2718
      Height          =   315
      Left            =   5280
      Picture         =   "EstiESe02a.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Show Sales Orders For Parts"
      Top             =   1560
      Width           =   350
   End
   Begin VB.ComboBox cmbRfq 
      Height          =   315
      Left            =   4080
      TabIndex        =   10
      Tag             =   "3"
      ToolTipText     =   "Select RFQ From List"
      Top             =   1920
      Width           =   2040
   End
   Begin VB.CommandButton cmdPrt 
      Height          =   315
      Left            =   5760
      Picture         =   "EstiESe02a.frx":30CC
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "New Part Numbers Find A Part Number (F2 At Part Number)"
      Top             =   1560
      Width           =   350
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Class A-Z"
      Top             =   1080
      Width           =   495
   End
   Begin VB.ComboBox cmbBid 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "Select Or Enter A Bid Number (Contains Full Bids)"
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Tag             =   "4"
      Top             =   720
      Width           =   1250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   10260
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   60
      Width           =   1000
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6660
      Top             =   7440
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7395
      FormDesignWidth =   11415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quote date"
      Height          =   255
      Index           =   12
      Left            =   7080
      TabIndex        =   105
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lead Time"
      Height          =   255
      Index           =   13
      Left            =   7080
      TabIndex        =   104
      Top             =   2340
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "FOB"
      Height          =   255
      Index           =   14
      Left            =   7080
      TabIndex        =   103
      Top             =   2700
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   98
      Top             =   6600
      Width           =   2475
   End
   Begin VB.Label lblEstTotLessSET 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8760
      TabIndex        =   97
      ToolTipText     =   "Total Of This Estimate"
      Top             =   5880
      Width           =   1050
   End
   Begin VB.Label lblPreScrapUnit 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4140
      TabIndex        =   96
      ToolTipText     =   "Unit Scrap"
      Top             =   4800
      Width           =   1050
   End
   Begin VB.Label lblPreGNAUnit 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4140
      TabIndex        =   95
      ToolTipText     =   "Unit General And Administration"
      Top             =   5160
      Width           =   1050
   End
   Begin VB.Label lblPreProfitUnit 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4140
      TabIndex        =   94
      ToolTipText     =   "Unit Profit"
      Top             =   5520
      Width           =   1050
   End
   Begin VB.Label lblUnitTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7560
      TabIndex        =   93
      ToolTipText     =   "Total Of This Estimate"
      Top             =   5880
      Width           =   1050
   End
   Begin VB.Label lblUnitLaborCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5280
      TabIndex        =   92
      ToolTipText     =   "Unit Labor Hours"
      Top             =   3720
      Width           =   1050
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Overhead Burden"
      Height          =   420
      Left            =   6420
      TabIndex        =   91
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Unit Hours"
      Height          =   420
      Left            =   4200
      TabIndex        =   90
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Unit Cost"
      Height          =   420
      Left            =   5460
      TabIndex        =   89
      Top             =   3240
      Width           =   555
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Estimate  Totals"
      Height          =   420
      Left            =   8820
      TabIndex        =   88
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblEstTotalLabor 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8760
      TabIndex        =   87
      ToolTipText     =   "Total Unit Labor"
      Top             =   3720
      Width           =   1050
   End
   Begin VB.Label lblEstTotMatl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8760
      TabIndex        =   86
      ToolTipText     =   "Total Unit Material"
      Top             =   4080
      Width           =   1050
   End
   Begin VB.Label lblEstTotScrap 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8760
      TabIndex        =   85
      ToolTipText     =   "Unit Scrap"
      Top             =   4800
      Width           =   1050
   End
   Begin VB.Label lblEstTotGA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8760
      TabIndex        =   84
      ToolTipText     =   "Unit General And Administration"
      Top             =   5160
      Width           =   1050
   End
   Begin VB.Label lblEstTotProfit 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8760
      TabIndex        =   83
      ToolTipText     =   "Unit Profit"
      Top             =   5520
      Width           =   1050
   End
   Begin VB.Label lblEstTotSET 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8760
      TabIndex        =   82
      ToolTipText     =   "Total Other Costs"
      Top             =   6240
      Width           =   1050
   End
   Begin VB.Label lblEstTotServiceOld 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7800
      TabIndex        =   81
      ToolTipText     =   "Total Unit Services"
      Top             =   7680
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Unit Totals"
      Height          =   420
      Left            =   7740
      TabIndex        =   80
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buyer"
      Height          =   255
      Index           =   33
      Left            =   240
      TabIndex        =   79
      Top             =   2640
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "First Delivery Date"
      Height          =   255
      Index           =   23
      Left            =   7080
      TabIndex        =   78
      Top             =   1980
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      Height          =   255
      Index           =   24
      Left            =   7560
      TabIndex        =   77
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label txtRte 
      Caption         =   "Rate"
      Height          =   255
      Left            =   2220
      TabIndex        =   75
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimator"
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   74
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label lblRouting 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   300
      TabIndex        =   72
      Top             =   8280
      Width           =   3495
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1680
      TabIndex        =   71
      Top             =   2280
      Width           =   3852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bid Due"
      Height          =   255
      Index           =   34
      Left            =   3360
      TabIndex        =   70
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblEstimator 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3780
      TabIndex        =   68
      ToolTipText     =   "Labor Hours"
      Top             =   8280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblUnitServices 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7560
      TabIndex        =   67
      ToolTipText     =   "Total Unit Services"
      Top             =   4440
      Width           =   1050
   End
   Begin VB.Label lblRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5220
      TabIndex        =   66
      ToolTipText     =   "Labor Hours"
      Top             =   7560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblFohRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4260
      TabIndex        =   65
      ToolTipText     =   "Labor Hours"
      Top             =   7560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup / Expediting / Tooling"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   64
      Top             =   6240
      Width           =   2475
   End
   Begin VB.Label PALEVEL 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2940
      TabIndex        =   63
      ToolTipText     =   "Material"
      Top             =   7680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEstGrandTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8760
      TabIndex        =   62
      ToolTipText     =   "Total Of This Estimate"
      Top             =   6600
      Width           =   1050
   End
   Begin VB.Label lblPrf 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7560
      TabIndex        =   61
      ToolTipText     =   "Unit Profit"
      Top             =   5520
      Width           =   1050
   End
   Begin VB.Label lblGNA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7560
      TabIndex        =   60
      ToolTipText     =   "Unit General And Administration"
      Top             =   5160
      Width           =   1050
   End
   Begin VB.Label lblScrap 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7560
      TabIndex        =   59
      ToolTipText     =   "Unit Scrap"
      Top             =   4800
      Width           =   1050
   End
   Begin VB.Label lblUnitOverheadCost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6420
      TabIndex        =   58
      ToolTipText     =   "Labor Overhead"
      Top             =   3720
      Width           =   1050
   End
   Begin VB.Label lblTotMat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7560
      TabIndex        =   57
      ToolTipText     =   "Total Unit Material"
      Top             =   4080
      Width           =   1050
   End
   Begin VB.Label lblBurden 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6420
      TabIndex        =   56
      ToolTipText     =   "Material Burden"
      Top             =   4080
      Width           =   1050
   End
   Begin VB.Label LblUnitLabor 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7560
      TabIndex        =   55
      ToolTipText     =   "Total Unit Labor"
      Top             =   3720
      Width           =   1050
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   51
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Totals"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   50
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Profit"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   49
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "% of"
      Height          =   255
      Index           =   15
      Left            =   3660
      TabIndex        =   48
      Top             =   5580
      Width           =   450
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "% of"
      Height          =   255
      Index           =   27
      Left            =   3660
      TabIndex        =   47
      Top             =   5220
      Width           =   450
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "G&A (General And Admin)"
      Height          =   255
      Index           =   28
      Left            =   240
      TabIndex        =   46
      Top             =   5160
      UseMnemonic     =   0   'False
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reduction (Scrap)"
      Height          =   255
      Index           =   17
      Left            =   240
      TabIndex        =   45
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "% of"
      Height          =   255
      Index           =   18
      Left            =   3660
      TabIndex        =   44
      Top             =   4860
      Width           =   450
   End
   Begin VB.Label lblTotServices 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8760
      TabIndex        =   43
      ToolTipText     =   "Total Services"
      Top             =   4440
      Width           =   1050
   End
   Begin VB.Label lblMaterial 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5280
      TabIndex        =   42
      ToolTipText     =   "Material"
      Top             =   4080
      Width           =   1050
   End
   Begin VB.Label lblHours 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4140
      TabIndex        =   41
      ToolTipText     =   "Unit Labor Hours"
      Top             =   3720
      Width           =   1050
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Services"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   40
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Material (Parts List)"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   39
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Labor (Routing)"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   38
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "RFQ"
      Height          =   255
      Index           =   25
      Left            =   3360
      TabIndex        =   36
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   35
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   34
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate Number"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   33
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bid Date"
      Height          =   255
      Index           =   26
      Left            =   240
      TabIndex        =   32
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Next Estimate"
      Height          =   255
      Index           =   31
      Left            =   240
      TabIndex        =   31
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblNxt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   30
      Top             =   360
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Complete"
      Height          =   255
      Index           =   35
      Left            =   3360
      TabIndex        =   29
      ToolTipText     =   "Bid Is Marked Complete"
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "EstiESe02a"
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

'Dim bBomLevel As Byte
Dim bCanceled As Byte
Dim bOnLoad As Byte
Dim bOpenKey As Byte
Dim bGoodCust As Byte
Dim bGoodList As Byte
Dim bGoodPart As Byte
Dim bGoodRout As Byte

Dim iBids As Integer
Dim bNoEdit As Boolean


'General defaults
Dim cFoh As Currency
Dim cMBurden As Currency
Dim cGna As Currency
Dim cProfit As Currency
Dim cRate As Currency
Dim cOldQty As Currency
Dim cScrap As Currency

'Bid stuff
Dim cBFoh As Currency
Dim cBGna As Currency
Dim cBprofit As Currency
Dim cBScrap As Currency

Dim cBGnaRate As Currency
Dim cBProfitRate As Currency
Dim cBScrapRate As Currency

Dim sBomRev As String
Dim sEstimator As String
Dim sBEstimator As String
Dim sRouting As String

Dim lBids(300) As Long

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
'10/9/06

Private Sub GetRates()
   On Error GoTo DiaErr1
   GetEstimatingRates cMBurden, cFoh, cRate, cGna, cProfit, cScrap
   
   '    sSql = "Qry_EstParameters"
   '    bSqlRows = clsADOCon.GetDataSet(sSql,RdoPar, ES_FORWARD)
   '        If bSqlRows Then
   '            With RdoPar
   '                txtGna = Format(!EstGenAdmnExp, ES_QuantityDataFormat)
   '                cGna = Format(!EstGenAdmnExp / 100, ES_QuantityDataFormat)
   '                cProfit = Format(!EstProfitOfSale / 100, ES_QuantityDataFormat)
   '                txtPrf = Format(cProfit * 100, ES_QuantityDataFormat)
   '                txtScr = Format(!EstScrapRate, ES_QuantityDataFormat)
   '                cScrap = Format(!EstScrapRate / 100, ES_QuantityDataFormat)
   '
   '                'Defaults
   '            End With
   '        End If
   'Defaults
   cBGnaRate = cGna
   cBProfitRate = cProfit
   cBScrapRate = cScrap
   
   sEstimator = GetSetting("Esi2000", "EsiEngr", "Estimator", sEstimator)
   sCurrEstimator = GetSetting("Esi2000", "EsiEngr", "Estimator", sCurrEstimator)
   Exit Sub
   
DiaErr1:
   sProcName = "getrates"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub cmbBid_Click()
   bGoodBid = GetThisFullBid(True)
   'GetBidLabor Compress(txtPrt), Val(cmbBid), txtQty
End Sub


Private Sub cmbBid_LostFocus()
   cmbBid = CheckLen(cmbBid, 6)
   cmbBid = Format(Abs(Val(cmbBid)), "000000")
   If bCanceled Then Exit Sub
   bGoodBid = GetThisFullBid(False)
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
      'RdoFull.Edit
      RdoFull!BIDPRE = Trim(cmbCls)
      RdoFull.Update
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
      'RdoFull.Edit
      RdoFull!BIDCUST = Compress(cmbCst)
      If bGoodPart Then
         RdoFull!BIDLOCKED = 0
      Else
         RdoFull!BIDLOCKED = 1
      End If
      RdoFull.Update
   End If
   '        If bGoodCust = 0 Or bGoodPart = 0 Then
   '            MsgBox "Bids Without A Valid Customer And Valid " _
   '                & "Part Number Will Not Be Saved.", _
   '                    vbInformation, Caption
   '        End If
   FillCustomerRFQs Me, cmbCst, True
   On Error Resume Next
   If Trim(RdoFull!BIDRFQ) <> "NONE" Then cmbRfq = Trim(RdoFull!BIDRFQ)
   
End Sub


Private Sub cmbPrt_Click()
  ' txtPrt = cmbPrt
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
      'RdoFull.Edit
      RdoFull!BIDRFQ = cmbRfq
      RdoFull.Update
   End If
   
End Sub


Private Sub cmdCan_Click()
   Dim bByte As Byte
   bByte = CheckBidEntries(bGoodPart, bGoodCust)
   Unload Me
   
End Sub

Private Sub cmdCom_Click()
   cmdCom_MouseDown 0, 0, 0, 0
   
End Sub

Private Sub cmdCom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim bResponse As Byte
   Dim sMsg As String
   
   If cmdCom.Caption = "C&omplete" Then
      sMsg = "Do You Wish To Mark This Bid As Completed?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         sSql = "UPDATE EstiTable SET BIDCOMPLETE=1, " _
                & "BIDCOMPLETED='" & Format(ES_SYSDATE, "mm/dd/yyyy") & "' " _
                & "WHERE BIDREF=" & Val(cmbBid) & " "
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         If clsADOCon.ADOErrNum = 0 Then
            MsgBox "Bid Was Successfully Marked Complete.", _
               vbInformation, Caption
            bGoodBid = GetThisFullBid()
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
         sSql = "UPDATE EstiTable SET BIDCOMPLETE=0, " _
                & "BIDCOMPLETED=Null " _
                & "WHERE BIDREF=" & Val(cmbBid) & " "
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         If clsADOCon.ADOErrNum = 0 Then
            MsgBox "Bid Was Successfully Marked Incomplete.", _
               vbInformation, Caption
            bGoodBid = GetThisFullBid()
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
   EstiESe01c.optFrom.Value = vbChecked
   EstiESe01c.Show
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   optVew.Value = vbChecked
   ViewParts.Show
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3502
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


'Should have enough options here

Private Sub cmdLbr_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   bResponse = CheckBidEntries(bGoodPart, bGoodCust)
   If bResponse = 1 Then Exit Sub
   
   bGoodRout = GetBidRouting()
   If bGoodRout = 1 Then
      EstiESe02b.lblBid = cmbBid
      EstiESe02b.Show
   Else
      bGoodRout = GetPartRouting()
      If bGoodRout = 0 Then
         'See if they want to use one
         sMsg = "This Part Does Not Have A Routing Would " & vbCrLf _
                & "You Like To Select An Existing Routing?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            EstiEsf04a.optFrom = vbUnchecked
            EstiEsf04a.Show 1
            If sRouting <> "" Then
               CopyARouting
               Exit Sub
            End If
         End If
         
         sMsg = "This Part Does Not Have A Routing." & vbCrLf _
                & "Would You Like To Create One?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            EstiESe02b.lblBid = cmbBid
            EstiESe02b.Show
         Else
            CancelTrans
         End If
      Else
         sMsg = "This Part Has A Routing." & vbCrLf _
                & "Would You Like To Copy It?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            CopyARouting
         Else
            sMsg = "Make A New Bid Only Routing?"
            bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
            If bResponse = vbYes Then
               EstiESe02b.lblBid = cmbBid
               EstiESe02b.Show
            Else
               CancelTrans
            End If
         End If
      End If
   End If
   
End Sub

Private Sub cmdMat_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   bResponse = CheckBidEntries(bGoodPart, bGoodCust)
   If bResponse = 1 Then Exit Sub
   bGoodList = GetBidPartsList()
   If bGoodList = 1 Then
      EstiESe02c.cmbPls = txtPrt
      EstiESe02c.lblBid = cmbBid
      EstiESe02c.Show
   Else
      bGoodList = GetPartsList()
      If bGoodList = 1 Then
         sMsg = "This Part Number Has A Parts List." & vbCrLf _
                & "Copy The Parts List To This Estimate?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            CopyPartsList
         Else
            CancelTrans
         End If
      Else
         sMsg = "This Part Does Not Have A Parts List." & vbCrLf _
                & "Make A New Bid Only Parts List?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            EstiESe02c.cmbPls = Compress(txtPrt)
            EstiESe02c.lblBid = cmbBid
            EstiESe02c.Show
         Else
            CancelTrans
         End If
      End If
   End If
   
End Sub


Private Sub cmdNew_Click()
   EstiESe02h.Show
   
End Sub

Private Sub cmdOth_Click()
   Dim bByte As Byte
   bByte = CheckBidEntries(bGoodPart, bGoodCust)
   If bByte = 1 Then Exit Sub
   
   On Error Resume Next
   EstiESe02e.txtSet = Format(RdoFull!BIDSETUP, "####0.00")
   EstiESe02e.txtEpd = Format(RdoFull!BIDEXPEDITE, "####0.00")
   EstiESe02e.txtTlg = Format(RdoFull!BIDTOOLING, "####0.00")
   EstiESe02e.txtDue = Format(RdoFull!BIDDUE, "mm/dd/yyyy")
   EstiESe02e.txtFst = Format(RdoFull!BIDFIRSTDELIVERY, "mm/dd/yyyy")
   EstiESe02e.txtByr = "" & Trim(RdoFull!BIDBUYER)
   EstiESe02e.txtEst = lblEstimator
   EstiESe02e.txtCmt = "" & Trim(RdoFull!BIDCOMMENT)
   EstiESe02e.Show
   
End Sub

Private Sub cmdPrt_Click()
   optPart.Value = vbChecked
   cmbBid.Enabled = False
   EstiESe01p.optFull.Value = vbChecked
   EstiESe01p.txtPrt = txtPrt
   EstiESe01p.Show
   
End Sub

Private Sub cmdSrv_Click()
   Dim bByte As Byte
   bByte = CheckBidEntries(bGoodPart, bGoodCust)
   If bByte = 1 Then Exit Sub
   
   EstiESe02d.lblBid = cmbBid
   EstiESe02d.Show
   
End Sub

Private Sub cmdVew_Click()
   optSle.Value = vbChecked
   ViewSales.lblPrt = txtPrt
   ViewSales.Show
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   sCurrEstimator = GetSetting("Esi2000", "EsiEngr", "Estimator", sCurrEstimator)
   If bOnLoad Then
      OpenBoxes True
      GetNextBid Me
      GetRates
      FillCustomers
      FillCustomerRFQs Me, cmbCst, True
      FillCombo
      
      Dim bPartSearch As Boolean
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillPartCombo cmbPrt
      
      cmbPrt = txtPrt
      GetBidLabor Compress(txtPrt), Val(cmbBid), CCur("0" & txtQty)
      bOnLoad = 0
   Else
      If optFrom.Value = vbChecked Then
         Unload EstiESe01c
         optFrom.Value = vbUnchecked
      End If
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   'lblEstGrandTotal.ForeColor = ES_BLUE
   bOnLoad = 1
   bNoEdit = True
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bResponse As Byte
   If Trim(cmbCst) = "" Or Trim(txtPrt) = "" Then
      bResponse = MsgBox("The Current Estimate Is Missing a Part Number And/Or " & vbCrLf _
                  & "Customer And Will Be Removed. Do You Still Wish To Quit?", _
                  ES_NOQUESTION, Caption)
      If bResponse = vbNo Then
         Cancel = True
      Else
         bGoodList = GetBidPartsList()
         If bGoodList = 0 Then
            bResponse = GetBidRouting()
            If bResponse = 0 Then
               sSql = "DELETE FROM EstiTable WHERE BIDREF=" & Val(cmbBid) & " "
               clsADOCon.ExecuteSQL sSql 'rdExecDirect
            End If
         End If
      End If
   Else
      If optPart.Value = vbChecked Then Unload EstiESe01p
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   On Error Resume Next
   RdoFull.Close
   Set EstiESe02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   Dim bScrap As Byte
   Dim bGna As Byte
   Dim bProfit As Byte
   
   bOnLoad = 1
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   For b = 65 To 88
      cmbCls.AddItem Chr$(b)
   Next
   cmbCls.AddItem Chr$(b)
   cmbCls = "F"
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtDte.ToolTipText = "The Date Of The Current Bid. Default Today For New Bids."
   txtDue = Format(ES_SYSDATE + 10, "mm/dd/yyyy")
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
   lblUnitTotal.BackColor = Es_TextDisabled
   
   If Trim(txtFst) = "" Then txtFst = Format(Now + 60, "mm/dd/yyyy")
   txtEstimator.BackColor = &H8000000F
End Sub

Private Sub FillCombo()
   cmbBid.Clear
   On Error GoTo DiaErr1
   FillEstimateCombo Me, "FULL"
   If cmbBid.ListCount > 0 Then
      cmbBid = cmbBid.List(0)
      bGoodBid = GetThisFullBid(True)
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

Public Function GetThisFullBid(Optional bHideMsg As Boolean) As Byte
   Dim b As Byte
   Dim lBidNo As Long
   Dim BidQty As Currency
   
   lBidNo = Val(cmbBid)
   bGoodBid = 0
   sRouting = ""
   GetThisFullBid = 0
   lblRouting = ""
      
   bNoEdit = True
   txtEstimator.BackColor = &H8000000F
   
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM EstiTable WHERE BIDREF=" & lBidNo & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFull, ES_KEYSET)
   If bSqlRows Then
      bOpenKey = 1
      With RdoFull
         If !BIDCANCELED = 1 Then
            cmdCom.Enabled = False
            ClearResultSet RdoFull
            OpenBoxes False
            bCanceled = 1
            Exit Function
         End If
         If !BIDACCEPTED = 1 Then
            cmdCom.Enabled = False
            If Not bHideMsg Then
               MsgBox "Bid " & cmbBid & " Was Accepted And Cannot Be Edited.", _
                  vbInformation
            End If
            ClearResultSet RdoFull
            OpenBoxes False
            bCanceled = 1
            Exit Function
         End If
         bCanceled = 0
         If !BIDCOMPLETE = 1 Then
            OpenBoxes False
            cmdCom.Enabled = True
            cmdCom.Caption = "Unc&omplete"
            cmdCom.ToolTipText = "Remove Complete Flag From Bid"
            optCom.Value = vbChecked
         Else
            OpenBoxes True
            cmdCom.Caption = "C&omplete"
            cmdCom.ToolTipText = "Mark Bid As Completed"
            cmdCom.Enabled = True
            txtPrt.Enabled = True
            cmbCst.Enabled = True
            txtQty.Enabled = True
            optCom.Value = vbUnchecked
         End If
         If Trim(!BidClass) = "QWIK" Then
            MsgBox "Estimate Is A Qwik Bid.  Edit From The Qwik Bid Area.", _
               vbInformation, Caption
            ClearResultSet RdoFull
            OpenBoxes False
            bCanceled = 1
            Exit Function
         End If
         cmbCls = "" & Trim(!BIDPRE)
         txtPrt = "" & Trim(!BidPart)
         cmbPrt = "" & Trim(!BidPart)
         bGoodPart = GetBidPart(Me)
         If "" & Trim(!BIDCUST) <> "" Then cmbCst = "" & Trim(!BIDCUST)
         bGoodCust = GetBidCustomer(Me, cmbCst)
         txtDte = Format(!BIDDATE, "mm/dd/yyyy")
         If Not IsNull(!BIDDUE) Then
            txtDue = Format(!BIDDUE, "mm/dd/yyyy")
         Else
            txtDue = Format(!BIDDATE + 10, "mm/dd/yyyy")
         End If
         If !BidQuantity > 0 Then
            txtQty = Format(!BidQuantity, ES_QuantityDataFormat)
         Else
            txtQty = "1.000"
         End If
         BidQty = txtQty
         cOldQty = Val(txtQty)
         
         
         lblFohRate = Format(!BIDFOHRATE, ES_MoneyFormat)
         
         cBGnaRate = !BIDGNARATE
         cBProfitRate = !BIDPROFITRATE
         cBprofit = !BIDPROFIT
         cBScrapRate = !BIDSCRAPRATE
         
         ' MM 2/13/2009 - Not need to set the global setting.
         ' Value need to be set from the database.
         
'         If !BIDGNARATE = 0 Then cBGnaRate = cGna _
'                          Else cBGnaRate = !BIDGNARATE
'
'         If !BIDPROFITRATE = 0 Then cBProfitRate = cProfit _
'                             Else cBProfitRate = !BIDPROFITRATE
'
'         If !BIDPROFIT = 0 Then cBprofit = cProfit _
'                        Else cBprofit = !BIDPROFIT
'
'         If !BIDSCRAPRATE = 0 Then cBScrapRate = cScrap _
'                           Else cBScrapRate = !BIDSCRAPRATE
         
         If Trim(!BIDESTIMATOR) = "" Then
            sBEstimator = sEstimator
         Else
            sBEstimator = "" & Trim(!BIDESTIMATOR)
         End If
         txtEstimator = sBEstimator
         lblEstimator = sBEstimator
         
         'Labor
         lblHours = Format(!BIDHOURS, ES_QuantityDataFormat)
         'lblUnitLaborCost
         lblUnitOverheadCost = Format(!BIDFOH, ES_MoneyFormat)
         lblRate = Format(!BIDRATE, ES_MoneyFormat)
         LblUnitLabor = Format(!BIDTOTLABOR, ES_MoneyFormat)
         lblEstTotalLabor = Format(BidQty * !BIDTOTLABOR, ES_MoneyFormat)
         
         'material
         lblMaterial = Format(!BIDMATL, ES_MoneyFormat)
         lblBurden = Format(!BIDBURDEN, ES_MoneyFormat)
         lblTotMat = Format(!BIDTOTMATL, ES_MoneyFormat)
         lblEstTotMatl = Format(BidQty * !BIDTOTMATL, ES_MoneyFormat)
         
         'services
         lblTotServices = Format(!BIDOSP, ES_MoneyFormat)
         'lblOther = Format(!BIDSETUP + !BIDEXPEDITE + !BIDTOOLING, "#####0.00")
         
         'setup, expediting, tooling
         txtSet = Format(!BIDSETUP, "#####0.00")
         txtEpd = Format(!BIDEXPEDITE, "#####0.00")
         txtTlg = Format(!BIDTOOLING, "#####0.00")
         TotalSetExpTool
         
         'percentages
         txtScr = Format(cBScrapRate * 100, ES_MoneyFormat)
         txtGna = Format(cBGnaRate * 100, ES_MoneyFormat)
         txtPrf = Format(cBProfitRate * 100, ES_MoneyFormat)
         lblScrap = Format(!BIDSCRAP, ES_MoneyFormat)
         lblGNA = Format(!BIDGNA, ES_MoneyFormat)
         lblPrf = Format(!BIDPROFIT, ES_MoneyFormat)
         
         FindCustomer Me, cmbCst
         FillCustomerRFQs Me, cmbCst, True
         cmbRfq = "" & Trim(!BIDRFQ)
         If Trim(cmbRfq) = "" Then cmbRfq = "NONE"
         lblUnitTotal = Format(!BIDUNITPRICE, "#####0.00")
         Me.txtCmt = !BIDCOMMENT
         Me.txtByr = Trim(!BIDBUYER)
         If IsNull(!BIDFIRSTDELIVERY) Then
            txtFst = Format(Now, "MM/DD/yyyy")
         Else
            txtFst = Format(!BIDFIRSTDELIVERY, "MM/DD/yyyy")
         End If
         
         
         Me.txtLeadTime = IIf(IsNull(!BIDLEADTIME), "", Trim(!BIDLEADTIME))
         Me.txtFOB = IIf(IsNull(!BIDFOB), "", Trim(!BIDFOB))
         
         If IsNull(!BIDQUOTEDATE) Then
            txtQuoteDt = Format(Now, "MM/DD/YY")
         Else
            txtQuoteDt = Format(!BIDQUOTEDATE, "MM/DD/YY")
         End If
                  
         b = GetBidServices(EstiESe02a, CCur("0" & txtQty))
         UnitPrice
      End With
      GetThisFullBid = 1
   Else
      'Reset defaults
      cBGnaRate = cGna
      cBProfitRate = cProfit
      cBScrapRate = cScrap
      cBprofit = 0
      GetThisFullBid = 0
   End If
   bOpenKey = 0
   GetBidLabor Compress(txtPrt), Val(cmbBid), CCur("0" & txtQty)
   Exit Function
   
DiaErr1:
   sProcName = "getthisbid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

'DOES NOTHING
'Private Sub lblBurden_Change()
'   If bGoodBid Then
'      On Error Resume Next
'      'RdoFull.Edit
'      RdoFull.Update
'   End If
'
'End Sub

Private Sub lblGNA_Change()
   Dim cGna As Currency
   Dim cGnr As Currency
   cGna = Format(Val(lblGNA), ES_QuantityDataFormat)
   cGnr = Format(Val(txtGna) / 100, ES_QuantityDataFormat)
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDGNA = Format(cGna, ES_QuantityDataFormat)
      RdoFull!BIDGNARATE = Format(cGnr, ES_QuantityDataFormat)
      RdoFull.Update
   End If
   
End Sub

Private Sub lblUnitLabor_Change()
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDRATE = Format(Val(lblRate), ES_QuantityDataFormat)
      RdoFull!BIDTOTLABOR = Format(Val(LblUnitLabor), ES_QuantityDataFormat)
      RdoFull!BIDHOURS = Format(Val(lblHours), ES_QuantityDataFormat)
      RdoFull!BIDFOH = Format(Val(lblUnitOverheadCost), ES_QuantityDataFormat)
      RdoFull!BIDFOHRATE = Format(Val(lblFohRate), ES_QuantityDataFormat)
      RdoFull.Update
      UnitPrice
   End If
   
End Sub

Private Sub lblOther_Change()
   TotalBid
   
End Sub

Private Sub lblPrf_Change()
   If bOpenKey = 1 Then Exit Sub
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDPROFIT = Format(Val(lblPrf), ES_QuantityDataFormat)
      RdoFull!BIDPROFITRATE = Format(Val(txtPrf) / 100, ES_QuantityDataFormat)
      RdoFull.Update
   End If
   
End Sub

Private Sub lblRouting_Change()
   sRouting = lblRouting
   
End Sub

Private Sub lblScrap_Change()
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDSCRAP = Format(Val(lblScrap), ES_QuantityDataFormat)
      RdoFull!BIDSCRAPRATE = Format(Val(txtScr) / 100, ES_QuantityDataFormat)
      RdoFull.Update
   End If
   
End Sub

Private Sub lblTotMat_Change()
   Dim cMat As Currency
   Dim cBur As Currency
   Dim cTot As Currency
   Dim cBbd As Currency
   If bGoodBid Then
      On Error Resume Next
      cTot = Format(Val(lblTotMat), ES_QuantityDataFormat)
      cMat = Format(Val(lblMaterial), ES_QuantityDataFormat)
      cBur = Format(Val(lblBurden), ES_QuantityDataFormat)
      'RdoFull.Edit
      RdoFull!BIDTOTMATL = Format(cTot, ES_QuantityDataFormat)
      RdoFull!BIDMATL = Format(cMat, ES_QuantityDataFormat)
      RdoFull!BIDBURDEN = Format(cBur, ES_QuantityDataFormat)
      If cMat > 0 Then
         cBbd = cBur / cMat
         RdoFull!BIDBURDENRATE = Format(cBbd, ES_QuantityDataFormat)
      Else
         RdoFull!BIDBURDENRATE = 0
      End If
      If Val(lblTotMat) > 0 Then
         RdoFull!BIDMATLDESC = "Total From Parts List"
      Else
         RdoFull!BIDMATLDESC = "No Parts List"
      End If
      RdoFull.Update
      UnitPrice
   End If
   
End Sub

Private Sub lblTotServices_Change()
   '
End Sub

Private Sub lblUnitServices_Change()
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDOSP = Val(lblUnitServices)
      If Val(lblUnitServices) > 0 Then
         RdoFull!BIDOSPDESC = "Total Entries"
      Else
         RdoFull!BIDOSPDESC = "No Services Entered"
      End If
      RdoFull.Update
    
      UnitPrice
   
   End If
   'MM 2/13/2009 Errr: should be inside the if statement as
   ' this will reset the Material cost.
   ' UnitPrice
   
End Sub

Private Sub optVew_Click()
   If optVew.Value = vbUnchecked Then txtPrt.SetFocus
   
End Sub


Private Sub Timer1_Timer()
   GetNextBid Me
   
End Sub

Private Sub txtByr_LostFocus()
   txtByr = CheckLen(txtByr, 30)
   txtByr = StrCase(txtByr)
   If bGoodBid Then
      'On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDBUYER = Trim(txtByr)
      RdoFull.Update
   End If
End Sub

Private Sub txtLeadTime_LostFocus()
   txtLeadTime = CheckLen(txtLeadTime, 128)
   txtLeadTime = UCase(txtLeadTime)
   If bGoodBid Then
      'On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDLEADTIME = Trim(txtLeadTime)
      RdoFull.Update
   End If
End Sub


Private Sub txtFOB_LostFocus()
   txtFOB = CheckLen(txtFOB, 128)
   txtFOB = UCase(txtFOB)
   If bGoodBid Then
      'On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDFOB = Trim(txtFOB)
      RdoFull.Update
   End If
End Sub



Private Sub txtCmt_Change()
   'txtCmt = CheckLen(txtCmt, 1024)
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDCOMMENT = Trim(txtCmt)
      RdoFull.Update
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
      'RdoFull.Edit
      RdoFull!BIDDATE = Format(txtDte, "mm/dd/yyyy")
      RdoFull.Update
   End If
   
End Sub


Private Sub txtDue_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtDue_LostFocus()
   txtDue = CheckDateEx(txtDue)
   cmbBid.Enabled = True
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDDUE = Format(txtDue, "mm/dd/yyyy")
      RdoFull.Update
   End If
   
End Sub

Private Sub txtEstimator_DblClick()
    bNoEdit = False
    txtEstimator.BackColor = &H80000005
    
End Sub

Private Sub txtEstimator_KeyPress(KeyAscii As Integer)
    If bNoEdit Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtEstimator_LostFocus()
    If bGoodBid Then
        On Error Resume Next
        RdoFull!BIDESTIMATOR = CheckLen(txtEstimator, 30)
        RdoFull.Update
    End If
    
End Sub
Private Sub txtFst_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub txtFst_LostFocus()
   If Trim(txtFst) = "" Then
      If bGoodBid Then
         'On Error Resume Next
         'RdoFull.Edit
         RdoFull!BIDFIRSTDELIVERY = Null
         RdoFull.Update
      End If
   Else
      txtFst = CheckDateEx(txtFst)
      If bGoodBid Then
         On Error Resume Next
         'RdoFull.Edit
         RdoFull!BIDFIRSTDELIVERY = Format(txtFst, "mm/dd/yyyy")
         RdoFull.Update
      End If
   End If
   
End Sub

Private Sub txtQuoteDt_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtQuoteDt_LostFocus()
   If Trim(txtQuoteDt) = "" Then
      If bGoodBid Then
         'On Error Resume Next
         'RdoFull.Edit
         RdoFull!BIDQUOTEDATE = Null
         RdoFull.Update
      End If
   Else
      txtQuoteDt = CheckDate(txtQuoteDt)
      If bGoodBid Then
         On Error Resume Next
         'RdoFull.Edit
         RdoFull!BIDQUOTEDATE = Format(txtQuoteDt, "mm/dd/yy")
         RdoFull.Update
      End If
   End If
   
End Sub

Private Sub txtGna_LostFocus()
   txtGna = Format(Abs(Val(txtGna)), "#####0.000")
   UnitPrice
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDGNARATE = Format(Val(txtGna) / 100)
      RdoFull.Update
   End If
   UnitPrice
   
End Sub


Private Sub lblunittotal_Change()
   If bOpenKey = 1 Or bOnLoad = 1 Then Exit Sub
   If bGoodBid = 1 Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDUNITPRICE = Format(Val(lblUnitTotal), ES_QuantityDataFormat)
      RdoFull.Update
      UpdateDiscounts
   End If
End Sub

Private Sub txtPrf_LostFocus()
   txtPrf = CheckLen(txtPrf, 7)
   txtPrf = Format(Abs(Val(txtPrf)), "##0.000")
   UnitPrice
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDPROFITRATE = Format(Val(txtPrf) / 100)
      RdoFull.Update
   End If
   UnitPrice
   
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
      EstiESe01p.optFull.Value = vbUnchecked
      EstiESe01p.txtPrt = txtPrt
      EstiESe01p.Show
   End If
   
End Sub

Private Sub cmbPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      optPart.Value = vbChecked
      cmbBid.Enabled = False
      EstiESe01p.optFull.Value = vbUnchecked
      EstiESe01p.txtPrt = cmbPrt
      EstiESe01p.Show
   End If
   
End Sub
Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   bGoodPart = GetBidPart(Me)
   If bGoodPart Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BidPart = Compress(txtPrt)
      If bGoodCust Then
         RdoFull!BIDLOCKED = 0
      Else
         RdoFull!BIDLOCKED = 1
      End If
      RdoFull.Update
   Else
      '   MsgBox "Bids Without A Valid PartNumber Will Not Be Saved.", _
      '       vbInformation, Caption
   End If
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If

   txtPrt = cmbPrt
   bGoodPart = GetBidPart(Me)
   If bGoodPart Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BidPart = Compress(cmbPrt)
      If bGoodCust Then
         RdoFull!BIDLOCKED = 0
      Else
         RdoFull!BIDLOCKED = 1
      End If
      RdoFull.Update
   Else
      '   MsgBox "Bids Without A Valid PartNumber Will Not Be Saved.", _
      '       vbInformation, Caption
   End If
   
End Sub


Private Sub OpenBoxes(bOpen As Boolean)
   Dim iList As Integer
   
   lblHours = ES_MoneyFormat
   lblUnitLaborCost = ES_MoneyFormat
   lblUnitOverheadCost = ES_MoneyFormat
   LblUnitLabor = ES_MoneyFormat
   lblEstTotalLabor = ES_MoneyFormat
   
   lblMaterial = ES_MoneyFormat
   lblBurden = ES_MoneyFormat
   lblTotMat = ES_MoneyFormat
   lblEstTotMatl = ES_MoneyFormat
   
   lblTotServices = ES_MoneyFormat
   lblUnitServices = ES_MoneyFormat
   'lblEstTotService = ES_MoneyFormat
   
   txtSet = ES_MoneyFormat
   txtEpd = ES_MoneyFormat
   txtTlg = ES_MoneyFormat
   'lblOther = ES_MoneyFormat
   lblEstTotSET = ES_MoneyFormat
   
   lblScrap = ES_MoneyFormat
   lblEstTotScrap = ES_MoneyFormat
   
   lblGNA = ES_MoneyFormat
   lblEstTotGA = ES_MoneyFormat
   
   lblPrf = ES_MoneyFormat
   lblEstTotProfit = ES_MoneyFormat
   
   lblEstGrandTotal = ES_MoneyFormat
   
   txtPrf = ES_MoneyFormat
   lblUnitTotal = ES_MoneyFormat
   
   If bOpen Then
      For iList = 0 To Controls.Count - 1
         If TypeOf Controls(iList) Is TextBox Or TypeOf Controls(iList) Is ComboBox _
                              Or TypeOf Controls(iList) Is CommandButton Or TypeOf Controls(iList) _
                              Is ComboBox Then
            Controls(iList).Enabled = True
         End If
      Next
   Else
      For iList = 0 To Controls.Count - 1
         If TypeOf Controls(iList) Is TextBox Or TypeOf Controls(iList) Is ComboBox _
                              Or TypeOf Controls(iList) Is CommandButton Or TypeOf Controls(iList) _
                              Is ComboBox Then
            Controls(iList).Enabled = False
         End If
      Next
   End If
   cmdCan.Enabled = True
   cmbBid.Enabled = True
   cmdNew.Enabled = True
   cmdVew.Enabled = True
   
End Sub

Private Sub UnitPrice()
   Dim cQuantity As Currency
   Dim cUnitLbr As Currency
   Dim cUnitMtl As Currency
   Dim cUnitOsp As Currency
   Dim cUnitPrc As Currency
   Dim cScrap As Currency
   Dim cGenAdm As Currency
   Dim cProfit As Currency
   
   cBGnaRate = Val(txtGna) / 100
   cBProfitRate = Val(txtPrf) / 100
   cScrap = Val(txtScr)
   
   cQuantity = cOldQty
   If cQuantity = 0 Then cQuantity = 1
   cUnitLbr = Val(LblUnitLabor)
   cUnitMtl = Val(lblTotMat)
   cUnitOsp = Val(lblUnitServices)
   cScrap = (cUnitLbr + cUnitMtl + cUnitOsp) * (cScrap / 100)
   lblScrap = Format(cScrap, "####0.00")
   cUnitPrc = cUnitLbr + cUnitMtl + cUnitOsp + cScrap
   cGenAdm = cUnitPrc * cBGnaRate
   lblGNA = Format(cGenAdm, "####0.00")
   cUnitPrc = cUnitPrc + cGenAdm
   cProfit = cUnitPrc * cBProfitRate
   cUnitPrc = cUnitPrc + cProfit
   lblPrf = Format(cProfit, "####0.00")
   lblUnitTotal = Format(cUnitPrc, "####0.00")
   TotalBid
   
   
End Sub


Private Function GetPartRouting() As Byte
   Dim RdoRte As ADODB.Recordset
   sSql = "SELECT PARTREF,PAROUTING FROM PartTable Where PARTREF='" _
          & Compress(txtPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      With RdoRte
         If Trim(!PAROUTING) = "" Then
            GetPartRouting = 0
            sRouting = ""
            lblRouting = ""
         Else
            sRouting = "" & Trim(!PAROUTING)
            lblRouting = sRouting
            GetPartRouting = 1
         End If
         ClearResultSet RdoRte
      End With
   Else
      GetPartRouting = 0
   End If
   Set RdoRte = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getbidrout"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub CopyARouting()
   Dim RdoRte As ADODB.Recordset
   Dim sMsg As String
   On Error GoTo DiaErr1
   
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   sSql = "SELECT OPREF,OPNO,OPSHOP,OPCENTER,OPSETUP,OPUNIT," & vbCrLf _
      & "OPQHRS,OPMHRS,WCNREF,WCNOHPCT,WCNESTRATE" & vbCrLf _
      & "FROM RtopTable" & vbCrLf _
      & "JOIN WcntTable ON OPSHOP=WCNSHOP AND OPCENTER=WCNREF" & vbCrLf _
      & "WHERE OPREF='" & sRouting & "' ORDER BY OPNO"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_STATIC)
   If bSqlRows Then
      MouseCursor 13
      With RdoRte
         'On Error Resume Next
         Do Until .EOF
'            sSql = "INSERT INTO EsrtTable" & vbCrLf _
'               & "(BIDRTEREF,BIDRTEOPNO,BIDRTESHOP,BIDRTECENTER," & vbCrLf _
'                  & "BIDRTESETUP,BIDRTEUNIT,BIDRTEQHRS," & vbCrLf _
'                  & "BIDRTEMHRS,BIDRTERATE,BIDRTEFOHRATE)" & vbCrLf _
'                  & "VALUES(" & cmbBid & "," & vbCrLf _
'                  & Str$(!OPNO) & ",'" & Trim(!OPSHOP) & "','" & Trim(!OPCENTER) & "'," & vbCrLf _
'                  & Format$(!OPSETUP) & "," & Format$(!OPUNIT) & "," & Format$(!OPQHRS) & "," & vbCrLf _
'                  & Format$(!OPMHRS) & "," & Format$(!WCNESTRATE) & "," & Format$(!WCNOHPCT / 100) & ")"
            sSql = "INSERT INTO EsrtTable" & vbCrLf _
               & "(BIDRTEREF,BIDRTEOPNO,BIDRTESHOP,BIDRTECENTER," & vbCrLf _
                  & "BIDRTESETUP,BIDRTEUNIT,BIDRTEQHRS," & vbCrLf _
                  & "BIDRTEMHRS,BIDRTERATE,BIDRTEFOHRATE)" & vbCrLf _
                  & "VALUES(" & cmbBid & "," & vbCrLf _
                  & !OPNO & ",'" & Trim(!OPSHOP) & "','" & Trim(!OPCENTER) & "'," & vbCrLf _
                  & !OPSETUP & "," & !OPUNIT & "," & !OPQHRS & "," & vbCrLf _
                  & !OPMHRS & "," & !WCNESTRATE & "," & !WCNOHPCT / 100 & ")"
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
            .MoveNext
         Loop
         MouseCursor 0
         If clsADOCon.ADOErrNum = 0 Then
            clsADOCon.CommitTrans
            Sleep 500
            MsgBox "Your Routing Is Ready. Reselect Labor.", vbInformation, Caption
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            MsgBox "Could Not Copy The Routing.", vbExclamation, Caption
         End If
         ClearResultSet RdoRte
      End With
   End If
   Set RdoRte = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "CopyARouting"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetPartsList() As Byte
   'returns 1 if there is a current parts list for this part
   'sBomRev = the revision
   
   Dim RdoPls As ADODB.Recordset
   On Error GoTo DiaErr1
'   sSql = "SELECT DISTINCT BMASSYPART,BMREV FROM BmplTable WHERE " _
'          & "BMASSYPART='" & Compress(txtPrt) & "' AND BMREV='" _
'          & sBomRev & "'"
'   bSqlRows = clsADOCon.GetDataSet(sSql,RdoPls, ES_FORWARD)
'   If bSqlRows Then
'      GetPartsList = 1
'      ClearResultSet RdoPls
'   Else
'      GetPartsList = 0
'   End If
'   Set RdoPls = Nothing

   sSql = "SELECT BMREV FROM BmplTable" & vbCrLf _
      & "JOIN PartTable ON BMASSYPART = PARTREF" & vbCrLf _
      & "AND BMREV = PABOMREV" & vbCrLf _
      & "WHERE BMASSYPART='" & Compress(txtPrt) & "'"
   If clsADOCon.GetDataSet(sSql, RdoPls) Then
      sBomRev = Trim(RdoPls.Fields(0))
      GetPartsList = 1
   End If
   
   Exit Function
   
DiaErr1:
   sProcName = "getpartslist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

'Changed 8/11/03
'Added BOMLEVEL 8/1/03

Private Sub CopyPartsList()
   Dim RdoPls As ADODB.Recordset
   Dim b As Byte
   Dim sPartNumber As String
   Dim sPartRev As String
   
   On Error GoTo DiaErr1
   clsADOCon.BeginTrans
   
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV,BMPARTREV,BMSEQUENCE,BMQTYREQD," _
          & "BMUNITS,BMCONVERSION,BMADDER,BMSETUP,BMCOMT,BMESTLABOR," _
          & "BMESTLABOROH FROM BmplTable WHERE BMASSYPART='" & Compress(txtPrt) _
          & "' AND BMREV='" & sBomRev & "'" & vbCrLf _
          & "ORDER BY BMSEQUENCE, BMPARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPls, ES_STATIC)
   If bSqlRows Then
      MouseCursor 13
      With RdoPls
         'On Error Resume Next
         
         Do Until .EOF
            b = b + 1
            'bBomLevel = 1
            sPartNumber = "" & Trim(!BMPARTREF)
            sPartRev = "" & Trim(!BMPARTREV)
            sSql = "INSERT INTO EsbmTable (BIDBOMREF,BIDBOMASSYPART," & vbCrLf _
                   & "BIDBOMSEQUENCE,BIDBOMPARTREF,BIDBOMLEVEL,BIDBOMQTYREQD,BIDBOMUNITS," & vbCrLf _
                   & "BIDBOMCONVERSION,BIDBOMADDER,BIDBOMSETUP," & vbCrLf _
                   & "BIDBOMCOMT,BIDBOMLABOR,BIDBOMLABOROH)" & vbCrLf _
                   & "VALUES(" & Val(cmbBid) & "," & vbCrLf _
                   & "'" & Compress(txtPrt) & "'," & vbCrLf _
                   & b & "," & vbCrLf _
                   & "'" & Trim(!BMPARTREF) & "'," & vbCrLf _
                   & "1," & vbCrLf _
                   & Trim(!BMQTYREQD) & "," & vbCrLf _
                   & "'" & Trim(!BMUNITS) & "'," & vbCrLf _
                   & Trim(!BMCONVERSION) & "," & vbCrLf _
                   & Trim(!BMADDER) & "," & vbCrLf _
                   & Trim(!BMSETUP) & "," & vbCrLf _
                   & "'" & Trim("" & !BMCOMT) & "'," & vbCrLf _
                   & Trim(!BMESTLABOR) & "," & vbCrLf _
                   & Trim(!BMESTLABOROH) & ")"
                        
                  Debug.Print sSql
                  
                  clsADOCon.ExecuteSQL sSql 'rdExecDirect
            'GetNextBomLevel sPartNumber, sPartRev, 2
            GetNextBomLevel sPartNumber, sBomRev, 2
            .MoveNext
         Loop
         ClearResultSet RdoPls
         MouseCursor 0
'         If Err = 0 Then
            clsADOCon.CommitTrans
            MsgBox "Your Bill Of Material Is Ready. Reselect Material.", _
               vbInformation, Caption
'         Else
'            MsgBox "Could Not Copy Some Of The Bill " & vbCrLf _
'               & "Possibly A Duplicate Or Circle." & vbCrLf _
'               & "Double Check Your Bill Of Material.", _
'               vbExclamation, Caption
'         End If
      End With
   End If
   Set RdoPls = Nothing
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   sProcName = "CopyPartsList"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtQty_LostFocus()
   Dim b As Byte
   txtQty = CheckLen(txtQty, 10)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   If Val(txtQty) = 0 Then txtQty = "1.000"
   If bGoodBid Then
      'On Error Resume Next
      'RdoFull.Edit
      RdoFull!BidQuantity = Val(txtQty)
      RdoFull.Update
      If cOldQty <> Val(txtQty) Then
         b = GetBidLabor(Compress(txtPrt), Val(cmbBid), CCur("0" & txtQty))
         TotalBidMatl CLng(cmbBid), Compress(txtPrt), CCur("0" & txtQty)
         b = GetBidServices(EstiESe02a, CCur("0" & txtQty))
         TotalSetExpTool
         TotalBid
      End If
   End If
   cOldQty = Val(txtQty)
   UnitPrice
   
End Sub



Private Sub TotalBid()
   Dim cUnitTot As Currency, cBidTot As Currency
   Dim unitScrap As Currency, unitGNA As Currency, unitProfit As Currency
   Dim qty As Currency
   
   'unit total
   qty = String2Currency(txtQty)
   'cUnitTot = String2Currency(LblUnitLabor) + String2Currency(lblTotMat) _
   '   + String2Currency(lblUnitServices) + String2Currency(lblUnitSET)
   cUnitTot = String2Currency(LblUnitLabor) + String2Currency(lblTotMat) _
      + String2Currency(lblUnitServices)
   
   lblPreScrapUnit = Format(cUnitTot, ES_MoneyFormat)
   unitScrap = String2Currency(txtScr) * cUnitTot / 100
   lblScrap = Format(unitScrap, ES_MoneyFormat)
   lblEstTotScrap = Format(unitScrap * qty, ES_MoneyFormat)
   cUnitTot = cUnitTot + unitScrap
   
   lblPreGNAUnit = Format(cUnitTot, ES_MoneyFormat)
   unitGNA = String2Currency(txtGna) * cUnitTot / 100
   lblGNA = Format(unitGNA, ES_MoneyFormat)
   lblEstTotGA = Format(unitGNA * qty, ES_MoneyFormat)
   cUnitTot = cUnitTot + unitGNA
   
   lblPreProfitUnit = Format(cUnitTot, ES_MoneyFormat)
   unitProfit = String2Currency(txtPrf) * cUnitTot / 100
   lblPrf = Format(unitProfit, ES_MoneyFormat)
   lblEstTotProfit = Format(unitProfit * qty, ES_MoneyFormat)
   cUnitTot = cUnitTot + unitProfit
         
   lblUnitTotal = Format(cUnitTot, ES_MoneyFormat)
   
   'make sure grand total column balances
   lblEstTotalLabor = Format(qty * String2Currency(LblUnitLabor), ES_MoneyFormat)
   lblEstTotMatl = Format(qty * String2Currency(lblTotMat), ES_MoneyFormat)
   lblTotServices = Format(qty * String2Currency(lblUnitServices), ES_MoneyFormat)
   'lblEstTotSET = Format(qty * String2Currency(lblUnitSET), ES_MoneyFormat)
   lblEstTotScrap = Format(qty * String2Currency(lblScrap), ES_MoneyFormat)
   lblEstTotGA = Format(qty * String2Currency(lblGNA), ES_MoneyFormat)
   lblEstTotProfit = Format(qty * String2Currency(lblPrf), ES_MoneyFormat)
   lblEstTotLessSET = Format(qty * String2Currency(lblUnitTotal), ES_MoneyFormat)
   lblEstGrandTotal = Format(String2Currency(Me.lblEstTotLessSET) _
      + String2Currency(lblEstTotSET), ES_MoneyFormat)
   
   'bid total
'   cBidTot = String2Currency(lblEstTotalLabor) + String2Currency(lblEstTotMatl) _
'      + String2Currency(lblTotServices) + String2Currency(lblEstTotSET) _
'      + String2Currency(lblEstTotScrap) + String2Currency(lblEstTotGA) _
'      + String2Currency(lblEstTotProfit)
'   lblEstGrandTotal = Format(cBidTot, ES_MoneyFormat)
   
   'if totals have changed, update them
   If Not RdoFull Is Nothing And bOnLoad = 0 Then
      With RdoFull
         If !BIDUNITPRICE <> String2Currency(lblUnitTotal) Then
            'RdoFull.Edit
            RdoFull!BIDTOTLABOR = String2Currency(LblUnitLabor)
            RdoFull!BIDMATL = String2Currency(lblMaterial)
            RdoFull!BIDBURDEN = String2Currency(lblBurden)
            RdoFull!BIDTOTMATL = String2Currency(lblTotMat)
            RdoFull!BIDFOH = String2Currency(lblUnitOverheadCost)
            RdoFull!BIDUNITPRICE = String2Currency(lblUnitTotal)
            RdoFull.Update
         End If
      End With
   End If
End Sub

Private Sub UpdateDiscounts()
   Dim cPer As Currency
   Dim cPrc As Currency
   On Error Resume Next
   bGoodBid = RefreshKeySet
   If bGoodBid = 1 Then
      With RdoFull
         '.Edit
         cPer = (!BIDQTYDISC1 / 100)
         cPrc = (1 - cPer) * Val(lblUnitTotal)
         !BIDQTYPRICE1 = cPrc
         
         cPer = (!BIDQTYDISC2 / 100)
         cPrc = (1 - cPer) * Val(lblUnitTotal)
         !BIDQTYPRICE2 = cPrc
         
         cPer = (!BIDQTYDISC3 / 100)
         cPrc = (1 - cPer) * Val(lblUnitTotal)
         !BIDQTYPRICE3 = cPrc
         
         cPer = (!BIDQTYDISC4 / 100)
         cPrc = (1 - cPer) * Val(lblUnitTotal)
         !BIDQTYPRICE4 = cPrc
         
         cPer = (!BIDQTYDISC5 / 100)
         cPrc = (1 - cPer) * Val(lblUnitTotal)
         !BIDQTYPRICE5 = cPrc
         
         cPer = (!BIDQTYDISC6 / 100)
         cPrc = (1 - cPer) * Val(lblUnitTotal)
         !BIDQTYPRICE6 = cPrc
         .Update
      End With
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "updatedisc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'8/7/03 - Adjusts for lose of KeySet when adding bom items

Private Function RefreshKeySet() As Byte
   On Error Resume Next
   sSql = "SELECT * FROM EstiTable WHERE BIDREF=" & Val(cmbBid) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFull, ES_KEYSET)
   If bSqlRows Then
      RefreshKeySet = 1
   Else
      RefreshKeySet = 0
   End If
   
End Function

'ensueing levels 8/11/03
'From Copy Parts List

'Private Sub GetNextBomLevel(BomPartRef As String, BomPartRev As String, BomLevel As Integer)
Private Sub GetNextBomLevel(BomPartRef As String, BomRev As String, BomLevel As Integer)
   Dim RdoBm2 As ADODB.Recordset
   Dim b As Byte
   Dim sPartNumber As String
   Dim sPartRev As String
   
   'don't allow BOM infinite loop
   If BomLevel > 10 Then
      MsgBox "Warning: BOM more than 10 levels deep.  Part " & BomPartRef & " is part of BOM loop."
      Exit Sub
   End If
   
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV,BMPARTREV,BMSEQUENCE,BMQTYREQD," & vbCrLf _
      & "BMUNITS,BMCONVERSION,BMADDER,BMSETUP,BMCOMT,BMESTLABOR,BMESTLABOROH" & vbCrLf _
      & "FROM BmplTable WHERE BMASSYPART='" & BomPartRef & "'" & vbCrLf _
      & "AND BMREV='" & BomRev & "'"
Debug.Print sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBm2, ES_STATIC)
   If bSqlRows Then
      With RdoBm2
         'bBomLevel = bBomLevel + 1
         Do Until .EOF
            sPartNumber = "" & Trim(!BMPARTREF)
            sPartRev = "" & Trim(!BMPARTREV)
Debug.Print "Level " & BomLevel & ": " & sPartNumber & " BOM rev " & BomRev
            'check for loops in bom
            If sPartNumber = BomPartRef Then
               MsgBox "WARNING: Circular BOM ignored." & vbCrLf _
                  & "Part " & sPartNumber & " is on its own part list."
            Else
               
               Dim RdoSon As ADODB.Recordset
               sSql = "SELECT * FROM EsbmTable " _
                     & " WHERE BIDBOMREF = " & Val(cmbBid) _
                        & " AND BIDBOMASSYPART = '" & BomPartRef & "'" _
                        & " AND BIDBOMPARTREF = '" & Trim(!BMPARTREF) & "'" _
                        & " AND BIDBOMLEVEL = '" & BomLevel & "'"
                        
               bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
               If bSqlRows Then
                  ClearResultSet RdoSon
                  sSql = "UPDATE EsbmTable SET BIDBOMQTYREQD = BIDBOMQTYREQD + " & Trim(!BMQTYREQD) _
                         & ", BIDBOMADDER = BIDBOMADDER + " & Trim(!BMADDER) _
                          & ",BIDBOMSETUP = BIDBOMSETUP + " & Trim(!BMSETUP) _
                     & " WHERE BIDBOMREF = " & Val(cmbBid) _
                        & " AND BIDBOMASSYPART = '" & BomPartRef & "'" _
                        & " AND BIDBOMPARTREF = '" & Trim(!BMPARTREF) & "'" _
                        & " AND BIDBOMLEVEL = '" & BomLevel & "'"
                  Debug.Print sSql
                  clsADOCon.ExecuteSQL sSql ' rdExecDirect
                  
               Else
                  sSql = "INSERT INTO EsbmTable (BIDBOMREF,BIDBOMASSYPART," & vbCrLf _
                         & "BIDBOMSEQUENCE,BIDBOMPARTREF,BIDBOMLEVEL,BIDBOMQTYREQD,BIDBOMUNITS," & vbCrLf _
                         & "BIDBOMCONVERSION,BIDBOMADDER,BIDBOMSETUP," & vbCrLf _
                         & "BIDBOMCOMT,BIDBOMLABOR,BIDBOMLABOROH)" & vbCrLf _
                         & "VALUES(" & Val(cmbBid) & "," & vbCrLf _
                         & "'" & BomPartRef & "'," & vbCrLf _
                         & b & "," & vbCrLf _
                         & "'" & Trim(!BMPARTREF) & "'," & vbCrLf _
                         & BomLevel & "," & vbCrLf _
                         & Trim(!BMQTYREQD) & "," & vbCrLf _
                         & "'" & Trim(!BMUNITS) & "'," & vbCrLf _
                         & Trim(!BMCONVERSION) & "," & vbCrLf _
                         & Trim(!BMADDER) & "," & vbCrLf _
                         & Trim(!BMSETUP) & "," & vbCrLf _
                         & "'" & Trim("" & Compress(!BMCOMT)) & "'," & vbCrLf _
                         & Trim(!BMESTLABOR) & "," & vbCrLf _
                         & Trim(!BMESTLABOROH) & ")"
                  
                  Debug.Print sSql
                  clsADOCon.ExecuteSQL sSql ', rdExecDirect
               End If
               
               Set RdoSon = Nothing
                           
                           
                           
                           sSql = "INSERT INTO EsbmTable (BIDBOMREF,BIDBOMASSYPART," & vbCrLf _
                      & "BIDBOMSEQUENCE,BIDBOMPARTREF,BIDBOMLEVEL,BIDBOMQTYREQD,BIDBOMUNITS," & vbCrLf _
                      & "BIDBOMCONVERSION,BIDBOMADDER,BIDBOMSETUP," & vbCrLf _
                      & "BIDBOMCOMT,BIDBOMLABOR,BIDBOMLABOROH)" & vbCrLf _
                      & "VALUES(" & Val(cmbBid) & "," & vbCrLf _
                      & "'" & BomPartRef & "'," & vbCrLf _
                      & b & "," & vbCrLf _
                      & "'" & Trim(!BMPARTREF) & "'," & vbCrLf _
                      & BomLevel & "," & vbCrLf _
                      & Trim(!BMQTYREQD) & "," & vbCrLf _
                      & "'" & Trim(!BMUNITS) & "'," & vbCrLf _
                      & Trim(!BMCONVERSION) & "," & vbCrLf _
                      & Trim(!BMADDER) & "," & vbCrLf _
                      & Trim(!BMSETUP) & "," & vbCrLf _
                      & "'" & Trim("" & !BMCOMT) & "'," & vbCrLf _
                      & Trim(!BMESTLABOR) & "," & vbCrLf _
                      & Trim(!BMESTLABOROH) & ")"
               clsADOCon.ExecuteSQL sSql 'rdExecDirect
               GetNextBomLevel sPartNumber, sBomRev, BomLevel + 1
            End If
            .MoveNext
         Loop
         ClearResultSet RdoBm2
      End With
   End If
   
End Sub

Private Sub txtScr_LostFocus()
   txtScr = CheckLen(txtScr, 7)
   txtScr = Format(Abs(Val(txtScr)), ES_MoneyFormat)
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDSCRAPRATE = Format(Val(txtScr) / 100)
      RdoFull.Update
   End If
   UnitPrice
   
End Sub

'Public Sub UpdateLaborTotals()
'   Dim rdo As ADODB.Recordset
'   sSql = "select sum(..."
'
'   lblHours = 0
'   lblUnitLaborCost = 0
'   lblUnitOverheadCost = 0
'   lblUnitLabor = 0   'total unit labor
'   lblEstTotalLabor = 0
'End Sub


Private Sub txtSet_LostFocus()
   txtSet = CheckLen(txtSet, 9)
   txtSet = Format(Abs(Val(txtSet)), "####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDSETUP = Val(txtSet)
      RdoFull.Update
      TotalSetExpTool
   End If
End Sub

Private Sub txtEpd_LostFocus()
   txtEpd = CheckLen(txtEpd, 9)
   txtEpd = Format(Abs(Val(txtEpd)), "####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDEXPEDITE = Val(txtEpd)
      RdoFull.Update
      TotalSetExpTool
   End If
End Sub

Private Sub txtTlg_LostFocus()
   txtTlg = CheckLen(txtTlg, 7)
   txtTlg = Format(Abs(Val(txtTlg)), "####0.00")
   If bGoodBid Then
      On Error Resume Next
      'RdoFull.Edit
      RdoFull!BIDTOOLING = Val(txtTlg)
      RdoFull.Update
      TotalSetExpTool
   End If
End Sub

Private Sub TotalSetExpTool()
   Dim total As Currency
   total = String2Currency(txtSet) + String2Currency(txtEpd) + String2Currency(txtTlg)
   Me.lblEstTotSET = Format(total, ES_MoneyFormat)
'   If String2Currency(txtQty) = 0 Then
'      lblUnitSET = Format(0, ES_MoneyFormat)
'   Else
'      lblUnitSET = Format(total / String2Currency(txtQty), ES_MoneyFormat)
'   End If
   TotalBid
End Sub

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

