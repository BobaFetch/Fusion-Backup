VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaAPe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cash Disbursements"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboDueDate 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Tag             =   "4"
      ToolTipText     =   "Include invoices due by this date"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   21
      ToolTipText     =   "Check To Pay"
      Top             =   4800
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   23
      ToolTipText     =   "Check For Partial Payment"
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   7
      Left            =   5400
      TabIndex        =   25
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   5160
      Width           =   1100
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   30
      ToolTipText     =   "Check To Pay"
      Top             =   5880
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   9
      Left            =   6600
      TabIndex        =   32
      ToolTipText     =   "Check For Partial Payment"
      Top             =   5880
      Width           =   255
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   9
      Left            =   5400
      TabIndex        =   31
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   5880
      Width           =   1100
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   8
      Left            =   5400
      TabIndex        =   28
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   8
      Left            =   6600
      TabIndex        =   29
      ToolTipText     =   "Check For Partial Payment"
      Top             =   5520
      Width           =   255
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   27
      ToolTipText     =   "Check To Pay"
      Top             =   5520
      Width           =   255
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   6
      Left            =   5400
      TabIndex        =   22
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   7
      Left            =   6600
      TabIndex        =   26
      ToolTipText     =   "Check For Partial Payment"
      Top             =   5160
      Width           =   255
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   24
      ToolTipText     =   "Check To Pay"
      Top             =   5160
      Width           =   255
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Checking Account To Debit"
      Top             =   320
      Width           =   1455
   End
   Begin VB.CheckBox optPad 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "Check To Pay"
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox optPad 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "Check To Pay"
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox optPad 
      Caption         =   "12"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   88
      ToolTipText     =   "Check To Pay"
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   15
      ToolTipText     =   "Check To Pay"
      Top             =   4080
      Width           =   255
   End
   Begin VB.CheckBox optPad 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   18
      ToolTipText     =   "Check To Pay"
      Top             =   4440
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   87
      ToolTipText     =   "Check For Partial Payment"
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   11
      ToolTipText     =   "Check For Partial Payment"
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   6600
      TabIndex        =   14
      ToolTipText     =   "Check For Partial Payment"
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   6600
      TabIndex        =   17
      ToolTipText     =   "Check For Partial Payment"
      Top             =   4080
      Width           =   255
   End
   Begin VB.CheckBox optPar 
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   6600
      TabIndex        =   20
      ToolTipText     =   "Check For Partial Payment"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox txtMemo 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Update"
      Height          =   315
      Left            =   9120
      TabIndex        =   33
      ToolTipText     =   "Updates Check Amount To Total"
      Top             =   1740
      Width           =   855
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   3360
      TabIndex        =   80
      Top             =   6480
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Height          =   315
      Left            =   9120
      TabIndex        =   12
      ToolTipText     =   "Select Invoices For This Vendor"
      Top             =   660
      Width           =   875
   End
   Begin VB.CommandButton cmdCnl 
      Caption         =   " &Reselect"
      Height          =   315
      Left            =   9120
      TabIndex        =   34
      ToolTipText     =   "Cancel Operation"
      Top             =   1380
      Width           =   855
   End
   Begin VB.CommandButton cmdPst 
      Caption         =   "&Post"
      Height          =   315
      Left            =   9120
      TabIndex        =   35
      ToolTipText     =   "Post This Check"
      Top             =   1020
      Width           =   855
   End
   Begin Threed.SSFrame z2 
      Height          =   30
      Left            =   120
      TabIndex        =   75
      Top             =   2640
      Width           =   9795
      _Version        =   65536
      _ExtentX        =   17277
      _ExtentY        =   53
      _StockProps     =   14
      Caption         =   "SSFrame1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   5
      Left            =   5400
      TabIndex        =   19
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   4440
      Width           =   1100
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   5400
      TabIndex        =   16
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   4080
      Width           =   1100
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   5400
      TabIndex        =   13
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   3720
      Width           =   1100
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   5400
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   3360
      Width           =   1100
   End
   Begin VB.TextBox txtPad 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   5400
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "Amount To Be Paid"
      Top             =   3000
      Width           =   1100
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      TabIndex        =   5
      Tag             =   "1"
      Top             =   1860
      Width           =   1215
   End
   Begin VB.TextBox txtChk 
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Tag             =   "3"
      Top             =   1860
      Width           =   1275
   End
   Begin VB.ComboBox cboCheckDate 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Tag             =   "4"
      ToolTipText     =   "Show discounts available through this date"
      Top             =   1860
      Width           =   1095
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Contains Vendors With PO's Unpaid Invoices"
      Top             =   720
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   9120
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   60
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6825
      FormDesignWidth =   10185
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   84
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPe03a.frx":0000
      PictureDn       =   "diaAPe03a.frx":0146
   End
   Begin Threed.SSCommand cmdDn 
      Height          =   375
      Left            =   7200
      TabIndex        =   85
      TabStop         =   0   'False
      ToolTipText     =   "Next Page (Page Down)"
      Top             =   6360
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaAPe03a.frx":028C
   End
   Begin Threed.SSCommand cmdUp 
      Height          =   375
      Left            =   7560
      TabIndex        =   86
      TabStop         =   0   'False
      ToolTipText     =   "Last Page (Page Up)"
      Top             =   6360
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaAPe03a.frx":078E
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      Height          =   285
      Index           =   19
      Left            =   240
      TabIndex        =   123
      Top             =   1500
      Width           =   1005
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disc Date   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   7980
      TabIndex        =   122
      Top             =   2760
      Width           =   825
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   7980
      TabIndex        =   121
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   7980
      TabIndex        =   120
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   7980
      TabIndex        =   119
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   7980
      TabIndex        =   118
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   7980
      TabIndex        =   117
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   7980
      TabIndex        =   116
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   7980
      TabIndex        =   115
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   7980
      TabIndex        =   114
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblDiscDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   7980
      TabIndex        =   113
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   3120
      TabIndex        =   112
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   600
      TabIndex        =   111
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   8940
      TabIndex        =   110
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   4200
      TabIndex        =   109
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label lblDis 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   6960
      TabIndex        =   108
      ToolTipText     =   "Tax And Freight Are Not Discounted"
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   3120
      TabIndex        =   107
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   600
      TabIndex        =   106
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   8940
      TabIndex        =   105
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   4200
      TabIndex        =   104
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label lblDis 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   6960
      TabIndex        =   103
      ToolTipText     =   "Tax And Freight Are Not Discounted"
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label lblDis 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   6960
      TabIndex        =   102
      ToolTipText     =   "Tax And Freight Are Not Discounted"
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   4200
      TabIndex        =   101
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   8940
      TabIndex        =   100
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   600
      TabIndex        =   99
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   3120
      TabIndex        =   98
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label lblDis 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   6960
      TabIndex        =   97
      ToolTipText     =   "Tax And Freight Are Not Discounted"
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   4200
      TabIndex        =   96
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   8940
      TabIndex        =   95
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   600
      TabIndex        =   94
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   3120
      TabIndex        =   93
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lbldsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3240
      TabIndex        =   92
      Top             =   320
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Checking Account"
      Height          =   285
      Index           =   16
      Left            =   240
      TabIndex        =   91
      Top             =   320
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   90
      Top             =   2760
      Width           =   345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   6600
      TabIndex        =   89
      Top             =   2760
      Width           =   345
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   4800
      Picture         =   "diaAPe03a.frx":0C90
      Top             =   6480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   4560
      Picture         =   "diaAPe03a.frx":1182
      Top             =   6480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   4320
      Picture         =   "diaAPe03a.frx":1674
      Top             =   6480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   5040
      Picture         =   "diaAPe03a.frx":1B66
      Top             =   6480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Memo"
      Height          =   285
      Index           =   18
      Left            =   240
      TabIndex        =   83
      Top             =   2280
      Width           =   1425
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   82
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Selected"
      Height          =   285
      Index           =   15
      Left            =   240
      TabIndex        =   81
      Top             =   6360
      Width           =   1395
   End
   Begin VB.Label lblPge 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   79
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label lblLst 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   78
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   255
      Index           =   14
      Left            =   5520
      TabIndex        =   77
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      Height          =   255
      Index           =   13
      Left            =   6480
      TabIndex        =   76
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount        "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   6960
      TabIndex        =   74
      Top             =   2760
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt Pay           "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   5400
      TabIndex        =   73
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt Due          "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   4200
      TabIndex        =   72
      Top             =   2760
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO No              "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   8940
      TabIndex        =   71
      Top             =   2760
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Number                            "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   600
      TabIndex        =   70
      Top             =   2760
      Width           =   2385
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date              "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   3120
      TabIndex        =   69
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   3120
      TabIndex        =   68
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   600
      TabIndex        =   67
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   8940
      TabIndex        =   66
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   4200
      TabIndex        =   65
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblDis 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   6960
      TabIndex        =   64
      ToolTipText     =   "Tax And Freight Are Not Discounted"
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   3120
      TabIndex        =   63
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   600
      TabIndex        =   62
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   8940
      TabIndex        =   61
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   4200
      TabIndex        =   60
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblDis 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   6960
      TabIndex        =   59
      ToolTipText     =   "Tax And Freight Are Not Discounted"
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   3120
      TabIndex        =   58
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   600
      TabIndex        =   57
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   8940
      TabIndex        =   56
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   4200
      TabIndex        =   55
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblDis 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   6960
      TabIndex        =   54
      ToolTipText     =   "Tax And Freight Are Not Discounted"
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   3120
      TabIndex        =   53
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   600
      TabIndex        =   52
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   8940
      TabIndex        =   51
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   4200
      TabIndex        =   50
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblDis 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   6960
      TabIndex        =   49
      ToolTipText     =   "Tax And Freight Are Not Discounted"
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   3120
      TabIndex        =   48
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   600
      TabIndex        =   47
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   8940
      TabIndex        =   46
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   4200
      TabIndex        =   45
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblDis 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   6960
      TabIndex        =   44
      ToolTipText     =   "Tax And Freight Are Not Discounted"
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblCnt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   7620
      TabIndex        =   43
      Top             =   720
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices Found"
      Height          =   285
      Index           =   5
      Left            =   6300
      TabIndex        =   42
      Top             =   720
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Amount"
      Height          =   285
      Index           =   4
      Left            =   5760
      TabIndex        =   41
      Top             =   1860
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number"
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   40
      Top             =   1860
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Date"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   39
      Top             =   1860
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   38
      Top             =   720
      Width           =   1425
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   37
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "diaAPe03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treates.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'************************************************************************************
' diaAPe03a - Cash Disbursements (External Check)
'
' Created: (cjs)
' Revions:
' 12/26/01 (nth) Added a ledger like look and feel
' 01/10/02 (nth) Changed GetVendorInvoices code to support AP discounts
' 09/10/02 (nth) Minor rework of UI and verbage in message boxs
' 12/11/02 (nth) Fixed postcheck debits and credits
' 07/07/03 (nth) Fixed the incrementing on DCTRAN and DCREF
' 08/25/03 (nth) Allow for 0 check amounts see incident # 18151
' 08/29/04 (nth) Added savelastcheck / getnextcheck
' 01/19/05 (nth) Added currencymask formatting to journal insert queries.
'
'************************************************************************************

'Dim RdoQry1 As rdoQuery
'Dim RdoQry2 As rdoQuery

Dim bOnLoad As Byte
Dim bGoodVendor As Byte
Dim bCancel As Byte

Dim iIndex As Integer
Dim iCurrPage As Integer
Dim iTotalItems As Integer

Dim lCheck As Double

Dim sApAcct As String
Dim sApDiscAcct As String
Dim sXcAcct As String
Dim sCcAcct As String
Dim sCgAcct As String
Dim sJournalID As String
Dim sMsg As String

' Grid Array
'
' 0 = Date
' 1 = Invoice #
Private Const INV_NO = 1
' 2 = PO
' 3 = Invoice Total
Private Const INV_TOTAL = 3
' 4 = Amount Paid
' 5 = Discount
Private Const INV_Discount = 5
' 6 = Freight
' 7 = Tax
' 8 = Discount Percentage
' 9 = Marked Paid? (0 or 1)
' 10 = Partial (0 or 1)
' 11 = Item Number
Private Const INV_DiscountCutoffDate = 12

'Dim sInvoices(500) As String    'not required
Dim vInvoice(500, 12) As Variant

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cboDueDate_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboDueDate_LostFocus()
   cboDueDate = CheckDate(cboDueDate)
   GetInvoiceCount
End Sub

'************************************************************************************

Private Sub cmbAct_Click()
   lbldsc = UpdateActDesc(cmbAct)
   sXcAcct = Compress(cmbAct)
   lCheck = GetNextCheck(sXcAcct)
   txtChk = lCheck
End Sub

Private Sub cmbAct_LostFocus()
   If Not bCancel Then
      lbldsc = UpdateActDesc(cmbAct)
      sXcAcct = Compress(cmbAct)
      lCheck = GetNextCheck(sXcAcct)
      txtChk = lCheck
   End If
End Sub

Private Sub cmbVnd_Click()
   txtAmt = "0.00"
   bGoodVendor = FindVendor(Me)
   If bGoodVendor Then GetInvoiceCount
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   If Len(cmbVnd) > 0 Then
      bGoodVendor = FindVendor(Me)
      If bGoodVendor Then GetInvoiceCount
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdCnl_Click()
   Dim bByte As Boolean
   Dim i As Integer
   Dim bResponse As Byte
   Dim sMsg As String
   
   For i = 1 To iTotalItems
      If vInvoice(i, 9) = vbChecked Then bByte = True
   Next
   If bByte Then
      sMsg = "You Have Items Selected. Are You" _
             & vbCrLf & "Sure That You Want To Cancel?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         bByte = False
      Else
         bByte = True
      End If
   End If
   If Not bByte Then
      
      cmbVnd.enabled = True
      cboCheckDate.enabled = True
      cboDueDate.enabled = True
      txtChk.enabled = True
      txtMemo.enabled = True
      txtAmt.enabled = True
      
      CloseBoxes
      txtDummy.enabled = False
      On Error Resume Next
      cmbVnd.SetFocus
   End If
End Sub

Private Sub cmdDn_Click()
   iCurrPage = iCurrPage + 1
   If iCurrPage > Val(lblLst) - 1 Then
      iCurrPage = Val(lblLst)
      cmdDn.enabled = False
      cmdDn.Picture = Dsdn
   Else
      cmdDn.enabled = True
      cmdDn.Picture = Endn
   End If
   If iCurrPage > 1 Then
      cmdUp.enabled = True
      cmdUp.Picture = Enup
   Else
      cmdUp.enabled = False
      cmdUp.Picture = Dsup
   End If
   lblPge = iCurrPage
   GetNextGroup
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Cash Disbursements"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdPst_Click()
   Dim bByte As Boolean
   Dim i As Integer
   
   For i = 1 To iTotalItems
      If vInvoice(i, 9) = vbChecked Then
         bByte = True
      End If
   Next
   If Not bByte Then
      MsgBox "You Have Nothing Selected To Pay.", vbInformation, Caption
   Else
      If CCur(txtAmt) <> CCur(lblTot) Then
         MsgBox "The Check Amount And Total Selected Must Agree.", vbInformation, Caption
         Exit Sub
      End If
      PostCheck
   End If
End Sub

Private Sub cmdSel_Click()
   Dim sNewJrnl As String
   Dim RdoChk As ADODB.Recordset
   
   If Len(Trim(txtChk)) = 0 Then
      MsgBox "You Must Have A Valid Check Number.", vbInformation, Caption
      txtChk.SetFocus
      Exit Sub
   End If
   
   If CCur(txtAmt) < 0 Then
      MsgBox "You Must Have A Valid Check Amount.", vbInformation, Caption
      txtAmt.SetFocus
      Exit Sub
   End If
   
   sSql = "SELECT CHKNUMBER FROM ChksTable WHERE " _
          & "CHKNUMBER = '" & Trim(txtChk) & "' AND " _
          & "CHKACCT = '" & Compress(cmbAct) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk, ES_FORWARD)
   If bSqlRows Then
      MsgBox "Check Number " & Trim(txtChk) & " For Account " & cmbAct & vbCrLf _
         & " Has Already Been Used.", vbInformation, Caption
      txtChk.SetFocus
      Set RdoChk = Nothing
      Exit Sub
   End If
   Set RdoChk = Nothing
   GetVendorInvoices
   
End Sub

Private Sub cmdUp_Click()
   iCurrPage = iCurrPage - 1
   If iCurrPage < 2 Then
      iCurrPage = 1
      cmdUp.enabled = False
      cmdUp.Picture = Dsup
   Else
      cmdUp.enabled = True
      cmdUp.Picture = Enup
   End If
   If iCurrPage < Val(lblLst) Then
      cmdDn.enabled = True
      cmdDn.Picture = Endn
   Else
      cmdDn.enabled = False
      cmdDn.Picture = Dsdn
   End If
   lblPge = iCurrPage
   GetNextGroup
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "Update Check Amount To Selected Amount?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then txtAmt = lblTot
End Sub



Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      'CurrentJournal "XC", ES_SYSDATE, sJournalID
      sJournalID = GetOpenJournal("XC", Format(ES_SYSDATE, "mm/dd/yy"))
      If sJournalID = "" Then
         MsgBox "There Is No Open External Checks Journal For The Period", _
            vbInformation, Caption
         Sleep 500
         MouseCursor 0
         Unload Me
         Exit Sub
      End If
      
      FillCombo
      cmbVnd = cUR.CurrentVendor
      bGoodVendor = FindVendor(Me)
      lCheck = GetNextCheck(cmbAct)
      txtChk = lCheck
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   
   'sSql = "SELECT DISTINCT PONUMBER,PONETDAYS," _
   '       & "PODISCOUNT FROM PohdTable WHERE " _
   '       & "PONUMBER= ? AND PORELEASE= ? "
   'Set RdoQry1 = RdoCon.CreateQuery("", sSql)
   'RdoQry1.MaxRows = 1
   
   'sSql = "SELECT VINO,VIVENDOR,VIDATE,VIFREIGHT,VITAX,VIPAY " _
   '       & "FROM VihdTable WHERE (VINO= ? AND VIVENDOR= ?)"
   'Set RdoQry2 = RdoCon.CreateQuery("", sSql)
   
   cboCheckDate = Format(ES_SYSDATE, "mm/dd/yy")
   cboDueDate = cboCheckDate
   CloseBoxes
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bGoodVendor Then
      cUR.CurrentVendor = cmbVnd
      SaveCurrentSelections
   End If
   'Set RdoQry1 = Nothing
   'Set RdoQry2 = Nothing
   FormUnload
   Set diaAPe03a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub lblDue_MouseDown(Index As Integer, Button As Integer, _
                             Shift As Integer, X As Single, Y As Single)
   'Dim cTotal As Currency
   'If Button = 2 Then
   '   cTotal = (Val(vInvoice(iIndex + Index, 4)) + Val(vInvoice(iIndex + Index, INV_Discount))) - Val(vInvoice(iIndex + Index, 6))
   
   'lblDet = "Freight and Tax: " & vInvoice(Index + iIndex, 6) _
   '    & vbCrLf & vbCrLf _
   '    & "All Items:         " _
   '    & Format(cTotal, "######0.00")
   
   'lblDet.Left = lblDue(Index).Left - lblDue(Index).Width
   'lblDet.Top = lblDue(Index).Top + 300
   'lblDet.Visible = True
   'lblDet.Refresh
   'Sleep 5000
   'lblDet.Visible = False
   'lblDet.Top = 0
   ' End If
End Sub

Private Sub optPad_Click(Index As Integer)
   If optPad(Index) = vbChecked Then
      vInvoice(Index + iIndex, 4) = (vInvoice(Index + iIndex, INV_TOTAL) _
               - vInvoice(Index + iIndex, INV_Discount))
      txtPad(Index) = Format(vInvoice(Index + iIndex, 4), CURRENCYMASK)
      txtPad(Index).enabled = True
      
   Else
      vInvoice(Index + iIndex, 4) = 0
      txtPad(Index) = Format(vInvoice(Index + iIndex, 4), CURRENCYMASK)
      vInvoice(Index + iIndex, 10) = vbUnchecked
      optPar(Index) = vbUnchecked
      txtPad(Index).enabled = False
      txtPad(Index) = ""
   End If
   vInvoice(Index + iIndex, 9) = optPad(Index)
   
   
   UpdateTotals
   
   
   
   'Dim cTotalItem As Currency
   '
   'If optPad(Index).Value = vbChecked Then
   '   If optPar(Index).Value = vbChecked Then
   '      ' don't calc discount
   '   Else
   '      ' calc discount
   '      If vInvoice(Index + iIndex, INV_TOTAL = Val(txtPad(Index)) Then
   '           ' If its the full amount then calc a discount
   '           CheckDiscount Index, Index + iIndex
   '           vInvoice(Index + iIndex, 4) = txtPad(Index)
   '       Else
   '           ' don't calc discount
   '           vInvoice(Index + iIndex, 4) = txtPad(Index)
   '       End If
   '   End If
   'End If
   '   'txtPad(Index) = Format(0, "0.00")
   '  vInvoice(Index + iIndex, 4) = txtPad(Index)
   
   
   'vInvoice(iIndex + Index, 9) = optPad(Index).Value
   'UpdateTotals
End Sub

Private Sub optPad_KeyPress(Index As Integer, KeyAscii As Integer)
   'KeyLock KeyAscii
   
End Sub


Private Sub optPad_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub optPad_LostFocus(Index As Integer)
   vInvoice(iIndex + Index, 9) = optPad(Index).Value
   
End Sub

Private Sub optPar_Click(Index As Integer)
   If optPar(Index).Value = vbChecked Then
      vInvoice(Index + iIndex, INV_Discount) = 0
      vInvoice(Index + iIndex, 10) = 1
      lblDis(Index) = "0.00"
      optPad(Index) = vbChecked
      lblDis(Index).ForeColor = vbBlack
   Else
      vInvoice(Index + iIndex, 10) = 0
   End If
   UpdateTotals
End Sub

Private Sub optPar_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
End Sub


Private Sub optPar_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
End Sub

Private Sub txtAmt_LostFocus()
   txtAmt = CheckLen(txtAmt, 10)
   txtAmt = Format(Abs(Val(txtAmt)), "######0.00")
End Sub

Private Sub txtChk_LostFocus()
   If Not bCancel Then
      txtChk = CheckLen(txtChk, 12)
      If Not ValidateCheck(txtChk, cmbAct) Then
         sMsg = "Check Already Exist For Account " & cmbAct
         MsgBox sMsg, vbInformation, Caption
         txtChk.SetFocus
      End If
   End If
End Sub

Private Sub cboCheckDate_DropDown()
   ShowCalendar Me
End Sub

Public Sub FillCombo()
   Dim RdoVed As ADODB.Recordset
   Dim rdoAct As ADODB.Recordset
   Dim b As Byte
   Dim lCheck As Double
   Dim sOptions As String
   cmbVnd.Clear
   MouseCursor 13
   On Error GoTo DiaErr1
   
   b = GetApAccounts
   If b = 3 Then
      MouseCursor 0
      MsgBox "One Or More Payables Accounts Are Not Active." & vbCr _
         & "Please Set All AP Accounts In The System." & vbCr _
         & "Company Setup, Administration Section.", _
         vbInformation, Caption
      Unload Me
      Exit Sub
   End If
   
   ' Vendors
   sSql = "SELECT DISTINCT VITVENDOR,VEREF,VENICKNAME " _
          & "FROM ViitTable,VndrTable WHERE VITVENDOR=VEREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVed)
   If bSqlRows Then
      With RdoVed
         cmbVnd = "" & Trim(!VENICKNAME)
         Do Until .EOF
            
            AddComboStr cmbVnd.hWnd, "" & Trim(!VENICKNAME)
            .MoveNext
         Loop
      End With
   End If
   Set RdoVed = Nothing
   If cmbVnd.ListCount > 0 Then
      bGoodVendor = FindVendor(Me)
      GetInvoiceCount
   End If
   txtAmt = "0.00"
   
   ' Accounts
   sSql = "SELECT GLACCTNO FROM GlacTable WHERE GLCASH = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct)
   If bSqlRows Then
      With rdoAct
         While Not .EOF
            AddComboStr cmbAct.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Wend
      End With
      cmbAct.ListIndex = 0
   Else
      cmbAct = sXcAcct
      cmbAct.enabled = False
   End If
   lbldsc = UpdateActDesc(cmbAct)
   Set rdoAct = Nothing
   Exit Sub
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub GetInvoiceCount()
   Dim RdoInv As ADODB.Recordset
   Dim iTotal As Integer
   Dim sVendor As String
   
   CloseBoxes
   'Erase sInvoices
   iTotal = 0
   sVendor = Compress(cmbVnd)
   txtDummy.enabled = False
   lblTot = ""
   On Error GoTo DiaErr1
   
'   sSql = "SELECT DISTINCT VITNO,VITVENDOR FROM ViitTable,VihdTable " _
'          & "WHERE ViitTable.VITNO = VihdTable.VINO AND ViitTable.VITVENDOR=VihdTable.VIVENDOR " _
'          & "AND VITVENDOR='" & sVendor & "' AND (NOT VIPIF=1)"
   
'   bSqlRows = clsAdoCon.GetDataSet(sSql,rdoInv)
'   If bSqlRows Then
'      With rdoInv
'         Do Until .EOF
'            iTotal = iTotal + 1
'            sInvoices(iTotal) = "" & Trim(!VITNO)
'            .MoveNext
'         Loop
'         .Cancel
'      End With
'      cmdSel.enabled = True
'   Else
'      cmdSel.enabled = False
'   End If
'   Set rdoInv = Nothing
'   cmdPst.enabled = False
'   cmdCnl.enabled = False
'   lblCnt = Format(iTotal, "##0")
'   Exit Sub
   
   'show invoices with discounts available on or after check date
   'also show invoices due on or before due date
   sSql = "select count(*)" & vbCrLf _
      & "from dbo.fnGetOpenAP('" & cboCheckDate & "', '" & Me.cboDueDate & "') fn" & vbCrLf _
      & "where VIVENDOR = '" & sVendor & "'"
   
   If clsADOCon.GetDataSet(sSql, RdoInv) Then
      iTotal = RdoInv.Fields(0)
   End If
   
   If iTotal > 0 Then
      cmdSel.enabled = True
   Else
      cmdSel.enabled = False
   End If
   
   Set RdoInv = Nothing
   cmdPst.enabled = False
   cmdCnl.enabled = False
   lblCnt = Format(iTotal, "##0")
   Exit Sub
   
DiaErr1:
   sProcName = "GetInvoiceCount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cboCheckDate_LostFocus()
   cboCheckDate = CheckDate(cboCheckDate)
   GetInvoiceCount
End Sub

Private Sub txtMemo_LostFocus()
   txtMemo = CheckLen(txtMemo, 40)
End Sub


Private Sub txtPad_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtPad_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtPad_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtPad_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtPad_LostFocus(Index As Integer)
   
   Dim cTotalItem As Currency
   Dim bResponse As Byte
   ' Format the currency
   txtPad(Index) = CheckLen(txtPad(Index), 10)
   If CheckCurrency(txtPad(Index), False) = "*" Then
    txtPad(Index).SetFocus
    Exit Sub
   End If
   
   
   If txtPad(Index) <> "" Then
      txtPad(Index) = Format(CCur(txtPad(Index)), CURRENCYMASK)
   End If
   ' If include button is checked record pay amt
   If optPad(Index).Value = vbChecked Then
      ' paid too much
      If CCur(txtPad(Index)) > CCur(lblDue(Index)) Then
         Beep
         'txtPad(Index) = lblDue(Index)
         'lblDis(Index) = "0.00"
         'optPad(Index).Value = vbUnchecked
         
         sMsg = "Due amount is less.Do you want to proceed?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbNo Then
           txtPad(Index) = lblDue(Index)
           lblDis(Index) = "0.00"
           optPad(Index).Value = vbUnchecked
         End If
         
         ' payment ok
      Else
         If CCur(txtPad(Index)) >= 0 And optPar(Index) = vbUnchecked Then
            cTotalItem = CCur(lblDue(Index)) - CCur(txtPad(Index))
            lblDis(Index) = Format(cTotalItem, CURRENCYMASK)
         ElseIf CCur(txtPad(Index)) >= 0 And optPar(Index) = vbUnchecked Then
            lblDis(Index) = "0.00" 'Format(0, CURRENCYMASK)
         End If
      End If
      vInvoice(Index + iIndex, 4) = txtPad(Index)
      'CheckDiscount Index, Index + iIndex
   Else
      lblDis(Index) = Format(0, CURRENCYMASK)
   End If
   
   UpdateTotals
End Sub

Public Sub CloseBoxes()
   Dim i As Integer
   For i = 1 To 9
      lblDte(i) = ""
      lblInv(i) = ""
      lblPon(i) = ""
      lblDue(i) = ""
      'lblDue(i).Enabled = False
      lblDis(i) = ""
      txtPad(i) = ""
      txtPad(i).enabled = False
      optPad(i).Caption = ""
      optPad(i).enabled = False
      optPad(i).Value = vbUnchecked
      optPar(i).Caption = ""
      optPar(i).enabled = False
      optPar(i).Value = vbUnchecked
   Next
   cmdCnl.enabled = False
   cmdPst.enabled = False
   cmdUpd.enabled = False
End Sub

'Public Sub GetVendorInvoices()
'   Dim rdoInv As ADODB.RecordSet
'   Dim RdoTol As ADODB.RecordSet
'   Dim rdoPo As ADODB.RecordSet
'   Dim rdoDis As ADODB.RecordSet
'   Dim rdoTemp As ADODB.RecordSet
'   Dim i As Integer
'   Dim cTotal As Currency
'   Dim cDiscount As Currency
'   Dim sVendor As String
'   Dim iDays As Integer
'   Dim iDiscFreight As Integer
'
'   MouseCursor 13
'
'   On Error GoTo DiaErr1
'
'   lblTot = "0.00"
'   iTotalItems = 0
'   Erase vInvoice
'   sVendor = Compress(cmbVnd)
'
'   ' Discounts on freight ?
'   sSql = "SELECT COAPDISC FROM ComnTable"
'   bSqlRows = clsAdoCon.GetDataSet(sSql,rdoTemp)
'   If Not IsNull(rdoTemp!COAPDISC) Then iDiscFreight = 1
'   Set rdoTemp = Nothing
'
'
'   sSql = "SELECT VINO,VIVENDOR,VIDATE,VIFREIGHT,VITAX,VIPAY,VENICKNAME," _
'          & "VIDUE,VIDUEDATE FROM VihdTable,VndrTable WHERE (VIVENDOR = VEREF) " _
'          & "AND (NOT VIPIF = 1) AND " _
'          & "(VIVENDOR = '" & sVendor & "')"
'
'   bSqlRows = clsAdoCon.GetDataSet(sSql,rdoInv)
'
'   With rdoInv
'      While Not .EOF
'         iTotalItems = iTotalItems + 1
'         sVendor = Trim(!VIVENDOR)
'
'         ' Calculate the invoice total
'         sSql = "SELECT ROUND(SUM((VITQTY * VITCOST)+VITADDERS), 2) FROM ViitTable WHERE (VITVENDOR = '" _
'                & sVendor & "') AND (VITNO = '" & Trim(!VINO) & "')"
'         bSqlRows = clsAdoCon.GetDataSet(sSql,RdoTol)
'         cTotal = RdoTol.Fields(0)
'         Set RdoTol = Nothing
'
'         ' Calculate discount
'         sSql = "SELECT DISTINCT VITPO,VITPORELEASE " _
'                & "FROM ViitTable WHERE (VITVENDOR = '" & sVendor _
'                & " ') AND (VITNO = '" & Trim(!VINO) & "' AND VITPO <> 0)"
'         bSqlRows = clsAdoCon.GetDataSet(sSql,rdoPo, ES_FORWARD)
'
'
'         If bSqlRows Then
'            vInvoice(iTotalItems, 2) = Format(rdoPo!VITPO, "000000")
'
'            If rdoPo!VITPORELEASE <> 0 Then
'               vInvoice(iTotalItems, 2) = vInvoice(iTotalItems, 2) & ("-" & Trim(rdoPo!VITPORELEASE))
'            End If
'
'            ' Invoice has a PO use PO terms
'            sSql = "SELECT PODDAYS,PODISCOUNT FROM PohdTable " _
'                   & "WHERE (PONUMBER = " & rdoPo!VITPO & ") AND " _
'                   & "(PORELEASE = " & rdoPo!VITPORELEASE & ")"
'            If GetDataSet(rdoDis) Then
'               iDays = rdoDis!PODDAYS
'               cDiscount = (rdoDis!PODISCOUNT / 100)
'               Set rdoDis = Nothing
'            End If
'
'         Else
'            vInvoice(iTotalItems, 2) = ""
'            ' No PO for this invoice get discount terms from vendor
'            sSql = "SELECT VEDDAYS,VEDISCOUNT FROM VndrTable " _
'                   & "WHERE VENICKNAME = '" & sVendor & "'"
'            If GetDataSet(rdoDis) Then
'               iDays = rdoDis!VEDDAYS
'               cDiscount = (rdoDis!VEDISCOUNT / 100)
'               Set rdoDis = Nothing
'            Else
'               iDays = 0
'               cDiscount = 0
'            End If
'         End If
'
'         Set rdoPo = Nothing
'
'         ' Assign values to invoice grid array
'         vInvoice(iTotalItems, 0) = Format(!VIDATE, "mm/dd/yy")
'         vInvoice(iTotalItems, INV_NO) = "" & Trim(!VINO)
'
'         vInvoice(iTotalItems, INV_TOTAL) = (cTotal + !VIFREIGHT + !VITAX) - !VIPAY
'         vInvoice(iTotalItems, 4) = 0#
'
'         ' Calculate discount dollars
'         If iDays > 0 Then
'            If CDate(DateAdd("d", (-1 * iDays), CDate(!VIDUEDATE))) >= CDate(Now) Then
'               If iDiscFreight Then
'                  vInvoice(iTotalItems, INV_Discount) = cTotal * cDiscount
'               Else
'                  vInvoice(iTotalItems, INV_Discount) = (cTotal + !VIFREIGHT) * cDiscount
'               End If
'            Else
'               vInvoice(iTotalItems, INV_Discount) = 0#
'            End If
'         Else
'            vInvoice(iTotalItems, INV_Discount) = 0#
'         End If
'
'         vInvoice(iTotalItems, 6) = !VIFREIGHT
'         vInvoice(iTotalItems, 7) = !VITAX
'         vInvoice(iTotalItems, 9) = 0
'         vInvoice(iTotalItems, 10) = 0
'         vInvoice(iTotalItems, 12) = Trim(!VENICKNAME)
'
'         .MoveNext
'      Wend
'   End With
'   Set rdoInv = Nothing
'
'   iIndex = 0
'   If iTotalItems > 0 Then
'      iCurrPage = 1
'      lblLst = CInt((((iTotalItems * 100) / 100) / 9) + 0.5)
'      For i = 1 To 9
'
'         If i > iTotalItems Then Exit For
'
'         lblDte(i) = Format(vInvoice(i, 0), "mm/dd/yy")
'         lblInv(i) = Trim(vInvoice(i, INV_NO))
'         lblPon(i) = Trim(vInvoice(i, 2))
'         lblDue(i) = Format(vInvoice(i, INV_TOTAL), CURRENCYMASK)
'         lblDis(i) = Format(vInvoice(i, INV_Discount), CURRENCYMASK)
'         'txtPad(i) = Format(vInvoice(i, 4), "#,###,##0.00")
'
'
'         'If Val(txtPad(i)) >= 0 Then
'         '    txtPad(i).Enabled = True
'         'Else
'         '    txtPad(i).Enabled = False
'         'End If
'
'         optPad(i).enabled = True
'         optPar(i).enabled = True
'
'         On Error Resume Next
'
'         txtDummy.enabled = True
'
'         cmdPst.enabled = True
'         cmdCnl.enabled = True
'         cmdSel.enabled = False
'
'         optPad(1).SetFocus
'      Next
'
'      If Val(lblLst) > 1 Then
'         cmdDn.enabled = True
'         cmdDn.Picture = Endn
'      End If
'      cboCheckDate.enabled = False
'      cmbVnd.enabled = False
'   End If
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "getvendorinv"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'End Sub
'

Public Sub GetVendorInvoices()
   Dim RdoInv As ADODB.Recordset
   'Dim RdoTol As ADODB.Recordset
   'Dim rdoPo As ADODB.Recordset
   'Dim rdoDis As ADODB.Recordset
   Dim rdoTemp As ADODB.Recordset
   Dim i As Integer
   Dim cTotal As Currency
   Dim cDiscount As Currency
   Dim sVendor As String
   Dim iDays As Integer
   Dim iDiscFreight As Integer
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   
   lblTot = "0.00"
   iTotalItems = 0
   Erase vInvoice
   sVendor = Compress(cmbVnd)
   
   ' Discounts on freight ?
   sSql = "SELECT COAPDISC FROM ComnTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTemp)
   If Not IsNull(rdoTemp!COAPDISC) Then iDiscFreight = 1
   Set rdoTemp = Nothing
   
   
'   sSql = "SELECT VINO,VIVENDOR,VIDATE,VIFREIGHT,VITAX,VIPAY,VENICKNAME," _
'          & "VIDUE,VIDUEDATE FROM VihdTable,VndrTable WHERE (VIVENDOR = VEREF) " _
'          & "AND (NOT VIPIF = 1) AND " _
'          & "(VIVENDOR = '" & sVendor & "')"
   
'   sSql = "select fn.*" & vbCrLf _
'      & "from dbo.fnGetOpenAP('" & GetServerDate & "', '" & cboCheckDate & "') fn" & vbCrLf _
'      & "where VIVENDOR = '" & sVendor & "'" & vbCrLf _
'      & "order by VIDATE"
   
   'show invoices with discounts available on or after check date
   'also show invoices due on or before due date
   sSql = "select fn.*" & vbCrLf _
      & "from dbo.fnGetOpenAP('" & cboCheckDate & "', '" & Me.cboDueDate & "') fn" & vbCrLf _
      & "where VIVENDOR = '" & sVendor & "'" & vbCrLf _
      & "order by VIDATE"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
   
   With RdoInv
      While Not .EOF
         iTotalItems = iTotalItems + 1
         'sVendor = Trim(!VIVENDOR)
         
'         ' Calculate the invoice total
'         sSql = "SELECT ROUND(SUM((VITQTY * VITCOST)+VITADDERS), 2) FROM ViitTable WHERE (VITVENDOR = '" _
'                & sVendor & "') AND (VITNO = '" & Trim(!VINO) & "')"
'         bSqlRows = clsAdoCon.GetDataSet(sSql,RdoTol)
'         cTotal = RdoTol.Fields(0)
'         Set RdoTol = Nothing
         cTotal = !CalcTotal
         
         ' Calculate discount
'         sSql = "SELECT DISTINCT VITPO,VITPORELEASE " _
'                & "FROM ViitTable WHERE (VITVENDOR = '" & sVendor _
'                & " ') AND (VITNO = '" & Trim(!VINO) & "' AND VITPO <> 0)"
'         bSqlRows = clsAdoCon.GetDataSet(sSql,rdoPo, ES_FORWARD)
'         If bSqlRows Then
'            vInvoice(iTotalItems, 2) = Format(rdoPo!VITPO, "000000")
'
'            If rdoPo!VITPORELEASE <> 0 Then
'               vInvoice(iTotalItems, 2) = vInvoice(iTotalItems, 2) & ("-" & Trim(rdoPo!VITPORELEASE))
'            End If

         vInvoice(iTotalItems, 2) = Format(!VITPO, "000000") _
            & "-" & Format(!VITPORELEASE, "##0")


'
'            ' Invoice has a PO use PO terms
'            sSql = "SELECT PODDAYS,PODISCOUNT FROM PohdTable " _
'                   & "WHERE (PONUMBER = " & rdoPo!VITPO & ") AND " _
'                   & "(PORELEASE = " & rdoPo!VITPORELEASE & ")"
'            If GetDataSet(rdoDis) Then
'               iDays = rdoDis!PODDAYS
'               cDiscount = (rdoDis!PODISCOUNT / 100)
'               Set rdoDis = Nothing
'            End If
'
'         Else
'            vInvoice(iTotalItems, 2) = ""
'            ' No PO for this invoice get discount terms from vendor
'            sSql = "SELECT VEDDAYS,VEDISCOUNT FROM VndrTable " _
'                   & "WHERE VENICKNAME = '" & sVendor & "'"
'            If GetDataSet(rdoDis) Then
'               iDays = rdoDis!VEDDAYS
'               cDiscount = (rdoDis!VEDISCOUNT / 100)
'               Set rdoDis = Nothing
'            Else
'               iDays = 0
'               cDiscount = 0
'            End If
'         End If
'
'         Set rdoPo = Nothing
         
         ' Assign values to invoice grid array
         vInvoice(iTotalItems, 0) = Format(!VIDATE, "mm/dd/yy")
         vInvoice(iTotalItems, INV_NO) = "" & Trim(!VINO)
         
         vInvoice(iTotalItems, INV_TOTAL) = cTotal - !VIPAY
         vInvoice(iTotalItems, 4) = 0#
         
'         ' Calculate discount dollars
'         If iDays > 0 Then
'            If CDate(DateAdd("d", (-1 * iDays), CDate(!VIDUEDATE))) >= CDate(Now) Then
'               If iDiscFreight Then
'                  vInvoice(iTotalItems, INV_Discount) = cTotal * cDiscount
'               Else
'                  vInvoice(iTotalItems, INV_Discount) = (cTotal + !VIFREIGHT) * cDiscount
'               End If
'            Else
'               vInvoice(iTotalItems, INV_Discount) = 0#
'            End If
'         Else
'            vInvoice(iTotalItems, INV_Discount) = 0#
'         End If
       
         vInvoice(iTotalItems, INV_Discount) = !DiscAvail

         vInvoice(iTotalItems, 6) = !VIFREIGHT
         vInvoice(iTotalItems, 7) = !VITAX
         vInvoice(iTotalItems, 9) = 0
         vInvoice(iTotalItems, 10) = 0
         vInvoice(iTotalItems, 12) = Trim(!VENICKNAME)
         
         If IsNull(!TakeDiscByDate) Then
            vInvoice(iTotalItems, INV_DiscountCutoffDate) = ""
         Else
            vInvoice(iTotalItems, INV_DiscountCutoffDate) = Format(!TakeDiscByDate, "mm/dd/yy")
         End If
         
         .MoveNext
      Wend
   End With
   Set RdoInv = Nothing
   
   iIndex = 0
   If iTotalItems > 0 Then
      iCurrPage = 1
'      lblLst = CInt((((iTotalItems * 100) / 100) / 9) + 0.5)
      lblLst = (iTotalItems - 1) \ 9 + 1 'calculate # of pages
      For i = 1 To 9
         
         If i > iTotalItems Then Exit For
         
         lblDte(i) = Format(vInvoice(i, 0), "mm/dd/yy")
         lblInv(i) = Trim(vInvoice(i, INV_NO))
         lblPon(i) = Trim(vInvoice(i, 2))
         lblDue(i) = Format(vInvoice(i, INV_TOTAL), CURRENCYMASK)
         lblDis(i) = Format(vInvoice(i, INV_Discount), CURRENCYMASK)
         Me.lblDiscDate(i) = vInvoice(i, INV_DiscountCutoffDate)
         'txtPad(i) = Format(vInvoice(i, 4), "#,###,##0.00")
         
         
         'If Val(txtPad(i)) >= 0 Then
         '    txtPad(i).Enabled = True
         'Else
         '    txtPad(i).Enabled = False
         'End If
         
         optPad(i).enabled = True
         optPar(i).enabled = True
         
         On Error Resume Next
         
         txtDummy.enabled = True
         
         cmdPst.enabled = True
         cmdCnl.enabled = True
         cmdSel.enabled = False
         
         optPad(1).SetFocus
      Next
      
      If Val(lblLst) > 1 Then
         cmdDn.enabled = True
         cmdDn.Picture = Endn
      End If
      cboCheckDate.enabled = False
      cboDueDate.enabled = False
      cmbVnd.enabled = False
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "GetVendorInvoices"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub



Public Sub GetNextGroup()
   Dim i As Integer
   iIndex = (Val(lblPge) - 1) * 9
   On Error GoTo DiaErr1
   For i = 1 To 9
      If i + iIndex > iTotalItems Then
         lblDte(i) = ""
         lblInv(i) = ""
         lblPon(i) = ""
         lblDue(i) = ""
         lblDue(i).enabled = False
         txtPad(i) = ""
         txtPad(i).enabled = False
         lblDis(i) = ""
         optPar(i).Caption = ""
         optPar(i).enabled = False
         optPar(i).Value = vbUnchecked
         optPad(i).Caption = ""
         optPad(i).enabled = False
         optPad(i).Value = vbUnchecked
         ' cmdCnl.Enabled = False
         ' cmdPst.Enabled = False
         Me.lblDiscDate(i) = ""
      Else
         lblDte(i) = vInvoice(i + iIndex, 0)
         lblInv(i) = vInvoice(i + iIndex, INV_NO)
         lblPon(i) = vInvoice(i + iIndex, 2)
         lblDue(i) = Format(vInvoice(i + iIndex, INV_TOTAL), CURRENCYMASK)
         lblDue(i).enabled = True
         txtPad(i) = Format(vInvoice(i + iIndex, 4), CURRENCYMASK)
         If Val(txtPad(i)) >= 0 Then
            txtPad(i).enabled = True
         Else
            txtPad(i).enabled = False
         End If
         lblDis(i) = Format(vInvoice(i + iIndex, INV_Discount), CURRENCYMASK)
         optPar(i).Caption = "__"
         optPar(i).Value = Val(vInvoice(i + iIndex, 10))
         optPar(i).enabled = True
         optPad(i).Caption = "__"
         optPad(i).Value = Val(vInvoice(i + iIndex, 9))
         optPad(i).enabled = True
         lblDiscDate(i) = vInvoice(i + iIndex, INV_DiscountCutoffDate)
      End If
   Next
   On Error Resume Next
   txtPad(1).SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "getnextgro"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub UpdateTotals()
   Dim i As Integer
   On Error Resume Next
   Dim cTotal As Currency
   For i = 1 To iTotalItems
      If vInvoice(i, 9) = vbChecked Then _
                  cTotal = cTotal + CCur(vInvoice(i, 4))
   Next
   If cTotal > 0 Then
      cmdUpd.enabled = True
   Else
      cmdUpd.enabled = False
   End If
   lblTot = Format(cTotal, CURRENCYMASK)
   
End Sub

Public Sub PostCheck()
   Dim RdoPst As ADODB.Recordset
   Dim bSuccess As Boolean
   Dim i As Integer
   Dim bResponse As Byte
   Dim iTrans As Integer
   Dim iRef As Integer
   Dim sCheckNo As String
   Dim cDiscount As Currency
   Dim cTotal As Currency
   Dim cDiscRate As Currency
   Dim cPaid As Currency
   Dim sCdCashAct As String
   Dim sNewAcct As String
   Dim sMsg As String
   Dim sVendor As String
   'Dim sToday     As String
   Dim sCredit As String
   Dim sDebit As String
   
   ' Check if journal exist
   sJournalID = GetOpenJournal("XC", Format(cboCheckDate, "mm/dd/yy"))
   If sJournalID = "" Then
      sMsg = "There Is No Open External Check Cash Disbursements" _
             & vbCrLf & "Journal For " & cboCheckDate
      MsgBox sMsg, vbInformation, Caption
      MouseCursor 0
      Exit Sub
   End If
   
   sMsg = "Are You Ready To Post This Payment?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   
   'sToday = Format(ES_SYSDATE, "mm/dd/yy")
   
   If bResponse = vbYes Then
      
      MouseCursor 13
      
      ' XC cash account
      sCdCashAct = sXcAcct
      
      sVendor = Compress(cmbVnd)
      
      ' Get the journal trans (should be a journal)
      iTrans = GetNextTransaction(sJournalID)
      
      cmdPst.enabled = False
      cmdCnl.enabled = False
      cmdUpd.enabled = False
      
      'On Error Resume Next   we need to diagnose errors as they happen
      Err = 0
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      sCheckNo = Trim(txtChk)
      
      For i = 1 To iTotalItems
         If vInvoice(i, 9) = vbChecked Then
            ' Reverse if credit memo
            If CCur(vInvoice(i, INV_TOTAL)) < 0 Then
               sDebit = "DCCREDIT"
               sCredit = "DCDEBIT"
            Else
               sDebit = "DCDEBIT"
               sCredit = "DCCREDIT"
            End If
            If IsEmpty((vInvoice(i, 10))) Or Val(vInvoice(i, 10)) = 0 Then
               ' No Partial Payment -- Might have discount though
               
               ' Using the discount rate in the item table
               If CCur(vInvoice(i, 4)) > 0 Then
                  cDiscRate = CCur(vInvoice(i, INV_Discount)) / CCur(vInvoice(i, 4))
               End If
               
               sSql = "UPDATE VihdTable SET VIPIF=1,VIPAY=VIPAY+" _
                  & Abs(Val(vInvoice(i, 4))) & ",VIDISCOUNT=VIDISCOUNT+" _
                  & CCur(vInvoice(i, INV_Discount)) & ",VICHECKNO='" & Trim(txtChk) & "'," & vbCrLf _
                  & "VICHKACCT = '" & Compress(sXcAcct) & "' WHERE VINO='" _
                  & vInvoice(i, INV_NO) & "' AND VIVENDOR='" _
                  & sVendor & "' "
               clsADOCon.ExecuteSql sSql
               
               sSql = "UPDATE ViitTable SET VITCHECKNO='" & Trim(txtChk) _
                      & "',VITCHECKDT='" & cboCheckDate & "',VITDISCOUNT=" _
                      & cDiscRate & ",VITPAID=1,VITTOTPAID=(VITQTY*VITCOST) " _
                      & "WHERE (VITNO='" & vInvoice(i, INV_NO) & "' AND VITVENDOR='" _
                      & sVendor & "') "
               clsADOCon.ExecuteSql sSql
               
               ' We are PIF and taking a discount.
               ' So make the appropriate journal entries.
               ' Debit Ap, Credit XC Cash , Credit Discount
               cTotal = Abs(CCur(vInvoice(i, INV_TOTAL)))
               cPaid = Abs(CCur(vInvoice(i, 4)))
               ' This discount assumes you know what your discount is
               cDiscount = CCur(vInvoice(i, INV_TOTAL)) - CCur(vInvoice(i, 4))
               
               ' 1 - Debit Accounts Payable
               'Format(cTotal, CURRENCYMASK) & ",'" can't handle commas
               iRef = iRef + 1
               sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," & sDebit & ",DCACCTNO," _
                  & "DCDATE,DCCHECKNO,DCCHKACCT, DCPONUMBER,DCVENDOR,DCVENDORINV,DCTYPE) " & vbCrLf _
                  & "VALUES('" _
                  & sJournalID & "'," _
                  & iTrans & "," _
                  & iRef & "," _
                  & cTotal & ",'" _
                  & Compress(sApAcct) & "','" _
                  & cboCheckDate & "','" _
                  & CStr(sCheckNo) & "','" _
                  & Compress(sXcAcct) & "'," _
                  & Val(vInvoice(i, 2)) & ",'" _
                  & sVendor & "','" _
                  & vInvoice(i, INV_NO) & "'," _
                  & DCTYPE_AP_ChkAPAcct & ")"
               clsADOCon.ExecuteSql sSql
               
               ' 2 - Credit XC Cash
               '  & Format(cPaid, CURRENCYMASK) & ",'"  doesn't work because of commas
               iRef = iRef + 1
               sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," & sCredit & ",DCACCTNO," _
                  & "DCDATE,DCCHECKNO,DCCHKACCT,DCPONUMBER,DCVENDOR,DCVENDORINV,DCTYPE) " & vbCrLf _
                  & "VALUES('" _
                  & sJournalID & "'," _
                  & iTrans & "," _
                  & iRef & "," _
                  & cPaid & ",'" _
                  & Compress(sXcAcct) & "','" _
                  & cboCheckDate & "','" _
                  & CStr(sCheckNo) & "','" _
                  & Compress(sXcAcct) & "'," _
                  & Val(vInvoice(i, 2)) & ",'" _
                  & sVendor & "','" _
                  & vInvoice(i, INV_NO) & "'," _
                  & DCTYPE_AP_ChkChkAcct & ")"
               clsADOCon.ExecuteSql sSql
               
               If cDiscount > 0 Then
                  ' 3 - Credit Discount
                  iRef = iRef + 1
                  sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," & sCredit & ",DCACCTNO," _
                     & "DCDATE,DCCHECKNO,DCCHKACCT,DCPONUMBER,DCVENDOR,DCVENDORINV,DCTYPE) " & vbCrLf _
                     & "VALUES('" _
                     & sJournalID & "'," _
                     & iTrans & "," _
                     & iRef & "," _
                     & cDiscount & ",'" _
                     & Compress(sApDiscAcct) & "','" _
                     & cboCheckDate & "','" _
                     & CStr(sCheckNo) & "','" _
                     & Compress(sXcAcct) & "'," _
                     & Val(vInvoice(i, 2)) & ",'" _
                     & sVendor & "','" _
                     & vInvoice(i, INV_NO) & "'," _
                     & DCTYPE_AP_ChkDiscAcct & ")"
                  clsADOCon.ExecuteSql sSql
               End If
            Else
               ' Partial Payment
               sSql = "UPDATE VihdTable SET VIPIF=0,VIPAY=VIPAY+" _
                      & Val(vInvoice(i, 4)) & ",VICHECKNO='" & Trim(txtChk) _
                      & "', VICHKACCT = '" & Compress(sXcAcct) & "' WHERE VINO='" & vInvoice(i, INV_NO) & "' AND " _
                      & "VIVENDOR='" & sVendor & "' "
               clsADOCon.ExecuteSql sSql
               ' Insert journal entries
               cPaid = vInvoice(i, 4)
               
               ' 1 - DEBIT Accounts Payable
               iRef = iRef + 1
               sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," & sDebit & ",DCACCTNO," _
                  & "DCDATE,DCCHECKNO,DCCHKACCT,DCPONUMBER,DCVENDOR,DCVENDORINV,DCTYPE) " & vbCrLf _
                  & "VALUES('" _
                  & sJournalID & "'," _
                  & iTrans & "," _
                  & iRef & "," _
                  & Format(cPaid, CURRENCYMASK) & ",'" _
                  & Compress(sApAcct) & "','" _
                  & cboCheckDate & "','" _
                  & CStr(sCheckNo) & "','" _
                  & Compress(sXcAcct) & "'," _
                  & Val(vInvoice(i, 2)) & ",'" _
                  & sVendor & "','" _
                  & vInvoice(i, INV_NO) & "'," _
                  & DCTYPE_AP_ChkAPAcct & ")"
               clsADOCon.ExecuteSql sSql
               
               ' 2 - Credit XC Cash
               ' Checking account
               iRef = iRef + 1
               sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," & sCredit & ",DCACCTNO," _
                  & "DCDATE,DCCHECKNO,DCCHKACCT,DCPONUMBER,DCVENDOR,DCVENDORINV,DCTYPE) " & vbCrLf _
                  & "VALUES('" _
                  & sJournalID & "'," _
                  & iTrans & "," _
                  & iRef & "," _
                  & cPaid & ",'" _
                  & Compress(sXcAcct) & "','" _
                  & cboCheckDate & "','" _
                  & CStr(sCheckNo) & "','" _
                  & Compress(sXcAcct) & "'," _
                  & Val(vInvoice(i, 2)) & ",'" _
                  & sVendor & "','" _
                  & vInvoice(i, INV_NO) & "'," _
                  & DCTYPE_AP_ChkChkAcct & ")"
               clsADOCon.ExecuteSql sSql
            End If
            iTrans = iTrans + 1
            iRef = 0
         End If
      Next
      
      ' Record the check.....
      sSql = "INSERT INTO ChksTable (CHKNUMBER,CHKVENDOR,CHKAMOUNT," _
             & "CHKPOSTDATE,CHKBY,CHKMEMO,CHKVOID,CHKPRINTED,CHKTYPE,CHKACCT)" & vbCrLf _
             & "VALUES('" _
             & Val(txtChk) & "','" _
             & Trim(sVendor) & "'," _
             & CCur(txtAmt) & ",'" _
             & Format(cboCheckDate, "mm/dd/yyyy") & "','" _
             & Replace(Secure.UserInitials, Chr(0), " ") & "','" _
             & Trim(txtMemo) & "'," _
             & "0,0,1,'" & sXcAcct & "')"
Debug.Print sSql
      clsADOCon.ExecuteSql sSql
      
      ' Check for errors, if none finialize the transaction
      If clsADOCon.ADOErrNum = 0 Then
         bSuccess = True
         clsADOCon.CommitTrans
         'sSql = "UPDATE Preferences SET PreLastCheck=" & Val(txtChk) & " " _
         '    & "WHERE PreRecord=1"
         'clsAdoCon.ExecuteSQL sSql
         SaveLastCheck txtChk, sXcAcct
      Else
         bSuccess = False
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
      End If
      MouseCursor 0
      If bSuccess Then
         sMsg = "Check " & Trim(txtChk) & " Successfully Posted."
         SysMsg sMsg, True
         lCheck = GetNextCheck(sXcAcct)
         txtChk = lCheck
      Else
         sMsg = "Unable To Post Check " & Trim(txtChk) & vbCrLf _
                & "Transaction Canceled."
         MsgBox sMsg, vbExclamation, Caption
      End If
      FillCombo
      cmbVnd.enabled = True
      cboCheckDate.enabled = True
      cboDueDate.enabled = True
      txtChk.enabled = True
      txtMemo.enabled = True
      txtAmt = "0.00"
      txtAmt.enabled = True
      On Error Resume Next
      cmbVnd.SetFocus
   Else
      CancelTrans
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "postcheck"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Function GetApAccounts() As Byte
   Dim RdoAps As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   
   ' Flash with 1 = no accounts nec, 2 = OK, 3 = Not enough accounts
   On Error GoTo DiaErr1
   sProcName = "getapacco"
   sSql = "SELECT COGLVERIFY,COAPACCT,COAPDISCACCT," _
          & "COXCCASHACCT " _
          & "FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAps, ES_FORWARD)
   If bSqlRows Then
      With RdoAps
         For i = 1 To 3
            If Not IsNull(.Fields(i)) Then
               If Trim(.Fields(i)) = "" Then b = 1
            Else
               b = 1
            End If
         Next
         sApAcct = "" & Trim(!COAPACCT)
         sApDiscAcct = "" & Trim(!COAPDISCACCT)
         sXcAcct = "" & Trim(!COXCCASHACCT)
         If b = 1 Then GetApAccounts = 3 Else GetApAccounts = 2
      End With
   Else
      GetApAccounts = 1 ' Fail
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getapaccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Public Sub CheckDiscount(CurBox As Integer, CurrentRecord As Integer)
   Dim cDiscount As Currency
   
   On Error GoTo DiaErr1
   If vInvoice(CurrentRecord, INV_TOTAL) > 0 Then
      cDiscount = 1 - (Val(vInvoice(CurrentRecord, 4)) / Val(vInvoice(CurrentRecord, INV_TOTAL)))
      If cDiscount > 0.1 Then
         If optPar(CurBox).Value = vbUnchecked Then
            lblDis(CurBox).ForeColor = ES_RED
         Else
            lblDis(CurBox).ForeColor = vbBlack
            lblDis(CurBox) = "0.00"
            vInvoice(CurrentRecord, INV_Discount) = "0.00"
         End If
      Else
         lblDis(CurBox).ForeColor = vbBlack
      End If
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "checkdiscount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

