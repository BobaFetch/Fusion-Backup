VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaARe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cash Receipts"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkShortPayment 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3060
      TabIndex        =   150
      Top             =   2580
      Width           =   255
   End
   Begin VB.CommandButton cmdDockLok 
      Enabled         =   0   'False
      Height          =   495
      Left            =   8280
      Picture         =   "diaARe04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   149
      ToolTipText     =   "Push Info to Document Imaging System "
      Top             =   2640
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   3960
      TabIndex        =   145
      Top             =   2520
      Width           =   3495
      Begin VB.OptionButton optOrderBy 
         Caption         =   "Invoice #"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   147
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optOrderBy 
         Caption         =   "Invoice Date"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   146
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label z1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sort By "
         Height          =   285
         Index           =   25
         Left            =   0
         TabIndex        =   148
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1800
      TabIndex        =   143
      Top             =   2880
      Width           =   255
   End
   Begin VB.ComboBox cmbPageNo 
      Height          =   315
      Left            =   6000
      TabIndex        =   141
      Top             =   7440
      Width           =   735
   End
   Begin VB.CommandButton cmdLastPage 
      Height          =   495
      Left            =   8400
      Picture         =   "diaARe04a.frx":05DA
      Style           =   1  'Graphical
      TabIndex        =   140
      ToolTipText     =   "Last Page"
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton cmdNextPage 
      Height          =   495
      Left            =   7920
      Picture         =   "diaARe04a.frx":0B52
      Style           =   1  'Graphical
      TabIndex        =   139
      ToolTipText     =   "Next Page"
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton cmdPrevPage 
      Height          =   495
      Left            =   7440
      Picture         =   "diaARe04a.frx":10C5
      Style           =   1  'Graphical
      TabIndex        =   138
      ToolTipText     =   "Previous Page"
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton cmdFirstPage 
      Height          =   495
      Left            =   6960
      Picture         =   "diaARe04a.frx":1637
      Style           =   1  'Graphical
      TabIndex        =   137
      ToolTipText     =   "First Page"
      Top             =   7320
      Width           =   495
   End
   Begin VB.CheckBox chkPIF 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   7320
      MaskColor       =   &H8000000F&
      TabIndex        =   16
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox chkPIF 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   7320
      TabIndex        =   20
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox chkPIF 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   7320
      TabIndex        =   24
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox chkPIF 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   7320
      TabIndex        =   28
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox chkPIF 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   7320
      TabIndex        =   32
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   5040
      Width           =   255
   End
   Begin VB.CheckBox chkPIF 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   6
      Left            =   7320
      TabIndex        =   135
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   5400
      Width           =   255
   End
   Begin VB.CheckBox chkPIF 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   7
      Left            =   7320
      TabIndex        =   134
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   5760
      Width           =   255
   End
   Begin VB.CheckBox chkPIF 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   8
      Left            =   7320
      TabIndex        =   133
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   6120
      Width           =   255
   End
   Begin VB.CheckBox chkPIF 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   9
      Left            =   7320
      TabIndex        =   45
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   6480
      Width           =   255
   End
   Begin VB.CommandButton cmdVew 
      Height          =   320
      Left            =   2760
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Other Payers"
      Top             =   720
      Width           =   350
   End
   Begin VB.TextBox txtAdvAmt 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Tag             =   "1"
      ToolTipText     =   "Amount Not Applied Directly To An Invoice"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox cmbAdvAct 
      Height          =   315
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   12
      Tag             =   "3"
      ToolTipText     =   "Account To Apply Advanced Payment To"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CheckBox chkAdv 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   8040
      TabIndex        =   13
      ToolTipText     =   "Create An Advanced Payment Invoice "
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtFee 
      Height          =   285
      Left            =   6240
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "Fee Associated With Wire Transfer"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox cmbapl 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   9
      Left            =   7620
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Tag             =   "8"
      ToolTipText     =   "Mark Invoice As Short Paid Or Apply Difference To Account "
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox txtPaid 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H80000012&
      Height          =   315
      Index           =   9
      Left            =   6140
      TabIndex        =   44
      ToolTipText     =   "Amount To Apply Againts This Invoice"
      Top             =   6480
      Width           =   1100
   End
   Begin VB.CheckBox chkPaid 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   43
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   6480
      Width           =   255
   End
   Begin VB.ComboBox cmbapl 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   8
      Left            =   7620
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Tag             =   "8"
      ToolTipText     =   "Mark Invoice As Short Paid Or Apply Difference To Account "
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox txtPaid 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H80000012&
      Height          =   315
      Index           =   8
      Left            =   6140
      TabIndex        =   41
      ToolTipText     =   "Amount To Apply Againts This Invoice"
      Top             =   6120
      Width           =   1100
   End
   Begin VB.CheckBox chkPaid 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   40
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   6120
      Width           =   255
   End
   Begin VB.CheckBox chkPaid 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   37
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   5760
      Width           =   255
   End
   Begin VB.TextBox txtPaid 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H80000012&
      Height          =   315
      Index           =   7
      Left            =   6140
      TabIndex        =   38
      ToolTipText     =   "Amount To Apply Againts This Invoice"
      Top             =   5760
      Width           =   1100
   End
   Begin VB.ComboBox cmbapl 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   7
      Left            =   7620
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Tag             =   "8"
      ToolTipText     =   "Mark Invoice As Short Paid Or Apply Difference To Account "
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CheckBox chkPaid 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   34
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   5400
      Width           =   255
   End
   Begin VB.TextBox txtPaid 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H80000012&
      Height          =   315
      Index           =   6
      Left            =   6140
      TabIndex        =   35
      ToolTipText     =   "Amount To Apply Againts This Invoice"
      Top             =   5400
      Width           =   1100
   End
   Begin VB.ComboBox cmbapl 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   6
      Left            =   7620
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Tag             =   "8"
      ToolTipText     =   "Mark Invoice As Short Paid Or Apply Difference To Account "
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txtChk 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      ToolTipText     =   "Check Number Or Wire Reference (12 Max)"
      Top             =   1440
      Width           =   1440
   End
   Begin VB.ComboBox txtCte 
      Height          =   315
      Left            =   3960
      TabIndex        =   5
      Tag             =   "4"
      ToolTipText     =   "Check / Wire Date"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.OptionButton optCsh 
      Caption         =   "C&ash"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Tag             =   "6"
      Top             =   1800
      Width           =   855
   End
   Begin VB.OptionButton optWir 
      Caption         =   "&Wire"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Tag             =   "7"
      Top             =   1800
      Width           =   735
   End
   Begin VB.OptionButton optChk 
      Caption         =   "&Check"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Cash Account"
      Top             =   240
      Width           =   1555
   End
   Begin VB.TextBox txtAmt 
      Height          =   285
      Left            =   6240
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Amount Recieved"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox cmbapl 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   5
      Left            =   7620
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Tag             =   "8"
      ToolTipText     =   "Mark Invoice As Short Paid Or Apply Difference To Account "
      Top             =   5040
      Width           =   1335
   End
   Begin VB.ComboBox cmbapl 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   4
      Left            =   7620
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Tag             =   "8"
      ToolTipText     =   "Mark Invoice As Short Paid Or Apply Difference To Account "
      Top             =   4680
      Width           =   1335
   End
   Begin VB.ComboBox cmbapl 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   3
      Left            =   7620
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Tag             =   "8"
      ToolTipText     =   "Mark Invoice As Short Paid Or Apply Difference To Account "
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ComboBox cmbapl 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   2
      Left            =   7620
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Tag             =   "8"
      ToolTipText     =   "Mark Invoice As Short Paid Or Apply Difference To Account "
      Top             =   3960
      Width           =   1335
   End
   Begin VB.ComboBox cmbapl 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   1
      Left            =   7620
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Tag             =   "8"
      ToolTipText     =   "Mark Invoice As Short Paid Or Apply Difference To Account "
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   6120
      TabIndex        =   79
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Height          =   315
      Left            =   8100
      TabIndex        =   48
      ToolTipText     =   "Select Invoices For This Vendor"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdCnl 
      Caption         =   "&Reselect"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8100
      TabIndex        =   50
      ToolTipText     =   "Cancel Operation"
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdPst 
      Caption         =   "&Post"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8100
      TabIndex        =   49
      ToolTipText     =   "Post This Cash Receipt"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtPaid 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H80000012&
      Height          =   315
      Index           =   5
      Left            =   6140
      TabIndex        =   31
      ToolTipText     =   "Amount To Apply Againts This Invoice"
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CheckBox chkPaid 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   30
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   5040
      Width           =   255
   End
   Begin VB.TextBox txtPaid 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H80000012&
      Height          =   315
      Index           =   4
      Left            =   6140
      TabIndex        =   27
      ToolTipText     =   "Amount To Apply Againts This Invoice"
      Top             =   4680
      Width           =   1100
   End
   Begin VB.CheckBox chkPaid 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   26
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   4680
      Width           =   255
   End
   Begin VB.TextBox txtPaid 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H80000012&
      Height          =   315
      Index           =   3
      Left            =   6140
      TabIndex        =   23
      ToolTipText     =   "Amount To Apply Againts This Invoice"
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CheckBox chkPaid 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   4320
      Width           =   255
   End
   Begin VB.TextBox txtPaid 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H80000012&
      Height          =   315
      Index           =   2
      Left            =   6140
      TabIndex        =   19
      ToolTipText     =   "Amount To Apply Againts This Invoice"
      Top             =   3960
      Width           =   1100
   End
   Begin VB.CheckBox chkPaid 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox txtPaid 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H80000012&
      Height          =   315
      Index           =   1
      Left            =   6140
      TabIndex        =   15
      ToolTipText     =   "Amount To Apply Againts This Invoice"
      Top             =   3600
      Width           =   1100
   End
   Begin VB.CheckBox chkPaid 
      Caption         =   "__"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   120
      MaskColor       =   &H8000000F&
      TabIndex        =   14
      ToolTipText     =   "Select Invoice For Payment"
      Top             =   3600
      Width           =   255
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   6240
      TabIndex        =   3
      Tag             =   "4"
      ToolTipText     =   "Date Enter (Today)"
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Customer Nicknames"
      Top             =   720
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   8100
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   0
      Width           =   855
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   54
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaARe04a.frx":1BB7
      PictureDn       =   "diaARe04a.frx":1CFD
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8025
      FormDesignWidth =   9255
   End
   Begin Threed.SSCommand cmdDn 
      Height          =   375
      Left            =   1680
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   "Next Page (Page Down)"
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaARe04a.frx":1E43
   End
   Begin Threed.SSCommand cmdUp 
      Height          =   375
      Left            =   1320
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Last Page (Page Up)"
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaARe04a.frx":2345
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   30
      Index           =   1
      Left            =   120
      TabIndex        =   99
      Top             =   3240
      Width           =   8775
      _Version        =   65536
      _ExtentX        =   15478
      _ExtentY        =   53
      _StockProps     =   14
      Caption         =   "SSFrame1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Create credit memo for  short payment"
      Height          =   285
      Index           =   26
      Left            =   120
      TabIndex        =   151
      Top             =   2580
      Width           =   2745
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay All on This Page"
      Height          =   285
      Index           =   24
      Left            =   120
      TabIndex        =   144
      Top             =   2880
      Width           =   1665
   End
   Begin VB.Label Label1 
      Caption         =   "Go to Page"
      Height          =   255
      Left            =   5040
      TabIndex        =   142
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PIF"
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
      Index           =   23
      Left            =   7320
      TabIndex        =   136
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label lblUnappliedTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Available"
      Height          =   285
      Left            =   2580
      TabIndex        =   132
      Top             =   6975
      Width           =   825
   End
   Begin VB.Label lblUnapplied 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3420
      TabIndex        =   131
      ToolTipText     =   "Total Amount Applied To Selected Invoices"
      Top             =   6975
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   130
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label lblOwe 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4320
      TabIndex        =   129
      ToolTipText     =   "Total Amount Due For Customer"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Adv. Payment"
      Height          =   285
      Index           =   20
      Left            =   6840
      TabIndex        =   128
      Top             =   2160
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unapplied Amount"
      Height          =   465
      Index           =   8
      Left            =   120
      TabIndex        =   127
      Top             =   2100
      Width           =   945
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   3960
      TabIndex        =   126
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   480
      TabIndex        =   51
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   480
      TabIndex        =   125
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   480
      TabIndex        =   124
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   480
      TabIndex        =   123
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   480
      TabIndex        =   122
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Date            "
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
      Left            =   480
      TabIndex        =   121
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   480
      TabIndex        =   120
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   480
      TabIndex        =   119
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   480
      TabIndex        =   118
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   480
      TabIndex        =   117
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Fee"
      Height          =   285
      Index           =   21
      Left            =   4680
      TabIndex        =   116
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   4995
      TabIndex        =   115
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lblPs 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   2730
      TabIndex        =   114
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   1605
      TabIndex        =   113
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lblInvAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   3855
      TabIndex        =   112
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   4995
      TabIndex        =   111
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblPs 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   2730
      TabIndex        =   110
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   1605
      TabIndex        =   109
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblInvAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   3855
      TabIndex        =   108
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblInvAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   3855
      TabIndex        =   107
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   1605
      TabIndex        =   106
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label lblPs 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   2730
      TabIndex        =   105
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   4995
      TabIndex        =   104
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label lblInvAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   3855
      TabIndex        =   103
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   1605
      TabIndex        =   102
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblPs 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   2730
      TabIndex        =   101
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   4995
      TabIndex        =   100
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Id/Check #"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   98
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   285
      Index           =   17
      Left            =   3120
      TabIndex        =   97
      Top             =   1440
      Width           =   585
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   96
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      Height          =   285
      Index           =   16
      Left            =   360
      TabIndex        =   95
      Top             =   240
      Width           =   705
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   285
      Index           =   4
      Left            =   5160
      TabIndex        =   94
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   4995
      TabIndex        =   93
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   4995
      TabIndex        =   92
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   4995
      TabIndex        =   91
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   4995
      TabIndex        =   90
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   4995
      TabIndex        =   89
      ToolTipText     =   "Amount Left To Pay On Invoice"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packslip #       "
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
      Index           =   19
      Left            =   2760
      TabIndex        =   88
      Top             =   3360
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Difference to       "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   18
      Left            =   7620
      TabIndex        =   87
      ToolTipText     =   "Mark Invoice As Short Paid Or Apply Difference To Account "
      Top             =   3360
      Width           =   1305
   End
   Begin VB.Label lblPs 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   2730
      TabIndex        =   86
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label lblPs 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   2730
      TabIndex        =   85
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblPs 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   2730
      TabIndex        =   84
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblPs 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   2730
      TabIndex        =   83
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblPs 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   2730
      TabIndex        =   82
      ToolTipText     =   "CM-Credit Memo,DM-Debit Memo,SO-Sales Order,PS-Pack Slip"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1260
      TabIndex        =   81
      ToolTipText     =   "Total Amount Applied To Selected Invoices"
      Top             =   6975
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Selected"
      Height          =   285
      Index           =   15
      Left            =   120
      TabIndex        =   80
      Top             =   6975
      Width           =   1125
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   2520
      Picture         =   "diaARe04a.frx":2847
      Top             =   7440
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   2880
      Picture         =   "diaARe04a.frx":2D39
      Top             =   7440
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   3240
      Picture         =   "diaARe04a.frx":322B
      Top             =   7440
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   2160
      Picture         =   "diaARe04a.frx":371D
      Top             =   7440
      Visible         =   0   'False
      Width           =   285
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
      Left            =   7560
      TabIndex        =   78
      Top             =   6960
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
      Left            =   8520
      TabIndex        =   77
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   255
      Index           =   14
      Left            =   6720
      TabIndex        =   76
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      Height          =   255
      Index           =   13
      Left            =   8040
      TabIndex        =   75
      Top             =   6960
      Width           =   375
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
      Left            =   120
      TabIndex        =   74
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Amt.   "
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
      Left            =   6120
      TabIndex        =   73
      ToolTipText     =   "Amount To Apply Againts This Invoice"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Due           "
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
      Left            =   4995
      TabIndex        =   72
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inv. Amount       "
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
      Left            =   3840
      TabIndex        =   71
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice #              "
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
      Left            =   1605
      TabIndex        =   70
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   1605
      TabIndex        =   69
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label lblInvAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   3855
      TabIndex        =   68
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   1605
      TabIndex        =   67
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblInvAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   3855
      TabIndex        =   66
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   1605
      TabIndex        =   65
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblInvAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   3855
      TabIndex        =   64
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   1605
      TabIndex        =   63
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblInvAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   3855
      TabIndex        =   62
      ToolTipText     =   "Includes Tax And Freight"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   1605
      TabIndex        =   61
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblInvAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   3855
      TabIndex        =   60
      ToolTipText     =   "Invoice Total Includes Tax And Freight"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblCnt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   3960
      TabIndex        =   59
      ToolTipText     =   "Unpaid Invoices For This Customer"
      Top             =   720
      Width           =   375
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices"
      Height          =   285
      Index           =   5
      Left            =   3120
      TabIndex        =   58
      Top             =   720
      Width           =   675
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Received"
      Height          =   285
      Index           =   2
      Left            =   5160
      TabIndex        =   57
      Top             =   720
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   56
      Top             =   720
      Width           =   945
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   55
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "diaARe04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

' See the UpdateTables prodecure for database revisions

'*************************************************************************************
' diaARe04a - Cash Receipts
'
' Created: (cjs)
'
' Revisions:
' 06/19/01 (nth) Allow transaction to post to journals other than the current month.
' 11/25/01 (nth) Improved with 'ledger' like interface
' 11/26/01 (nth) Added support for advanced payments
' 11/27/01 (nth) To support invoice type CA - cash advance invoice types
' 11/28/01 (nth) Added support for wire transfers and transaction fee's
' 11/28/01 (nth) Added 'short paid' and 'paid in full' to cmbapl
' 11/28/01 (nth) Prevented applied short payment amounts to be added to advanced
'                  payment amount
' 11/14/02 (nth) To remember last used checking account.
' 11/14/02 (nth) To remember last customer.
' 11/15/02 (nth) Added sFindMatchedAmounts automatically checks invoices or combination
'                  matching the check amount.
' 07/02/03 (nth) Added fDoesCheckExist to prevent check being applied mulitple
'                  times to the same invoice.
' 07/15/03 (nth) Fixed runtime error 13 reported by Jevco.
' 07/15/03 (nth) Remove multiple invoice matching.  Just match one for now.
' 09/18/03 (nth) Fixed runtime error with negative invoice amounts.
' 01/07/04 (nth) fixed CA type invoices not posting in the correct journal
' 01/20/04 (nth) Allow for a negative cash receipt
' 02/20/04 (nth) Show zero dollar invoices
' 05/11/04 (nth) Allow invoice number to be typed into customer dropdown.
' 07/15/04 (nth) Added support for multiple payers for a customer.
' 08/22/04 (nth) Change advanced payment logic to hit CR journal and not SJ
' 11/09/04 (nth) Added support for CihdTable.INVORIGIN in regards to advance payments
' 12/01/04 (nth) Added invoice filter by cash receipt received date.
' 01/03/05 (nth) Added total amount due field.
' 01/03/05 (nth) Increased invoice array from 200 to 512.
' 01/20/05 (nth) Changed iTotalItems to long datatype because of odd overflow error.
' 04/10/05 TEL Increased invoice array from 512 to 1024 - should use ReDim!
'
'*************************************************************************************

Option Explicit

'Private Const MAX_INVOICES = 1024


'Dim AdoQry As ADODB.Command
'Dim AdoParameter As ADODB.Parameter


Dim bCheckUsed As Boolean

Dim bOnLoad As Byte
Dim bNoCalc As Byte
Dim bCancel As Byte
Dim bNotNextGroup As Byte ' Disable the chkPaid_Click

' Cash Recipt Types (CATYPE)
' 0 = Cash
' 1 = Wire / Credit Card
' 2 = Check
Dim bCaType As Byte

Dim lInvoices() As Long


Dim iIndex As Integer
Dim iCurrPage As Integer

'(had to change to long because of overflow error on PROPLA database)
Dim iTotalItems As Long

Dim cCuArDisc As Currency
Dim cCuDays As Currency
Dim cCuNetDays As Currency
Dim cBottomLine() As Currency
Dim cAmount As Currency 'total amount paid
Dim cAdvAmt As Currency

' Debit \ Credit Accounts
Dim sAccount As String
Dim sCrCashAcct As String
Dim sCrDiscAcct As String
Dim sCrExpAcct As String
Dim sSJARAcct As String
Dim sCrRevAcct As String
Dim sCrCommAcct As String

Dim sTransFeeAcct As String
Dim sJournalID As String
Dim sApply(5) As String
Dim sMsg As String
Dim sIn As String

Dim optOrd As Integer

'Dim inShowAccounts As Boolean
Dim ignoreClicks As Boolean

' vInvoice Documentation
' 0  = Date
' 1  = Invoice
' 2  = PackSlip
' 3  = Total
' 4  = Total All (all meaning + freight and tax)
Private Const INV_TOTALAMOUNT = 4
' 5  = Total Due
Private Const INV_TOTALDUE = 5
' 6  = Amount Paid
Private Const INV_AMOUNTPAID = 6
' 7  = Freight
' 8  = Tax
' 9  = Discount Percentage
Private Const INV_DISCOUNTPERCENT = 9
' 10 = Days
' 11 = Marked Paid? (0 or 1)
Private Const INV_PAID = 11
' 12 = Apply to what account
Private Const INV_ACCOUNT = 12

'listindex values for accounts
Private Const INV_ACCOUNT_DISCOUNT = 0
Private Const INV_ACCOUNT_EXPENSE = 1
Private Const INV_ACCOUNT_COMMISSION = 2
Private Const INV_ACCOUNT_REVENUE = 3

' 13 = Invoice Type
' 14 = Paid in Full Flag
Private Const INV_PAIDINFULL = 14
Dim vInvoice() As Variant

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub ClearArrays()
    Erase vInvoice
    ReDim vInvoice(0 To 1024, 14)
End Sub

Private Sub chkAll_Click()
   Dim Index As Integer
   
   'iIndex = (Val(lblPge) - 1) * 9
   For Index = 1 To 9
      If Index + iIndex <= iTotalItems Then
         If chkPaid(Index) = vbUnchecked Then
            
            Dim pif As Boolean
            pif = False
            Dim due As Currency
            due = CCur(lblDue(Index))
            
            chkPaid(Index) = vbChecked
            'process credit - allow any or all to be applied
            If due < 0 Then
               If CCur(lblUnapplied) < 0 Then
                  If due < CCur(lblUnapplied) Then
                     txtPaid(Index) = Format(CCur(lblUnapplied), CURRENCYMASK)
                     pif = False
                  Else
                     txtPaid(Index) = due
                     pif = True
                  End If
               Else
                  txtPaid(Index) = "0.00"
               End If
               'process normal invoice
            Else
               
               If chkShortPayment.Value = 1 Then
                  txtPaid(Index) = Format(due, CURRENCYMASK)
               ElseIf cAmount <= CCur(lblTot) + due Then
                  ' not enough money
                  If cAmount < CCur(lblTot) Then
                     txtPaid(Index) = Format(0, CURRENCYMASK)
                  Else
                     txtPaid(Index) = Format(cAmount - CCur(lblTot), CURRENCYMASK)
                  End If
               Else
                  If CCur(vInvoice(iIndex + Index, INV_AMOUNTPAID)) = 0 Then
                     txtPaid(Index) = Format(vInvoice(iIndex + Index, INV_TOTALDUE), _
                             CURRENCYMASK)
                     'vInvoice(iIndex + Index, INV_AMOUNTPAID) = CCur(txtPaid(Index))
                     pif = True
                  Else
                     txtPaid(Index) = Format(vInvoice(iIndex + Index, INV_AMOUNTPAID), _
                             CURRENCYMASK)
                     pif = IIf(vInvoice(iIndex + Index, INV_PAIDINFULL) <> 0, True, False)
                  End If
               End If
            End If
            
            vInvoice(iIndex + Index, INV_AMOUNTPAID) = CCur(txtPaid(Index))
            
            'lblTot = Format(CCur(lblTot) + CCur(txtPaid(Index)), CURRENCYMASK)
            txtPaid(Index).enabled = True
            txtPaid(Index).SetFocus
            cmbapl(Index).ListIndex = vInvoice(iIndex + Index, INV_ACCOUNT)
            cmbapl(Index).enabled = True
            
         End If
      End If
      
   Next
   ' Remove the all selection
   chkAll.Value = False

End Sub

Private Sub chkPIF_Click(Index As Integer)
   
   If ignoreClicks Then
      Exit Sub
   End If
   
   vInvoice(iIndex + Index, INV_PAIDINFULL) = chkPIF(Index).Value
   ShowAccounts Index
End Sub

'*************************************************************************************

Private Sub cmbAct_Click()
   lblDsc(1) = UpdateActDesc(cmbAct)
   sCrCashAcct = cmbAct
End Sub

Private Sub cmbAct_LostFocus()
   lblDsc(1) = UpdateActDesc(cmbAct)
   sCrCashAcct = cmbAct
End Sub

Private Sub cmbAdvAct_Click()
   lblDsc(0) = UpdateActDesc(cmbAdvAct, lblDsc(0))
End Sub

Private Sub cmbAdvAct_LostFocus()
   lblDsc(0) = UpdateActDesc(cmbAdvAct, lblDsc(0))
End Sub

Private Sub cmbapl_Click(Index As Integer)
   If ignoreClicks Then
      Exit Sub
   End If
   vInvoice(iIndex + Index, INV_ACCOUNT) = cmbapl(Index).ListIndex
End Sub

Private Sub cmbapl_LostFocus(Index As Integer)
   ' if closing form, allow it to proceed
   If Me.ActiveControl.Name = "cmdCan" Then
      Exit Sub
   End If
   
   vInvoice(iIndex + Index, INV_ACCOUNT) = cmbapl(Index).ListIndex
End Sub

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
End Sub

Private Sub cmbCst_LostFocus()
   If Not bCancel Then
      cmbCst = CheckLen(cmbCst, 10)
      FindCustomer Me, cmbCst
      If lblNme = "" Then
         lblNme = "*** Customer Wasn't Found ***"
      End If
      sGetInvoices
   End If
End Sub

Private Sub cmbPageNo_Click()
    iCurrPage = Val(cmbPageNo)
    lblPge = iCurrPage
    sGetNextGroup
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   'cmbCst = ""
   bCancel = True
End Sub

Private Sub cmdCnl_Click()
   Dim bByte As Boolean
   Dim i As Integer
   Dim bResponse As Byte
   Dim sMsg As String
   
   For i = 1 To iTotalItems
      If vInvoice(i, 10) = vbChecked Then bByte = True
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
      cmbCst.enabled = True
      cmbAct.enabled = True
      txtDte.enabled = True
      txtChk.enabled = True
      txtAmt.enabled = True
      cmdSel.enabled = True
      txtCte.enabled = True
      txtFee.enabled = True
      optCsh.enabled = True
      optChk.enabled = True
      optWir.enabled = True
      sCloseBoxes
      txtDummy.enabled = False
      txtAmt.SetFocus
      ClearArrays
      
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
   sGetNextGroup
   
End Sub

Private Sub cmdDockLok_Click()
   Dim strCust As String
   Dim strDate As String
   Dim strChkNum As String
   
   Dim mfileInt As MfileIntegrator
   Set mfileInt = New MfileIntegrator
    
   Dim strIniPath As String
   strIniPath = App.Path & "\" & "MFileInit.ini"

   strCust = GetSectionEntry("MFILES_METADATA", "CUST", strIniPath)
   strDate = GetSectionEntry("MFILES_METADATA", "CRDATE", strIniPath)
   strChkNum = GetSectionEntry("MFILES_METADATA", "CHECKNUM", strIniPath)
    
   mfileInt.OpenXMLFile "Insert", "CashReceipts", "Cash Receipts", mfileInt.gsCRClassID, "Scan", txtChk.Text
   mfileInt.AddXMLMetaData "Customer", strCust, cmbCst.Text
   mfileInt.AddXMLMetaData "Date", strDate, txtDte.Text, MFDatatypeDate
   mfileInt.AddXMLMetaData "CheckNumber", strChkNum, txtChk.Text
        
   mfileInt.CloseXMLFile
   If mfileInt.SendXMLFileToMFile Then SysMsg "Record Successfully Indexed", True

   Set mfileInt = Nothing

End Sub

Private Function GetSectionEntry(ByVal strSectionName As String, ByVal strEntry As String, ByVal strIniPath As String) As String
   
   Dim X As Long
   Dim sSection As String, sEntry As String, sDefault As String
   Dim sRetBuf As String, iLenBuf As Integer, sFileName As String
   Dim sValue As String

   On Error GoTo modErr1
   
   sSection = strSectionName
   sEntry = strEntry
   sDefault = ""
   sRetBuf = String(256, vbNull) '256 null characters
   iLenBuf = Len(sRetBuf)
   sFileName = strIniPath
   X = GetPrivateProfileString(sSection, sEntry, _
                     "", sRetBuf, iLenBuf, sFileName)
   sValue = Trim(Left$(sRetBuf, X))
   
   If sValue <> "" Then
      GetSectionEntry = sValue
   Else
      GetSectionEntry = ""
   End If
   
   Exit Function
   
modErr1:
   sProcName = "GetSectionEntry"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function

Private Sub cmdFirstPage_Click()
    iCurrPage = 1
    lblPge = iCurrPage
    sGetNextGroup
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Cash Receipts"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdLastPage_Click()
    iCurrPage = Val(lblLst)
    lblPge = iCurrPage
    sGetNextGroup
End Sub

Private Sub cmdNextPage_Click()
    If iCurrPage >= Val(lblLst) Then Exit Sub
    iCurrPage = iCurrPage + 1
    lblPge = iCurrPage
    sGetNextGroup
    
'   iCurrPage = iCurrPage + 1
'   If iCurrPage > Val(lblLst) - 1 Then
'      iCurrPage = Val(lblLst)
'      cmdDn.enabled = False
'      cmdDn.Picture = Dsdn
'   Else
'      cmdDn.enabled = True
'      cmdDn.Picture = Endn
'   End If
'   If iCurrPage > 1 Then
'      cmdUp.enabled = True
'      cmdUp.Picture = Enup
'   Else
'      cmdUp.enabled = False
'      cmdUp.Picture = Dsup
'   End If
'   lblPge = iCurrPage
'   sGetNextGroup
End Sub

Private Sub cmdPrevPage_Click()
    If iCurrPage = 1 Then Exit Sub
    iCurrPage = iCurrPage - 1
    lblPge = iCurrPage
    sGetNextGroup

'   iCurrPage = iCurrPage - 1
'   If iCurrPage < 2 Then
'      iCurrPage = 1
'      cmdUp.enabled = False
'      cmdUp.Picture = Dsup
'   Else
'      cmdUp.enabled = True
'      cmdUp.Picture = Enup
'   End If
'   If iCurrPage < Val(lblLst) Then
'      cmdDn.enabled = True
'      cmdDn.Picture = Endn
'   Else
'      cmdDn.enabled = False
'      cmdDn.Picture = Dsdn
'   End If
'   lblPge = iCurrPage
'   sGetNextGroup
End Sub

Private Sub cmdPst_Click()
   Dim bResponse As Byte
   If Left(lblNme, 3) = "***" Then
      sMsg = "Customer Is Invalid."
      MsgBox sMsg, vbInformation, Caption
      cmbCst.SetFocus
      Exit Sub
   End If
   If Left(lblDsc(0), 3) = "***" Then
      sMsg = "Amount Not Applied Account Is Invalid."
      MsgBox sMsg, vbInformation, Caption
      cmbAdvAct.SetFocus
      Exit Sub
   End If
   
   If chkShortPayment.Value = vbChecked Then
      If CCur("0" & lblUnapplied) <> 0 Then
         sMsg = "When creating a credit memo, the Unapplied amount must be negative and the available amount at the bottom must be zero."
         MsgBox sMsg, vbInformation, Caption
         cmbAdvAct.SetFocus
         Exit Sub
      Else
         sMsg = "Create credit memo for " & Format(CCur(txtAdvAmt), CURRENCYMASK) _
            & " to account '" & lblDsc(0) & "'?"
         If MsgBox(sMsg, vbYesNo, "Create credit memo and cash receipt") <> vbYes Then
            Exit Sub
         End If
      End If
      
   Else
      If InStr(1, txtAdvAmt, "-") > 0 Then
         MsgBox "You must check the create credit memo box for short payments (negative amount)"
         Exit Sub
      End If
   End If
         
   sMsg = "Are You Ready To Add This Cash Receipt?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      If Not fCheckExists(Trim(txtChk)) Then
         MouseCursor 13
         sPostCheck
         cmdcnl.enabled = False
         cmdpst.enabled = False
         cmdDockLok.enabled = False
         MouseCursor 0
         sCloseBoxes
      Else
         txtChk.enabled = True
         txtChk.SetFocus
      End If
   Else
      CancelTrans
   End If
End Sub

Private Sub cmdSel_Click()
   Dim sMsg As String
   If Len(Trim(txtChk)) = 0 And Not optCsh Then
      If optChk Then sMsg = "Checks Require A Valid Check Number."
      If optWir Then sMsg = "Wires Require A Valid Wire Number."
      MsgBox sMsg, vbInformation, Caption
      txtChk.SetFocus
      Exit Sub
   End If
   If cAmount < 0 Then
      MsgBox "Amount Must Be Zero Or Greater.", vbInformation, Caption
      txtAmt.SetFocus
      Exit Sub
   End If
   
   fGetCustomerInvoices (optOrd)
   
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
   sGetNextGroup
End Sub

Private Sub cmdVew_Click()
   With diaARe12a
      .bRemote = True
      .Left = Me.Left + 200
      .Top = Me.Top + 200
      .cmbCst = cmbCst
      .Show
   End With
End Sub



Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CurrentJournal "CR", ES_SYSDATE, sJournalID
      b = fGetCashAccounts()
      If b = 3 Then
         MouseCursor 0
         MsgBox "One Or More Receivable Accounts Are Not Active." & vbCr _
            & "Please Set All Cash Accounts In The " & vbCr _
            & "System Setup, Administration Section.", _
            vbInformation, Caption
         Sleep 500
         Unload Me
         Exit Sub
      End If
      MouseCursor 13
      sFillCombo
      If cUR.CurrentCustomer <> "" Then
         cmbCst = cUR.CurrentCustomer
      End If
      FindCustomer Me, cmbCst
      If lblNme = "" Then
         lblNme = "*** Customer Wasn't Found ***"
      End If
      sGetInvoices
      If DocumentLokEnabled Then cmdDockLok.Visible = True Else cmdDockLok.Visible = False
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim j As Integer
   ReDim lInvoices(0 To 1)
   ReDim cBottomLine(0 To 1)
   ReDim vInvoice(0 To 1024, 14)
   
   FormLoad Me
   sFormatControls
   
   sCurrForm = Caption
   sAccount = GetSetting("Esi2000", "Fina", "LastCAccount", sAccount)
   

   
'   sSql = "SELECT CUREF,CUNICKNAME,CUNAME,CUARDISC," _
'          & "CUDAYS,CUNETDAYS FROM CustTable WHERE CUREF= ? "
'   Set AdoQry = New ADODB.Command
'   AdoQry.CommandText = sSql
   
'   Set AdoParameter = New ADODB.Parameter
'   AdoParameter.Type = adChar
'   AdoParameter.SIZE = 10
'   AdoQry.parameters.Append AdoParameter
   
   
   'sApply(0) = "Paid In Full"
   'sApply(1) = "Short Paid"
   sApply(INV_ACCOUNT_DISCOUNT) = "Discount"
   sApply(INV_ACCOUNT_EXPENSE) = "Expense"
   sApply(INV_ACCOUNT_COMMISSION) = "Commission"
   sApply(INV_ACCOUNT_REVENUE) = "Revenue"
   For i = 0 To 3
      For j = 1 To 9
         AddComboStr cmbapl(j).hWnd, sApply(i)
      Next
   Next
   
   optChk = True
      
   optOrd = 0
   
   txtAmt = "0.00"
   txtFee = "0.00"
   txtAdvAmt = "0.00"
   
   cmdVew.Picture = Resources.imgVew.Picture
   
   sCloseBoxes
   
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Len(cmbCst) Then
      cUR.CurrentCustomer = cmbCst
      SaveCurrentSelections
   End If
   If Len(cmbAct) Then SaveSetting "Esi2000", _
          "Fina", "LastCAccount", Trim(cmbAct)
   
   'Set AdoParameter = Nothing
   'Set AdoQry = Nothing
   
   FormUnload
   Set diaARe04a = Nothing
End Sub

Private Sub sFormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDte = Format(Now, "mm/dd/yy")
   txtCte = Format(Now, "mm/dd/yy")
End Sub


Private Sub optChk_Click()
   txtFee.enabled = True
   txtChk.enabled = True
   txtAmt.enabled = True
   txtCte.enabled = True
   bCaType = 2
   If txtChk = "Cash" Then
      txtChk = ""
   End If
End Sub

Private Sub optCsh_Click()
   txtFee.enabled = False
   'txtChk.Enabled = False
   txtAmt.enabled = True
   txtCte.enabled = True
   bCaType = 0
   txtChk = Format(ES_SYSDATE, "yymmdd") & "-CASH"
End Sub

Private Sub chkPaid_Click(Index As Integer)
   Dim cTotalItem As Currency
   'Dim cMoneyLeft As Currency
   
   If ignoreClicks Then
      Exit Sub
   End If
   
   If bNotNextGroup = 0 Then
      
      'cMoneyLeft = CCur(txtAmt) - CCur(lblTot)
      
      Dim pif As Boolean
      pif = False
      Dim due As Currency
      due = CCur(lblDue(Index))
      If chkPaid(Index) = vbChecked Then
         
         'process credit - allow any or all to be applied
         If due < 0 Then
            If CCur(lblUnapplied) < 0 Then
               If due < CCur(lblUnapplied) Then
                  txtPaid(Index) = Format(CCur(lblUnapplied), CURRENCYMASK)
                  pif = False
               Else
                  txtPaid(Index) = due
                  pif = True
               End If
            Else
               txtPaid(Index) = "0.00"
            End If
            
            
            'process normal invoice
         Else
            ' if credit memo is being created, just set pay amount = remaining balance
            If chkShortPayment.Value = 1 Then
                  txtPaid(Index) = Format(due, CURRENCYMASK)
            ElseIf cAmount <= CCur(lblTot) + due Then
               ' not enough money
               If cAmount < CCur(lblTot) And chkShortPayment.Value = 0 Then
                  txtPaid(Index) = Format(0, CURRENCYMASK)
               Else
                  txtPaid(Index) = Format(cAmount - CCur(lblTot), CURRENCYMASK)
               End If
            Else
               If CCur(vInvoice(iIndex + Index, INV_AMOUNTPAID)) = 0 Then
                  txtPaid(Index) = Format(vInvoice(iIndex + Index, INV_TOTALDUE), _
                          CURRENCYMASK)
                  'vInvoice(iIndex + Index, INV_AMOUNTPAID) = CCur(txtPaid(Index))
                  pif = True
               Else
                  txtPaid(Index) = Format(vInvoice(iIndex + Index, INV_AMOUNTPAID), _
                          CURRENCYMASK)
                  pif = IIf(vInvoice(iIndex + Index, INV_PAIDINFULL) <> 0, True, False)
               End If
            End If
         End If
         
         vInvoice(iIndex + Index, INV_AMOUNTPAID) = CCur(txtPaid(Index))
         
         'lblTot = Format(CCur(lblTot) + CCur(txtPaid(Index)), CURRENCYMASK)
         txtPaid(Index).enabled = True
         txtPaid(Index).SetFocus
         cmbapl(Index).ListIndex = vInvoice(iIndex + Index, INV_ACCOUNT)
         cmbapl(Index).enabled = True
      Else
         chkPaid(Index) = vbUnchecked
         txtPaid(Index) = "0.00"
         vInvoice(iIndex + Index, INV_AMOUNTPAID) = 0
         cmbapl(Index).ListIndex = -1
         'cmbapl(Index) = ""
         cmbapl(Index).enabled = False
         txtPaid(Index).enabled = False
         'sUpdateTotals
      End If
      chkPIF(Index).Value = IIf(pif, vbChecked, vbUnchecked)
      vInvoice(iIndex + Index, INV_PAIDINFULL) = pif
      vInvoice(iIndex + Index, INV_PAID) = chkPaid(Index).Value
      sUpdateTotals
   End If
   ShowAccounts Index
End Sub

Private Sub chkPaid_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub chkPaid_KeyUp(Index As Integer, _
                          KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub chkPaid_LostFocus(Index As Integer)
   ' if closing form, allow it to proceed
   If Me.ActiveControl.Name = "cmdCan" Then
      Exit Sub
   End If
   
   vInvoice(iIndex + Index, 10) = chkPaid(Index).Value
End Sub

Private Sub optOrderBy_Click(Index As Integer)
   If optOrderBy(0) Then
      optOrd = 0
   Else
      optOrd = 1
   End If
End Sub


Private Sub optWir_Click()
   txtFee.enabled = True
   txtChk.enabled = True
   txtAmt.enabled = True
   txtCte.enabled = True
   bCaType = 1
   If txtChk = "Cash" Then txtChk = ""
End Sub

Private Sub txtAdvAmt_LostFocus()
   Dim cMoneyLeft As Currency
   Dim bTurnOn As Byte
   
   ' if closing form, allow it to proceed
   If Me.ActiveControl.Name = "cmdCan" Then
      Exit Sub
   End If
   
   If Trim(txtAdvAmt) = "" Then txtAdvAmt = "0.00"
   txtAdvAmt = CheckLen(txtAdvAmt, 10)
   txtAdvAmt = Format(txtAdvAmt, "#,###,##0.00")
   cAdvAmt = CCur(txtAdvAmt)
   bTurnOn = False
   If cAdvAmt > 0 Then
      'cMoneyLeft = cAmount - fSumSelectedItems + CCur(txtFee & "0")
      cMoneyLeft = cAmount - CashApplied()
      If cAdvAmt > cMoneyLeft Then
         'sMsg = "Allocated Amount Cannot Exceed Cash Receipt."
         'MsgBox sMsg, vbInformation, Caption
         'txtAdvAmt.SetFocus
      Else
         'lblTot = Format(cAmount - cMoneyLeft + cAdvAmt, "#,###,##0.00")
         'bTurnOn = True
         'If lblTot = txtAmt Then
         '    cmdPst.Enabled = True
         'End If
      End If
   End If
   bTurnOn = cAdvAmt <> 0
   chkAdv.enabled = bTurnOn
   sUpdateTotals
End Sub

Private Sub txtAmt_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtAmt_LostFocus()
   If Trim(txtAmt) = "" Then
      txtAmt = "0.00"
   Else
      txtAmt = CheckLen(txtAmt, 10)
      txtAmt = Format(txtAmt, "#,###,##0.00")
   End If
   cAmount = CCur("0" & txtAmt) '+ CCur("0" & txtFee)
End Sub

Private Sub txtChk_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtChk_LostFocus()
   If Not bCancel Then
      txtChk = CheckLen(txtChk, 12)
      If fCheckExists(txtChk) Then
         txtChk.SetFocus
      End If
   End If
End Sub

Private Sub txtCte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtCte_LostFocus()
   txtCte = CheckDate(txtCte)
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub sFillCombo()
   Dim rdoCst As ADODB.Recordset
   Dim rdoAct As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   Dim lCheck As Long
   Dim sOptions As String
   
   On Error GoTo DiaErr1
   
   ' Customer
   sSql = "SELECT DISTINCT CUNICKNAME FROM CustTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst)
   If bSqlRows Then
      With rdoCst
         cmbCst = "" & Trim(!CUNICKNAME)
         While Not .EOF
            AddComboStr cmbCst.hWnd, "" & Trim(!CUNICKNAME)
            .MoveNext
         Wend
      End With
   End If
   Set rdoCst = Nothing
   
   ' Accounts
   sSql = "Qry_FillLowAccounts"
   'bSqlRows = clsAdoCon.GetDataSet(sSql,rdoAct, ES_FORWARD)
   'sSql = "SELECT GLACCTNO FROM GlacTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         While Not .EOF
            AddComboStr cmbAdvAct.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Wend
      End With
      cmbAdvAct.ListIndex = 0
   End If
   Set rdoAct = Nothing
   lblDsc(0) = UpdateActDesc(cmbAdvAct)
   
   
   sSql = "SELECT GLACCTNO FROM GlacTable WHERE GLCASH = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         While Not .EOF
            AddComboStr cmbAct.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Wend
      End With
      
      '        cmbAct.ListIndex = 0
      '        If Trim(sAccount) <> "" Then
      '            'cmbAct = sAccount
      '        End If
      SetComboBox cmbAct, sAccount
   Else
      ' Multiple cash accounts not found so use
      ' the default cash account
'      cmbAct.enabled = False
'      cmbAct = sCrCashAcct
      AddComboStr cmbAct.hWnd, sCrCashAcct
      cmbAct.ListIndex = 0
      cmbAct.enabled = False
   
   End If
   lblDsc(1) = UpdateActDesc(cmbAct)
   
   Set rdoCst = Nothing
   Set rdoAct = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub sGetInvoices()
   Dim RdoInv As ADODB.Recordset
   'Dim rdoPay As ADODB.Recordset
   Dim rdoCst As ADODB.Recordset
   Dim iTotal As Integer
   Dim cInvTot As Currency
   Dim sCust As String

   On Error GoTo DiaErr1
   sCloseBoxes
   ClearArrays
   
   sCust = Compress(cmbCst)
   txtDummy.enabled = False
   
   sIn = "('" & sCust & "'"
   sSql = "SELECT CPPAYER FROM CpayTable WHERE CPCUST = '" & sCust & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst)
   If bSqlRows Then
      With rdoCst
         While Not .EOF
            sIn = sIn & ",'" & "" & Trim(.Fields(0)) & "'"
            .MoveNext
         Wend
         .Cancel
      End With
      Set rdoCst = Nothing
   End If
   sIn = sIn & ")"
   
   sSql = "SELECT DISTINCT INVNO,INVPRE,INVCUST,INVTYPE,INVTOTAL," _
          & "INVPAY,INVFREIGHT,INVTAX FROM CihdTable WHERE " _
          & "INVCUST IN " & sIn & " AND INVTYPE<>'TM' AND INVPIF=0 " _
          & "AND INVCANCELED=0 AND INVDATE<='" & txtDte & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_STATIC)
   
'   RdoInv.MoveLast   'record count not accurate until last record seen
'   RdoInv.MoveFirst
'   ReDim lInvoices(0 To RdoInv.RecordCount + 1)

   If bSqlRows Then
      With RdoInv
         .MoveLast   'record count not accurate until last record seen
         .MoveFirst
         ReDim lInvoices(0 To .RecordCount + 1)

         Do Until .EOF
            iTotal = iTotal + 1
            lInvoices(iTotal) = !InvNo
            cInvTot = cInvTot + (!INVTOTAL - !INVPAY)
            .MoveNext
         Loop
         .Cancel
      End With
   Else
      iTotal = 0
   End If
   
   Set RdoInv = Nothing
   cmdpst.enabled = False
   cmdcnl.enabled = False
   cmdDockLok.enabled = False

   lblCnt = Format(iTotal, "##0")
   lblOwe = Format(cInvTot, CURRENCYMASK)
   Exit Sub
DiaErr1:
   sProcName = "sGetInvoices"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtDte_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   sGetInvoices
End Sub

Private Sub txtFee_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtFee_LostFocus()
   If txtFee = "" Then
      txtFee = "0.00"
   Else
      txtFee = CheckLen(txtFee, 10)
      txtFee = Format(Abs(txtFee), CURRENCYMASK)
   End If
   'cAmount = CCur("0" & txtAmt) + CCur("0" & txtFee)
End Sub

Private Sub txtPaid_GotFocus(Index As Integer)
   ' if closing form, allow it to proceed
   '    If Me.ActiveControl.Name = "xx" Then
   '        Exit Sub
   '    End If
   '
   SelectFormat Me
   'On Error Resume Next
   '    If CCur(lblInvAmt(index)) = CCur(txtPaid(index)) Then
   '        cmbapl(index).ListIndex = 0
   '    Else
   '        If cmbapl(index).ListIndex < 1 Then
   '            cmbapl(index).ListIndex = 1
   '        End If
   '    End If
End Sub

Private Sub txtPaid_KeyDown(Index As Integer, _
                            KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
End Sub

Private Sub txtPaid_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
End Sub

Private Sub txtPaid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim cPay As Currency
   Dim cDue As Currency
   
   If KeyCode = vbKeyPageUp Then
      cmdUp_Click
   ElseIf KeyCode = vbKeyPageDown Then
      cmdDn_Click
   Else
      '        If IsNumeric(txtPaid(index)) Then
      '            cPay = CCur(txtPaid(index))
      '            cDue = CCur(lblInvAmt(index))
      '
      '            If cPay = cDue Then
      '                cmbapl(index).ListIndex = 0
      '            End If
      '
      '            If Abs(cPay) < Abs(cDue) Then
      '                If cmbapl(index).ListIndex < 1 Then
      '                    cmbapl(index).ListIndex = 1
      '                End If
      '            End If
      '        End If
   End If
End Sub


Private Sub txtPaid_LostFocus(Index As Integer)
   Dim i As Integer
   Dim cTotalItem As Currency
   
   On Error GoTo DiaErr1
   
   If Me.ActiveControl.Name = "cmdCan" Then
      Exit Sub
   End If
   
   'if credit, require a negative entry amount
   If CCur(lblDue(Index)) < 0 And CCur(txtPaid(Index)) > 0 Then
      MsgBox "You must enter a negative payment amount for a credit", vbInformation, Caption
      txtPaid(Index).SetFocus
      Exit Sub
   End If
   
   ' Why did you start at one instead of zero Colin?
   i = Index + iIndex
   
   txtPaid(Index) = CheckLen(txtPaid(Index), 11)
   txtPaid(Index) = Format(txtPaid(Index), "#,###,##0.00")
   
   If Val(txtPaid(Index)) = 0 And CCur(lblDue(Index)) > 0 Then
      'chkPaid(Index).Value = vbUnchecked
   Else
      vInvoice(i, INV_AMOUNTPAID) = txtPaid(Index)
   End If
   
   vInvoice(i, INV_TOTALDUE) = lblDue(Index)
   vInvoice(i, INV_AMOUNTPAID) = txtPaid(Index)
   'vInvoice(i, INV_ACCOUNT) = cmbapl(index).ListIndex
   
   sUpdateTotals
   ShowAccounts (Index)
   Exit Sub

DiaErr1:
   sProcName = "txtPaidLostFocus"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub sCloseBoxes()
   Dim i As Integer
   
   For i = 1 To 9
      chkPaid(i).Caption = ""
      chkPaid(i).enabled = False
      chkPaid(i).Value = vbUnchecked
      lblDte(i) = ""
      lblInv(i) = ""
      lblPs(i) = ""
      lblPs(i).enabled = False
      lblInvAmt(i) = "0.00"
      lblInvAmt(i).enabled = False
      txtPaid(i) = "0.00"
      txtPaid(i).enabled = False
      lblDue(i) = "0.00"
      lblDue(i).enabled = False
      cmbapl(i).ListIndex = -1
      chkPIF(i).enabled = False
      chkPIF(i).Value = vbUnchecked
   Next
   
   cmdcnl.enabled = False
   cmdpst.enabled = False
   
   cmdDn.enabled = False
   cmdUp.enabled = False
   cmdDockLok.enabled = False
   
   cmdUp.Picture = Dsup
   cmdDn.Picture = Dsdn
   
   lblPge = ""
   lblLst = ""
   cmbPageNo.Clear
   cmbPageNo.enabled = False
   cmdFirstPage.enabled = False
   cmdLastPage.enabled = False
   cmdNextPage.enabled = False
   cmdPrevPage.enabled = False
   
   lblTot = "0.00"
End Sub

Private Sub fGetCustomerInvoices(optOrd As Integer)
   Dim RdoInv As ADODB.Recordset
   Dim i As Integer
   Dim iNetDays As Integer
   Dim cTotal As Currency
   Dim cDiscount As Currency
   Dim cPaid As Currency
   Dim sPaidDate As Date
   Dim sPoInvDate As Date
   Dim sCust As String
   Dim sType As String
   Dim lTemp As Long
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   lblTot = 0#
   iTotalItems = 0
   
   Erase cBottomLine
   Erase lInvoices
   ClearArrays
   
   sCust = Compress(cmbCst)
   sPaidDate = Format(txtDte, "mm/dd/yyyy")
   cTotal = 0
   
'   sSql = "SELECT INVNO,INVPRE,INVCUST,INVPACKSLIP,INVTOTAL,INVPAY," _
'          & "INVTAX,INVFREIGHT,INVDATE,INVTYPE FROM CihdTable WHERE " _
'          & "INVCUST IN " & sIn & " AND (INVPIF=0 AND INVTYPE<>'TM' AND " _
'          & "INVCANCELED=0 AND INVDATE <= '" & txtDte & "') ORDER BY INVDATE, INVNO desc"
   
   If optOrd = 0 Then
        sSql = "SELECT INVNO,INVPRE,INVCUST,INVPACKSLIP,INVTOTAL,INVPAY," _
            & "INVTAX,INVFREIGHT,INVDATE,INVTYPE FROM CihdTable WHERE " _
            & "INVCUST IN " & sIn & " AND (INVPIF=0 AND INVTYPE<>'TM' AND " _
            & "INVCANCELED=0 AND INVDATE <= '" & txtDte & "') ORDER BY INVDATE, INVNO"
   Else
        sSql = "SELECT INVNO,INVPRE,INVCUST,INVPACKSLIP,INVTOTAL,INVPAY," _
            & "INVTAX,INVFREIGHT,INVDATE,INVTYPE FROM CihdTable WHERE " _
            & "INVCUST IN " & sIn & " AND (INVPIF=0 AND INVTYPE<>'TM' AND " _
            & "INVCANCELED=0 AND INVDATE <= '" & txtDte & "') ORDER BY INVNO, INVDATE"
   End If
   
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_STATIC)
   
  
   If bSqlRows Then
      With RdoInv
         .MoveLast   'record count not accurate until last record seen
         .MoveFirst
         ReDim lInvoices(0 To .RecordCount + 1)
         ReDim cBottomLine(0 To .RecordCount + 1)
         ReDim vInvoice(0 To .RecordCount + 10, 14)    'BBS Changed for Ticket #41369 because this array has to be an even divisible of 9
         While Not .EOF
            
            sPoInvDate = Format(!INVDATE, "mm/dd/yyyy")
            
            'If Left(sType, 2) = "-S" Or Left(sType, 2) = "-P" Then
            '    If (sPaidDate - cCuDays) <= sPoInvDate Then
            '        cDiscount = ((!INVTOTAL) * cCuArDisc) / 100
            '    End If
            'End If
            
            cTotal = !INVTOTAL
            cPaid = (cTotal - !INVPAY)
            
            'If cTotal <> 0 Then
            iTotalItems = iTotalItems + 1
            cBottomLine(iTotalItems) = !INVFREIGHT + !INVTAX
            lInvoices(iTotalItems) = !InvNo
            
            vInvoice(iTotalItems, 0) = Format(!INVDATE, "mm/dd/yyyy")
            vInvoice(iTotalItems, 1) = "" & Trim(!INVPRE) & Format(!InvNo, "000000")
            vInvoice(iTotalItems, 2) = Format("" & (!INVPACKSLIP), "000000")
            vInvoice(iTotalItems, INV_TOTALAMOUNT) = Format(cTotal, "0.00")
            vInvoice(iTotalItems, INV_TOTALDUE) = Format(cPaid, "0.00")
            vInvoice(iTotalItems, INV_AMOUNTPAID) = Format(0, "0.00")
            vInvoice(iTotalItems, 7) = Format(!INVFREIGHT, "0.00")
            vInvoice(iTotalItems, 8) = Format(!INVTAX, "0.00")
            vInvoice(iTotalItems, INV_DISCOUNTPERCENT) = Format(cCuArDisc, "0.000")
            vInvoice(iTotalItems, 10) = Format(cCuDays, "0")
            vInvoice(iTotalItems, INV_PAID) = 0
            vInvoice(iTotalItems, INV_ACCOUNT) = -1
            vInvoice(iTotalItems, 13) = !INVTYPE
            vInvoice(iTotalItems, INV_PAIDINFULL) = 0
            'Debug.Print !INVTYPE & " " & !INVNO & " " & cPaid
            'End If
            
            .MoveNext
            
         Wend
         .Cancel
      End With
   End If
   
   iIndex = 0
   If iTotalItems > 0 Then
      iIndex = fFindMatchedAmounts
      iCurrPage = 1
      If iIndex > 0 Then
         If iIndex > 9 Then
            lTemp = (iIndex * 100&)
            iCurrPage = CInt(((lTemp / 100) / 9) + 0.5)
            iIndex = ((iCurrPage - 1) * 9)
         Else
            iIndex = 0
         End If
      End If
      
      
      lblLst = CLng((((iTotalItems * 100) / 100) / 9) + 0.5)
      cmbPageNo.Clear
      For i = 1 To lblLst
        cmbPageNo.AddItem Trim(str(i))
      Next i
      If cmbPageNo.ListCount > 0 Then cmbPageNo.enabled = True
      
      For i = 1 To 9
         If i > iTotalItems Then
            Exit For
         End If
         
         lblDte(i) = Format(vInvoice(i + iIndex, 0), "mm/dd/yy")
         lblInv(i) = vInvoice(i + iIndex, 1)
         lblPs(i) = vInvoice(i + iIndex, 2)
         lblPs(i).enabled = True
         lblInvAmt(i) = Format(vInvoice(i + iIndex, INV_TOTALAMOUNT), "#,###,##0.00")
         lblInvAmt(i).enabled = True
         txtPaid(i) = Format(vInvoice(i + iIndex, INV_AMOUNTPAID), "#,###,##0.00")
         lblDue(i) = Format(vInvoice(i + iIndex, INV_TOTALDUE), "#,###,##0.00")
         lblDue(i).enabled = True
         chkPaid(i).Caption = "___"
         chkPaid(i).enabled = True
         chkPaid(i).Value = vInvoice(i + iIndex, INV_PAID)
         chkPIF(i).Value = vInvoice(i + iIndex, INV_PAIDINFULL)
         
         '                On Error Resume Next
         '                txtDummy.enabled = True
         '                cmdPst.enabled = True
         '                cmdCnl.enabled = True
         '                cmdSel.enabled = False
         '                txtPaid(1).SetFocus
      Next
      
      On Error Resume Next
      txtDummy.enabled = True
      cmdpst.enabled = True
      cmdcnl.enabled = True
      cmdSel.enabled = False
      cmdDockLok.enabled = True
    
          txtPaid(1).SetFocus
      
      lblPge = iCurrPage
      
      If Val(lblLst) > 1 Then
         cmdDn.enabled = True
         cmdDn.Picture = Endn
      End If
      
      If Val(lblPge) > 1 Then
         cmdUp.enabled = True
         cmdUp.Picture = Enup
      End If
      
      cmbCst.enabled = False
      cmbAct.enabled = False
      txtDte.enabled = False
      txtCte.enabled = False
      txtChk.enabled = False
      txtAmt.enabled = False
      txtFee.enabled = False
      optCsh.enabled = False
      optChk.enabled = False
      optWir.enabled = False
      
      chkPaid(1).SetFocus
      sUpdateTotals
      
      cmdFirstPage.enabled = True
      cmdLastPage.enabled = True
      cmdNextPage.enabled = True
      cmdPrevPage.enabled = True

   End If
   
   Set RdoInv = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "fGetCustomerInv"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub sGetNextGroup()
   Dim i As Integer
   Dim sType As String
   
   ignoreClicks = True
   
   bNotNextGroup = 1
   
   iIndex = (Val(lblPge) - 1) * 9
   On Error GoTo DiaErr1
   For i = 1 To 9
      
      '            If i = 6 Then
      '                Debug.Print "6"
      '            End If
      If i + iIndex > iTotalItems Then
         chkPaid(i).Caption = ""
         chkPaid(i).enabled = False
         chkPaid(i).Value = vbUnchecked
         lblDte(i) = ""
         lblInv(i) = ""
         lblPs(i) = ""
         lblPs(i).enabled = False
         lblInvAmt(i) = ""
         lblInvAmt(i).enabled = False
         lblDue(i) = ""
         lblDue(i).enabled = False
         txtPaid(i) = ""
         txtPaid(i).enabled = False
         cmbapl(i).enabled = False
         cmbapl(i).ListIndex = -1
         chkPIF(i).Value = 0
         chkPIF(i).enabled = False
      Else
         lblDte(i) = Format(vInvoice(i + iIndex, 0), "mm/dd/yy")
         lblInv(i) = vInvoice(i + iIndex, 1)
         lblPs(i) = Format(vInvoice(i + iIndex, 2), "000000")
         lblPs(i).enabled = True
         lblInvAmt(i) = Format(vInvoice(i + iIndex, INV_TOTALAMOUNT), "#,###,##0.00")
         lblInvAmt(i).enabled = True
         lblDue(i) = Format(vInvoice(i + iIndex, INV_TOTALDUE), "#,###,##0.00")
         lblDue(i).enabled = True
         
         
         'txtPaid(i) = Format(vInvoice(i + iIndex, INV_AMOUNTPAID), "#,###,##0.00")
         
         If Val(vInvoice(i + iIndex, INV_PAID)) > 0 Then
            'chkPaid(i).Value vInvoice(i + Index, INV_PAID)
            txtPaid(i) = Format(vInvoice(i + iIndex, INV_AMOUNTPAID), "#,###,##0.00")
            txtPaid(i).enabled = True
         Else
            txtPaid(i).Text = "0.00"
            txtPaid(i).enabled = False
         End If
         
         chkPaid(i).Caption = "___"
         chkPaid(i).Value = vInvoice(i + iIndex, INV_PAID)
         chkPaid(i).enabled = True
         
         chkPIF(i).Value = vInvoice(i + iIndex, INV_PAIDINFULL)
         If chkPaid(i).Value = 1 And (CCur(lblDue(i).Caption) - CCur(txtPaid(i))) = 0 Then
            chkPIF(i).enabled = True
         Else
            chkPIF(i).enabled = False
            'chkPIF(i).Value = 0
         End If
         
         cmbapl(i).ListIndex = Val(vInvoice(i + iIndex, INV_ACCOUNT))
      End If
   Next
   On Error Resume Next
   bNotNextGroup = 0
   
   chkPaid(1).SetFocus
   ignoreClicks = False
   Exit Sub
   
DiaErr1:
   sProcName = "getnextgr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   ignoreClicks = False
End Sub

Private Sub sUpdateTotals()
   Dim cTotal As Currency
   
   On Error GoTo DiaErr1
   
   'cTotal = fSumSelectedItems
   'cTotal = cTotal + CCur(txtAdvAmt)
   cTotal = CashApplied()
   lblTot = Format(cTotal, "#,###,##0.00")
   lblUnapplied = Format(cAmount - cTotal, "#,###,##0.00")
   'txtAdvAmt = lblUnapplied
   
   If cTotal = cAmount Then
      cmdpst.enabled = True
      cmdDockLok.enabled = True
   Else
      cmdpst.enabled = False
      cmdDockLok.enabled = False
   End If
   Exit Sub
DiaErr1:
   sProcName = "sUpdateTotals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub sPostCheck()
   Dim i As Integer
   Dim lTrans As Long
   Dim iRef As Integer
   Dim sCust As String
   Dim sCDate As String
   Dim sRDate As String
   Dim sAdvPay As String
   Dim sAdvAct As String
   Dim sCheck As String
   Dim sEDate As String
   Dim sColumn As String
   Dim sNow As String
   Dim lNewInv As Long
   Dim rdoAct As ADODB.Recordset
   Dim TotalReceipt As Currency
   Dim CashAvailable As Currency
   Dim RefundAvailable As Currency
   Dim unappliedAmt As Currency
   
   On Error GoTo DiaErr1
   
   sCDate = txtCte
   sRDate = txtDte
   sEDate = Format(ES_SYSDATE, "mm/dd/yy")
   sCust = Compress(cmbCst)
   sAdvAct = Compress(cmbAdvAct)
   sCrCashAcct = Compress(cmbAct)
   sCheck = Trim(txtChk)
   TotalReceipt = CCur("0" & txtAmt)
   CashAvailable = TotalReceipt
   unappliedAmt = CCur(txtAdvAmt & "0")

   
   ' Allow posting to journals other than that of the current month
   sJournalID = GetOpenJournal("CR", sRDate)
   If sJournalID <> "" Then
      'lTrans = GetNextTransaction(sJournalID)
   Else
      sMsg = "No Open Cash Recipts Journal Found For " & sRDate & "."
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   
'   'not required
'   ' if short payment, an open sales journal is required
'   Dim SalesJournalID As String
'   If unappliedAmt < 0 Then
'      SalesJournalID = GetOpenJournal("SJ", sRDate)
'      If SalesJournalID = "" Then
'         sMsg = "No Open Sales Journal Found For " & sRDate & "."
'         MsgBox sMsg, vbInformation, Caption
'         Exit Sub
'      End If
'   End If
'
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   ignoreClicks = True
   
   'create a general ledger transaction object for creation of CR debits and credits
   Dim gl As GLTransaction
   Set gl = New GLTransaction
   gl.JournalID = sJournalID 'automatically sets next transaction
   gl.InvoiceDate = CDate(sRDate)
   
'   'if short payment create a sales hournal transaction object for credit memo debit and credit
'   Dim glSJ As GLTransaction
'   If unappliedAmt < 0 Then
'      Set glSJ = New GLTransaction
'      glSJ.JournalID = SalesJournalID 'automatically sets next transaction
'      glSJ.InvoiceDate = CDate(sRDate)
'   End If
   
   ''''''''''''''''''''''''''''''''''''''''''
   
   'PAYMENT NOT APPLIED TO INVOICES
   ' Deal with overpayments and record advance payments if nessary
'   Dim unappliedAmt As Currency
'   'unappliedAmt = CCur("0" & txtAdvAmt)  ' chage logic to work with negative number
'   unappliedAmt = CCur(txtAdvAmt & "0")
   If unappliedAmt > 0 Then
      If chkAdv Then
         ' Create CA Invoice
         lNewInv = fGetNextInvoice
         sSql = "INSERT INTO CihdTable (INVNO,INVTYPE,INVTOTAL,INVCUST," _
                & "INVDATE,INVSHIPDATE,INVORIGIN) " _
                & "VALUES(" & lNewInv & ",'CA'," _
                & (-1 * unappliedAmt) & ",'" _
                & Compress(cmbCst) & "','" _
                & txtDte & "','" & txtDte _
                & "','" & sCheck & "')"
         clsADOCon.ExecuteSql sSql
         
         Dim inv As New ClassARInvoice
         inv.SaveLastInvoiceNumber lNewInv
         
      End If
      
      
      ' Debit Cash
      '        iRef = iRef + 1
      '        sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT," _
      '            & "DCACCTNO,DCCUST,DCDATE,DCCHECKNO,DCINVNO) " _
      'x            & "VALUES('" & Trim(sJournalID) & "'," _
      'x            & lTrans & "," _
      'x            & iRef & "," _
      'x            & cAdvAmt & ",'" _
      'x            & sCrCashAcct & "','" _
      'x            & sCust & "','" _
      'x            & sRDate & "','" _
      'x            & sCheck & "'," _
      'x            & lNewInv & ")"
      '        clsAdoCon.ExecuteSQL sSQL
      gl.InvoiceNumber = lNewInv
      gl.AddDebitCredit CCur(cAdvAmt), CCur(0), sCrCashAcct, "", 0, 0, "", sCust, sCheck
      
      ' Credit AdvAct
      gl.AddDebitCredit 0, CCur(cAdvAmt), Compress(cmbAdvAct), "", 0, 0, "", sCust, sCheck
      
      gl.AddCashReceipt sCheck, sCust, bCaType, sRDate, TotalReceipt, unappliedAmt, 0, 0, 0, 0, sCDate, sCrCashAcct
   
   ' if unapplied amount < 0 create a credit memo for the amount
   ' IMAINC REQUESTED THIS IN JUNE 2018 TO APPLY SHORT PAYMENTS TO SALES REFUNDS
   ElseIf unappliedAmt < 0 Then
      RefundAvailable = -unappliedAmt
      lNewInv = fGetNextInvoice
      sSql = "INSERT INTO CihdTable (INVNO,INVTYPE,INVTOTAL,INVPAY,INVCUST," _
             & "INVDATE,INVSHIPDATE,INVORIGIN) " _
             & "VALUES(" & lNewInv & ",'CM'," _
             & unappliedAmt & "," & unappliedAmt & ",'" _
             & Compress(cmbCst) & "','" _
             & txtDte & "','" & txtDte _
             & "','" & sCheck & "')"
      clsADOCon.ExecuteSql sSql
      
      Dim inv2 As New ClassARInvoice
      inv2.SaveLastInvoiceNumber lNewInv
         
'      glSJ.InvoiceNumber = lNewInv
'
'      ' credit AR in SJ for amount of shortage. (Trade receivables 11100.000 for IMAINC)
'      glSJ.AddDebitCredit 0, -CCur(cAdvAmt), sSJARAcct, "", 0, 0, "", sCust, sCheck
'
'      ' debit sales in SJ(returns account 401500.000 for IMAINC)
'      glSJ.AddDebitCredit -CCur(cAdvAmt), 0, sAdvAct, "", 0, 0, "", sCust, sCheck
'
      'add zero dollar cash receipt.  report joins require this
      gl.AddCashReceipt sCheck, sCust, bCaType, sRDate, TotalReceipt, unappliedAmt, 0, 0, 0, 0, sCDate, sCrCashAcct
   
   End If
   
   
   ''''''''''''''''''''''''''''''''''''''''''
   
   
   'TRANSACTION FEE
   Dim feeAmt As Currency
   feeAmt = CCur("0" & txtFee)
   If feeAmt > 0 Then
      ' Debit Transaction Fee Account
      gl.InvoiceNumber = 0
      gl.AddDebitCredit feeAmt, 0, sTransFeeAcct, "", 0, 0, "", sCust, sCheck
      
      ' Credit Cash
      gl.AddDebitCredit 0, feeAmt, sCrCashAcct, "", 0, 0, "", sCust, sCheck
      
   End If
   
   
   ''''''''''''''''''''''''''''''''''''''''''
   
   'INVOICES
   Dim pif As Integer '=1 if item is paid in full
   For i = 1 To iTotalItems
      If vInvoice(i, INV_PAID) = vbChecked Then
         pif = vInvoice(i, INV_PAIDINFULL)
         gl.InvoiceNumber = lInvoices(i)
         
         'calculate amount applied.  if pif, amount = inv total
         'if pif, also calculate discount
         Dim amountPaid As Currency
         Dim debitCash As Currency
         Dim creditAR As Currency
         Dim creditOther As Currency
         
         Dim discountAmount As Currency
         Dim commissionAmount As Currency
         Dim revenueAmount As Currency
         Dim iAdjst As Integer
         iAdjst = 0
         
         debitCash = CCur(vInvoice(i, INV_AMOUNTPAID))
         If pif = 1 Then
            amountPaid = CCur(vInvoice(i, INV_TOTALAMOUNT))
            creditAR = CCur(vInvoice(i, INV_TOTALDUE))
            creditOther = debitCash - creditAR
            If creditOther <> 0 Then iAdjst = 1
            '                If Val(vInvoice(i, INV_ACCOUNT)) = INV_ACCOUNT_DISCOUNT Then
            '                    discountAmount = creditOther
            '                End If
         Else
            'amount paid so far
            amountPaid = (CCur(vInvoice(i, INV_TOTALAMOUNT)) - CCur(vInvoice(i, INV_TOTALDUE))) _
                         + debitCash
            creditAR = debitCash
            creditOther = 0
            '                discountAmount = 0
         End If
         
         ' create miscellaneous debit or credit if required
         Dim expenseAmount As Currency
         
         If creditOther <> 0 And Val(vInvoice(i, INV_ACCOUNT)) <> -1 Then
            Dim acct As String
            Select Case Val(vInvoice(i, INV_ACCOUNT))
               Case INV_ACCOUNT_DISCOUNT
                  acct = sCrDiscAcct
                  discountAmount = creditOther
               Case INV_ACCOUNT_EXPENSE
                  acct = sCrExpAcct
                  expenseAmount = creditOther
               Case INV_ACCOUNT_COMMISSION
                  acct = sCrCommAcct
                  commissionAmount = creditOther
               Case INV_ACCOUNT_REVENUE
                  acct = sCrRevAcct
                  revenueAmount = creditOther
            End Select
            
            'create debit or credit
            'note: AddDebitCredit turns a - credit into a debit
            If Len(acct) > 0 Then
               gl.AddDebitCredit 0, creditOther, acct, "", 0, 0, "", sCust, sCheck
            End If
         End If
         
         sSql = "UPDATE CihdTable SET " _
                & "INVCHECKNO='" & sCheck & "'," _
                & "INVPAY=" & amountPaid & "," _
                & "INVPIF=" & pif & "," _
                & "INVADJUST=" & iAdjst & "," _
                & "INVARDISC=" & CSng(vInvoice(i, INV_DISCOUNTPERCENT)) & "," _
                & "INVDAYS=" & Val(vInvoice(i, 10)) & "," _
                & "INVCHECKDATE='" & txtDte & "'  " _
                & "WHERE INVNO=" & lInvoices(i) & " "
         clsADOCon.ExecuteSql sSql
         
         ' Add a cash receipt record
         gl.AddCashReceipt sCheck, sCust, bCaType, sRDate, TotalReceipt, _
            creditAR, discountAmount, commissionAmount, expenseAmount, revenueAmount, sCDate, sCrCashAcct
         
         ' Check if invoice is type CA (cash advance), if so, credit deposit liability
         ' and debit the cash account that was selected
         If vInvoice(i, 13) = "CA" Then
            sSql = "SELECT DCACCTNO, ISNULL(DCDEBIT,0) as DCDEBIT, ISNULL(DCCREDIT,0) AS DCCREDIT FROM JritTable WHERE " _
                   & "DCINVNO = " & vInvoice(i, 1) & " ORDER BY DCDATE,DCTRAN,DCREF"
            bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_KEYSET)
            If bSqlRows Then
               
               '                    With rdoAct
               '                        ' Debit Advpay (note negative amount will be reversed in AddDebitCredit
               '                        gl.InvoiceNumber = vInvoice(i, 1)
               '                        gl.AddDebitCredit 0, debitCash, Trim(.Fields(0)), _
               '                            "", 0, 0, "", sCust, sCheck
               '                        .MoveNext
               '
               '                        ' Credit Cash
               '                        gl.AddDebitCredit debitCash, 0, Trim(.Fields(0)), _
               '                            "", 0, 0, "", sCust, sCheck
               '                        .Cancel
               '                    End With
               
               gl.InvoiceNumber = vInvoice(i, 1)
               With rdoAct
                  Do While Not .EOF
                     Dim priorDebit As Currency
                     priorDebit = !DCDEBIT
                     
                     ' Debit Advpay liability (note negative amount will be reversed in AddDebitCredit
                     If priorDebit = 0 Then
                        gl.AddDebitCredit 0, debitCash, Trim(.Fields(0)), _
                                                             "", 0, 0, "", sCust, sCheck
                        '.Cancel
                        Exit Do
                     Else
                        .MoveNext
                     End If
                  Loop
                  .Cancel
               End With
               
               ' Credit Cash
               gl.AddDebitCredit debitCash, 0, sCrCashAcct, _
                  "", 0, 0, "", sCust, sCheck
            End If
            Set rdoAct = Nothing
            
            'invoice other than cash advance
            'if credit memo, debit AR and credit cash
            '(negative amount will make this happen automatically)
         Else
            '                If vInvoice(i, 13) = "CM" Then
            '                    gl.AddDebitCredit creditAR, 0, sSJARAcct, "", 0, 0, "", sCust, sCheck
            '                    gl.AddDebitCredit 0, debitCash, sCrCashAcct, _
            '                        "", 0, 0, "", sCust, sCheck
            '
            '                'if debit memo or regular invoice, credit AR and debit cash
            '                Else
            
            gl.AddDebitCredit 0, creditAR, sSJARAcct, "", 0, 0, "", sCust, sCheck
            
            ' create short payment debit(s) to cash and/or refund
            Dim debitSales As Currency
            debitSales = 0
            If unappliedAmt < 0 Then
               If debitCash > CashAvailable Then
                  debitCash = CashAvailable
                  'CashAvailable = CashAvailable - debitCash
                  If debitCash < creditAR Then
                     debitSales = creditAR - debitCash
                     If debitSales > RefundAvailable Then
                        debitSales = RefundAvailable
                     End If
                     'RefundAvailable = RefundAvailable - debitSales
                  End If
               End If
               
               'short payment debit to cash until cash is used up
               If debitCash > 0 Then
                  gl.AddDebitCredit debitCash, 0, sCrCashAcct, "", 0, 0, "", sCust, sCheck
                  CashAvailable = CashAvailable - debitCash
               End If
               
               'short payment debit to sales refund after cash is used up
               If debitSales > 0 Then
                  gl.AddDebitCredit debitSales, 0, sAdvAct, "", 0, 0, "", sCust, sCheck
                  RefundAvailable = RefundAvailable - debitSales
               End If
            
            'normal non-short payment debit to cash
            Else
               gl.AddDebitCredit debitCash, 0, sCrCashAcct, "", 0, 0, "", sCust, sCheck
            End If
         End If
      End If
   Next
   
   'glSJ.ShowDebitsAndCredits
   'gl.ShowDebitsAndCredits
   
   Dim success As Boolean
   
   success = gl.Commit
'   If success And unappliedAmt < 0 Then
'      success = glSJ.Commit
'   End If
   
   Dim bDockLok As Boolean
   Dim bResponse As Byte
   bResponse = 0
   
   If success And clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      If lNewInv > 0 Then
         If Me.chkShortPayment.Value = vbChecked Then
            sMsg = "Cash Receipt Successfully Posted And" & vbCrLf _
                   & "Short payment credit memo invoice " & Format(lNewInv, "000000") & " Created." & vbCrLf _
                   & "Revised Invoice Comments Now?"
         Else
            sMsg = "Cash Receipt Successfully Posted And" & vbCrLf _
                   & "Advanced Payment Invoice " & Format(lNewInv, "000000") & " Created." & vbCrLf _
                   & "Revised Invoice Comments Now?"
         End If
         bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      Else
         SysMsg "Cash Receipt Successfully Posted", 1
      End If
          
      If (cmdDockLok.enabled = True) And (cmdDockLok.Visible = True) Then
         bDockLok = True
         sMsg = "Do you want to scan a Cash Receipt ?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            cmdDockLok_Click
         End If
      End If
          
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      sMsg = "Cash Receipt Not Successfully Posted." _
             & vbCrLf & "Transaction Canceled."
      MsgBox sMsg, vbExclamation, Caption
   End If
   
   ' Clear grid and set focus back to customer turn buttons back on.
   sCloseBoxes
   ClearArrays
   sNow = Format(ES_SYSDATE, "mm/dd/yy")
   txtDte.enabled = True
   txtDte = sNow
   txtCte.enabled = True
   txtCte = sNow
   txtChk.enabled = True
   txtChk = ""
   txtAmt.enabled = True
   txtAmt = "0.00"
   txtFee.enabled = True
   txtFee = "0.00"
   cmdSel.enabled = True
   cmbCst.enabled = True
   optWir.enabled = True
   optCsh.enabled = True
   optChk.enabled = True
   cmbAct.enabled = True
   chkAdv.Value = vbUnchecked
   txtAdvAmt = "0.00"
   cmbCst.SetFocus
   If ((bResponse = vbYes) And (bDockLok = False)) Then
      Load diaARe05a
      diaARe05a.Move Me.Left + 400, Me.Top + 400
      diaARe05a.chkCalledByCR = vbChecked
      diaARe05a.Show
      Sleep 500
      SendKeys Format(lNewInv, "000000") & "{TAB}"
   End If
   bDockLok = False
   ignoreClicks = False
   sCloseBoxes
   Exit Sub
DiaErr1:
   ignoreClicks = False
   sProcName = "sPostCheck"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

'Local Errors - from FillCombo
'Flash with 1 = no accounts nec, 2 = OK, 3 = Not enough accounts

Public Function fGetCashAccounts() As Byte
   Dim rdoCsh As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT COGLVERIFY,COCRCASHACCT,COCRDISCACCT,COSJARACCT," _
          & "COCRCOMMACCT,COCRREVACCT,COCREXPACCT,COTRANSFEEACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCsh, ES_FORWARD)
   sProcName = "getcashacct"
   If bSqlRows Then
      With rdoCsh
         For i = 1 To 7
            If "" & Trim(.Fields(i)) = "" Then
               b = 1
               Exit For
            End If
         Next
         sCrCashAcct = "" & Trim(!COCRCASHACCT)
         sCrDiscAcct = "" & Trim(!COCRDISCACCT)
         sSJARAcct = "" & Trim(!COSJARACCT)
         sCrCommAcct = "" & Trim(!COCRCOMMACCT)
         sCrRevAcct = "" & Trim(!COCRREVACCT)
         sCrExpAcct = "" & Trim(!COCREXPACCT)
         sTransFeeAcct = "" & Trim(!COTRANSFEEACCT)
         .Cancel
         If b = 1 Then fGetCashAccounts = 3 Else fGetCashAccounts = 2
      End With
   Else
      fGetCashAccounts = 0
   End If
   Set rdoCsh = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "fgetcashacco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function fGetNextInvoice() As Long
'   Dim RdoInv As ADODB.RecordSet
'   On Error GoTo DiaErr1
'   sSql = "SELECT MAX(INVNO) FROM CihdTable WHERE INVNO<999999"
'   bSqlRows = clsAdoCon.GetDataSet(sSql,RdoInv)
'   If bSqlRows Then
'      If IsNull(RdoInv.Fields(0)) Then
'         fGetNextInvoice = 1
'      Else
'         fGetNextInvoice = (RdoInv.Fields(0) + 1)
'      End If
'   Else
'      fGetNextInvoice = 1
'   End If
'   Set RdoInv = Nothing
'   Exit Function
'DiaErr1:
'   sProcName = "getnextinv"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me


   Dim inv As New ClassARInvoice
   fGetNextInvoice = inv.GetNextInvoiceNumber

End Function

Private Function fFindMatchedAmounts() As Integer
   ' Look for invoice matching check amount.
   Dim i As Integer
   Dim K As Integer
   Dim bMatched As Byte
   On Error Resume Next
   For i = 0 To UBound(vInvoice)
      If CCur(vInvoice(i, INV_TOTALDUE)) = cAmount Then
         vInvoice(i, INV_PAID) = 1
         bMatched = True
         fFindMatchedAmounts = i
         Exit For
      End If
   Next
End Function

'Private Function fSumSelectedItems() As Currency
'    Dim i           As Integer
'    Dim cTotal      As Currency
'    Dim cLessAdv    As Currency
'
'    On Error Resume Next
'    For i = 1 To iTotalItems
'        If vInvoice(i, INV_PAID) = vbChecked Then
'            If vInvoice(i, INV_ACCOUNT) > 1 Then
'                cLessAdv = cLessAdv + (vInvoice(i, INV_TOTALDUE) - vInvoice(i, INV_AMOUNTPAID))
'            End If
'            cTotal = cTotal + vInvoice(i, INV_AMOUNTPAID)
'        End If
'    Next
'
'    fSumSelectedItems = cTotal
'End Function
'

Private Function CashApplied() As Currency
   Dim i As Integer
   Dim cTotal As Currency
   Dim cLessAdv As Currency
   
   'sum invoices
   'On Error Resume Next
       For i = 1 To iTotalItems
      If vInvoice(i, INV_PAID) = vbChecked Then
         If vInvoice(i, INV_ACCOUNT) > 1 Then
            cLessAdv = cLessAdv + (vInvoice(i, INV_TOTALDUE) - vInvoice(i, INV_AMOUNTPAID))
         End If
         '            If vInvoice(i, INV_AMOUNTPAID) <> 0 Then
         '                MsgBox vInvoice(i, INV_AMOUNTPAID)
         '            End If
         cTotal = cTotal + vInvoice(i, INV_AMOUNTPAID)
      End If
   Next
   
   'now add unapplied cash
   cTotal = cTotal + CCur(txtAdvAmt)
   
   CashApplied = cTotal
End Function


Private Function fCheckExists(sCheck As String) As Byte
   Dim RdoChk As ADODB.Recordset
   
   On Error GoTo DiaErr1
   fCheckExists = False
   sSql = "SELECT COUNT(CACHECKNO) FROM CashTable WHERE CACHECKNO = '" _
          & sCheck & "' AND CACUST = '" & Compress(cmbCst) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      With RdoChk
         If .Fields(0) > 0 Then
            fCheckExists = True
            sMsg = "Check # " & sCheck & " All Ready Exists " & vbCrLf _
                   & "For Customer " & cmbCst
            MsgBox sMsg, vbInformation, Caption
         End If
         .Cancel
      End With
   End If
   Set RdoChk = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "fcheckexists"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Function

Sub ShowAccounts(Index As Integer)
   
   '    If inShowAccounts Then
   '        Exit Sub
   '    End If
   '    inShowAccounts = True
   Dim oldIgnoreClicks As Boolean
   oldIgnoreClicks = ignoreClicks
   ignoreClicks = True
   
   Dim enabled As Boolean
   enabled = False
   chkPIF(Index).enabled = True
   
   If chkPaid(Index).Value > 0 Then
      
      Dim due As Currency, applied As Currency
      due = CCur(lblDue(Index))
      applied = CCur(txtPaid(Index))
      
      If due < 0 Then
         Select Case due - applied
            Case Is = 0
               chkPIF(Index).Value = vbChecked 'if exact amt, must be pd in full
               chkPIF(Index).enabled = False
            Case Is > 0
            Case Is < 0
         End Select
         enabled = False
         
      Else
         
         'process normal invoice
         Select Case due - applied
            
            'exact payment
            Case Is = 0
               enabled = False
               chkPIF(Index).Value = vbChecked 'if exact amt, must be pd in full
               chkPIF(Index).enabled = False
               
               'underpayment
            Case Is > 0
               If chkPIF(Index).Value > 0 Then
                  enabled = True
               End If
               
               'overpayment
            Case Is < 0
               enabled = True
               chkPIF(Index).Value = vbChecked
               chkPIF(Index).enabled = False
         End Select
      End If
   Else
      chkPIF(Index).Value = vbUnchecked
      chkPIF(Index).enabled = False
   End If
   
   If enabled Then
      cmbapl(Index).ListIndex = 0
   Else
      cmbapl(Index).ListIndex = -1
   End If
   cmbapl(Index).enabled = enabled
   
   'update array
   vInvoice(iIndex + Index, INV_PAIDINFULL) = chkPIF(Index).Value
   If iIndex > 9 Then
      'MsgBox "here"
   End If
   vInvoice(iIndex + Index, INV_ACCOUNT) = cmbapl(Index).ListIndex
   
   'inShowAccounts = False
   
   ignoreClicks = oldIgnoreClicks
End Sub


Private Function DocumentLokEnabled() As Boolean
    Dim rdoDocLok As ADODB.Recordset
    Dim iDL As Integer
    
    DocumentLokEnabled = False
    sSql = "SELECT CODOCLOK FROM ComnTable WHERE COREF=1"
    bSqlRows = clsADOCon.GetDataSet(sSql, rdoDocLok, ES_FORWARD)
    If bSqlRows Then
        iDL = 0 & rdoDocLok!CODOCLOK
        If iDL = 1 Then DocumentLokEnabled = True
    End If
    Set rdoDocLok = Nothing
End Function

