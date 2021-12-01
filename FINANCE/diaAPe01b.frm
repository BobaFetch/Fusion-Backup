VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaAPe01b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Invoice Items (Purchase Order)"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDockLok 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6840
      Picture         =   "diaAPe01b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   105
      ToolTipText     =   "Push Info to M-Files"
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   5
      Left            =   7560
      TabIndex        =   92
      Tag             =   "2"
      ToolTipText     =   "Select Run From List"
      Top             =   6000
      Width           =   975
   End
   Begin VB.ComboBox cmbMon 
      Height          =   315
      Index           =   5
      Left            =   4320
      TabIndex        =   91
      Tag             =   "3"
      ToolTipText     =   "Blank, Enter Run or Select "
      Top             =   6000
      Width           =   2775
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   4
      Left            =   7560
      TabIndex        =   90
      Tag             =   "2"
      ToolTipText     =   "Select Run From List"
      Top             =   5160
      Width           =   975
   End
   Begin VB.ComboBox cmbMon 
      Height          =   315
      Index           =   4
      Left            =   4320
      TabIndex        =   89
      Tag             =   "3"
      ToolTipText     =   "Blank, Enter Run or Select "
      Top             =   5160
      Width           =   2775
   End
   Begin VB.ComboBox cmbMon 
      Height          =   315
      Index           =   3
      Left            =   4320
      TabIndex        =   22
      Tag             =   "3"
      ToolTipText     =   "Blank, Enter Run or Select "
      Top             =   4320
      Width           =   2775
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   3
      Left            =   7560
      TabIndex        =   23
      Tag             =   "2"
      ToolTipText     =   "Select Run From List"
      Top             =   4320
      Width           =   975
   End
   Begin VB.ComboBox cmbMon 
      Height          =   315
      Index           =   2
      Left            =   4320
      TabIndex        =   16
      Tag             =   "3"
      ToolTipText     =   "Blank, Enter Run or Select "
      Top             =   3480
      Width           =   2775
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   2
      Left            =   7560
      TabIndex        =   17
      Tag             =   "2"
      ToolTipText     =   "Select Run From List"
      Top             =   3480
      Width           =   975
   End
   Begin VB.ComboBox cmbMon 
      Height          =   315
      Index           =   1
      Left            =   4320
      TabIndex        =   10
      Tag             =   "3"
      ToolTipText     =   "Blank, Enter Run or Select "
      Top             =   2640
      Width           =   2775
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   1
      Left            =   7560
      TabIndex        =   11
      Tag             =   "2"
      ToolTipText     =   "Select Run From List"
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Index           =   5
      Left            =   765
      TabIndex        =   82
      Top             =   6000
      Width           =   3045
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Index           =   4
      Left            =   765
      TabIndex        =   21
      Top             =   5160
      Width           =   3045
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Index           =   3
      Left            =   765
      TabIndex        =   15
      Top             =   4320
      Width           =   3045
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Index           =   2
      Left            =   765
      TabIndex        =   9
      Top             =   3480
      Width           =   3045
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Index           =   1
      Left            =   765
      TabIndex        =   5
      Top             =   2640
      Width           =   3045
   End
   Begin VB.TextBox txtAdd 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   6540
      TabIndex        =   25
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox txtAdd 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   6540
      TabIndex        =   19
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtAdd 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   6540
      TabIndex        =   13
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtAdd 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   6540
      TabIndex        =   7
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtAdd 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   6540
      TabIndex        =   3
      Tag             =   "1"
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdPO 
      Height          =   320
      Index           =   5
      Left            =   120
      Picture         =   "diaAPe01b.frx":05DA
      Style           =   1  'Graphical
      TabIndex        =   77
      TabStop         =   0   'False
      ToolTipText     =   "Revise PO Item Comments"
      Top             =   6000
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CommandButton cmdPO 
      Height          =   320
      Index           =   4
      Left            =   120
      Picture         =   "diaAPe01b.frx":0B5C
      Style           =   1  'Graphical
      TabIndex        =   76
      TabStop         =   0   'False
      ToolTipText     =   "Revise PO Item Comments"
      Top             =   5160
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CommandButton cmdPO 
      Height          =   320
      Index           =   3
      Left            =   120
      Picture         =   "diaAPe01b.frx":10DE
      Style           =   1  'Graphical
      TabIndex        =   75
      TabStop         =   0   'False
      ToolTipText     =   "Revise PO Item Comments"
      Top             =   4320
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CommandButton cmdPO 
      Height          =   320
      Index           =   2
      Left            =   120
      Picture         =   "diaAPe01b.frx":1660
      Style           =   1  'Graphical
      TabIndex        =   74
      TabStop         =   0   'False
      ToolTipText     =   "Revise PO Item Comments"
      Top             =   3480
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CommandButton cmdPO 
      Height          =   320
      Index           =   1
      Left            =   120
      Picture         =   "diaAPe01b.frx":1BE2
      Style           =   1  'Graphical
      TabIndex        =   73
      TabStop         =   0   'False
      ToolTipText     =   "Show PO Item Comments"
      Top             =   2640
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   1
      Left            =   120
      TabIndex        =   72
      Top             =   1920
      Width           =   8415
   End
   Begin VB.CommandButton cmdItemNoPo 
      Caption         =   "&GL"
      Height          =   315
      Left            =   7680
      TabIndex        =   71
      ToolTipText     =   "Add GL Account Distribution"
      Top             =   1080
      Width           =   875
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Update"
      Height          =   315
      Left            =   7680
      TabIndex        =   69
      TabStop         =   0   'False
      ToolTipText     =   "Update Invoice Amount To Total Selected"
      Top             =   1440
      Width           =   875
   End
   Begin VB.TextBox txtCst 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   5460
      TabIndex        =   24
      Tag             =   "1"
      ToolTipText     =   "Unit Cost"
      Top             =   5640
      Width           =   1020
   End
   Begin VB.TextBox txtCst 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   5460
      TabIndex        =   18
      Tag             =   "1"
      ToolTipText     =   "Unit Cost"
      Top             =   4800
      Width           =   1020
   End
   Begin VB.TextBox txtCst 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   5460
      TabIndex        =   12
      Tag             =   "1"
      ToolTipText     =   "Unit Cost"
      Top             =   3960
      Width           =   1020
   End
   Begin VB.TextBox txtCst 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   5460
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Unit Cost"
      Top             =   3120
      Width           =   1020
   End
   Begin VB.TextBox txtCst 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   5460
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "Unit Cost"
      Top             =   2280
      Width           =   1020
   End
   Begin VB.CheckBox optPst 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   8580
      TabIndex        =   26
      Top             =   5640
      Width           =   585
   End
   Begin VB.TextBox txtFrt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4980
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Freight and Misc"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtTax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4980
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Sales Tax For Invoice"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox optPst 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   8580
      TabIndex        =   20
      Top             =   4800
      Width           =   600
   End
   Begin VB.CheckBox optPst 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   8580
      TabIndex        =   14
      Top             =   3960
      Width           =   600
   End
   Begin VB.CheckBox optPst 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   8580
      TabIndex        =   8
      Top             =   3120
      Width           =   600
   End
   Begin VB.CheckBox optPst 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   8580
      TabIndex        =   4
      ToolTipText     =   "Invoice This Item"
      Top             =   2280
      Width           =   600
   End
   Begin VB.CommandButton cmdPst 
      Caption         =   "&Post"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7680
      TabIndex        =   39
      ToolTipText     =   "Save and Post This Item"
      Top             =   720
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7680
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   28
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
      PictureUp       =   "diaAPe01b.frx":2164
      PictureDn       =   "diaAPe01b.frx":22AA
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8760
      Top             =   6360
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7035
      FormDesignWidth =   9435
   End
   Begin Threed.SSCommand cmdDn 
      Height          =   375
      Left            =   8160
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   "Next Page (Page Down)"
      Top             =   6480
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaAPe01b.frx":23F0
   End
   Begin Threed.SSCommand cmdUp 
      Height          =   375
      Left            =   7800
      TabIndex        =   61
      TabStop         =   0   'False
      ToolTipText     =   "Last Page (Page Up)"
      Top             =   6480
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaAPe01b.frx":28F2
   End
   Begin VB.Label lblPostDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4980
      TabIndex        =   104
      Tag             =   "1"
      ToolTipText     =   "Amount Of Invoice"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Date"
      Height          =   255
      Index           =   16
      Left            =   4020
      TabIndex        =   103
      Top             =   120
      Width           =   810
   End
   Begin VB.Label lblRun 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   5
      Left            =   7200
      TabIndex        =   102
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblRun 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   4
      Left            =   7200
      TabIndex        =   101
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblRun 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   3
      Left            =   7200
      TabIndex        =   100
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblRun 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   2
      Left            =   7200
      TabIndex        =   99
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMon 
      BackStyle       =   0  'Transparent
      Caption         =   "MO"
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   98
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMon 
      BackStyle       =   0  'Transparent
      Caption         =   "MO"
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   97
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMon 
      BackStyle       =   0  'Transparent
      Caption         =   "MO"
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   96
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMon 
      BackStyle       =   0  'Transparent
      Caption         =   "MO"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   95
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblRun 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   94
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMon 
      BackStyle       =   0  'Transparent
      Caption         =   "MO"
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   93
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblExt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   7500
      TabIndex        =   88
      Top             =   5640
      Width           =   1020
   End
   Begin VB.Label lblExt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   7500
      TabIndex        =   87
      Top             =   4800
      Width           =   1020
   End
   Begin VB.Label lblExt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   7500
      TabIndex        =   86
      Top             =   3960
      Width           =   1020
   End
   Begin VB.Label lblExt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   7500
      TabIndex        =   85
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Label lblExt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   7500
      TabIndex        =   84
      Top             =   2280
      Width           =   1020
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   7500
      TabIndex        =   83
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   81
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO #"
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   80
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   6540
      TabIndex        =   79
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost               "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   5460
      TabIndex        =   78
      Top             =   2040
      Width           =   1020
   End
   Begin VB.Label lblDue 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   70
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include             "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   8580
      TabIndex        =   68
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblRel 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   67
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   66
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      Height          =   255
      Index           =   10
      Left            =   6960
      TabIndex        =   65
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   255
      Index           =   9
      Left            =   5880
      TabIndex        =   64
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label lblLst 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   63
      Top             =   6600
      Width           =   375
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
      Left            =   6480
      TabIndex        =   62
      Top             =   6600
      Width           =   375
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   2400
      Picture         =   "diaAPe01b.frx":2DF4
      Top             =   6480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   2160
      Picture         =   "diaAPe01b.frx":32E6
      Top             =   6480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   1920
      Picture         =   "diaAPe01b.frx":37D8
      Top             =   6480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   2640
      Picture         =   "diaAPe01b.frx":3CCA
      Top             =   6480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   59
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   765
      TabIndex        =   58
      Top             =   5640
      Width           =   3045
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   3960
      TabIndex        =   57
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight"
      Height          =   255
      Index           =   8
      Left            =   4020
      TabIndex        =   56
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax"
      Height          =   255
      Index           =   7
      Left            =   4020
      TabIndex        =   55
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   3960
      TabIndex        =   54
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   765
      TabIndex        =   53
      Top             =   4800
      Width           =   3045
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   52
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   3960
      TabIndex        =   51
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   765
      TabIndex        =   50
      Top             =   3960
      Width           =   3045
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   49
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   3960
      TabIndex        =   48
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   765
      TabIndex        =   47
      Top             =   3120
      Width           =   3045
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   46
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   3960
      TabIndex        =   45
      ToolTipText     =   "Quantity"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   765
      TabIndex        =   44
      Top             =   2280
      Width           =   3045
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   43
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity / Lot Qty                  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3960
      TabIndex        =   42
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number/Description                              "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   765
      TabIndex        =   41
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   40
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblIdt 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblPdt 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   36
      Top             =   120
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   35
      Top             =   480
      Width           =   690
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   34
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   33
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   255
      Index           =   2
      Left            =   4020
      TabIndex        =   32
      Top             =   480
      Width           =   810
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected"
      Height          =   255
      Index           =   3
      Left            =   4020
      TabIndex        =   31
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4980
      TabIndex        =   30
      Tag             =   "1"
      ToolTipText     =   "Amount Of Invoice"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4980
      TabIndex        =   29
      Tag             =   "1"
      ToolTipText     =   "Accumulated Total"
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "diaAPe01b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*************************************************************************************
'
' diaAPe01b - Add PO items and GL distributions to invoice.
'
' Notes:
'
' Created: (cjs)
' Revisions:
' 08/02/01 (nth) Add the abliblity to add items not found on PO (GL Dis).
' 08/07/01 (nth) Fixed error, not showing even if no items not on PO where present.
' 11/27/02 (nth) Plug lot cost of PO item into lot header.
' 12/10/02 (nth) Changed adjustments to allow negative numbers.
' 12/26/02 (nth) Added ablility to revised GL distribution comments after adding per JLH
' 01/24/03 (nth) Fixed error, invoice with no PO and 6 or more GL distributions
'                does record any more than 5.
' 01/24/03 (nth) Added view PO item comments.
' 01/28/03 (nth) Fixed invoice Post, Invoice, and Due date.
' 11/24/03 (nth) Added extended per incident 18394.
' 02/19/04 (nth) Divide freight by number of PO items and add to lot cost.
' 05/20/04 (nth) Added MO run allocations to invoice line item per ANDELE.
' 11/08/04 (nth) Correct adding of extened totals.
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bPosted As Boolean
Dim bGoodRun As Boolean

Dim iIndex As Integer
Dim iCurrPage As Integer
Dim iTotalItems As Integer
Dim iPOItems As Integer

'Dim sAccount    As String
Dim sApAcct As String
Dim sFrAcct As String
Dim sTxAcct As String
Dim sMsg As String
Dim sJournalID As String
Dim cPrevAmt As Currency
Dim ADOErrNum As Double
Dim ADOErrDesc As String


' Five invoice items per page
Const NEXTPAGE = 5

' vInvoice Documentation
Dim vInvoice(300, 14) As Variant

' 0 = Item
Const VINV_ITEMNO = 0
' 1 = Item Rev
Const VINV_ITEMREV = 1
' 2 = Part Number
Const VINV_PARTNO = 2
' 3 = Part Desc
Const VINV_PARTDESC = 3
' 4 = PO Qty
Const VINV_QTY = 4
' 5 = PO Cost
Const VINV_POCOST = 5
' 6 = Invoiced Cost
Const VINV_INVOICEDCOST = 6
' 7 = Account
Const VINV_ACCOUNT = 7
' 8 = To Be Invoiced (0 or 1)
Const VINV_INCLUDE = 8
' 9 = Type PO / GL Distribution (no PO)
Const VINV_GLDISTRIBUTION = 9
' 10 = Adjustments
Const VINV_ADJUSTMENTS = 10
' 11 = Lot Number
Const VINV_LOTNO = 11
' 12 = MO
Const VINV_MO = 12
' 13 = Run
Const VINV_RUNNO = 13
' 13 = Run
Const VINV_LOTQTY = 14

Private Sub CheckIfBalanced()
   If (CCur(lblTot) + CCur(Trim(txtFrt)) + CCur(txtTax)) = CCur(lblAmt) Then
      cmdpst.enabled = True
      cmdDockLok.enabled = True
      cmdUpd.enabled = False
   Else
      cmdpst.enabled = False
      cmdDockLok.enabled = False
      cmdUpd.enabled = True
   End If
End Sub


Private Function GetDBAccounts(ApAcct As String, FrAcct As String, _
                               TxAcct As String) As Byte
   Dim RdoCdm As ADODB.Recordset
   Dim b As Byte
   Dim i As Integer
   
   On Error GoTo DiaErr1
   sSql = "SELECT COAPACCT,COPJTAXACCT,COPJTFRTACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCdm, ES_FORWARD)
   If bSqlRows Then
      With RdoCdm
         For i = 0 To 2
            If Not IsNull(.Fields(i)) Then
               If Trim(.Fields(i)) = "" Then b = 1
            Else
               b = 1
            End If
         Next
         ApAcct = "" & Trim(!COAPACCT)
         TxAcct = "" & Trim(!COPJTAXACCT)
         FrAcct = "" & Trim(!COPJTFRTACCT)
         .Cancel
      End With
      If b = 0 Then GetDBAccounts = 1
   Else
      ApAcct = ""
      FrAcct = ""
      TxAcct = ""
      GetDBAccounts = 0
   End If
   Set RdoCdm = Nothing
   Exit Function
DiaErr1:
   sProcName = "getdbaccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub cmbMon_Click(Index As Integer)
   bGoodRun = GetRuns(Index)
End Sub

Private Sub cmbMon_LostFocus(Index As Integer)
   cmbMon(Index) = CheckLen(cmbMon(Index), 30)
   If Len(cmbMon(Index)) Then
      bGoodRun = GetRuns(Index)
   Else
      cmbRun(Index).Clear
   End If
End Sub

Private Sub cmbRun_LostFocus(Index As Integer)
   vInvoice(Index + iIndex, VINV_MO) = Trim(cmbMon(Index))
   vInvoice(Index + iIndex, VINV_RUNNO) = Trim(cmbRun(Index))
End Sub

Private Sub cmdCan_Click()
   Unload Me
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


Private Sub cmdDockLok_Click()
   Dim mfileInt As MfileIntegrator
   Set mfileInt = New MfileIntegrator
    
   Dim strInvNum As String
   Dim strVendor As String
   Dim strPONum As String
    
   Dim strIniPath As String
   strIniPath = App.Path & "\" & "MFileInit.ini"
    
   strInvNum = GetSectionEntry("MFILES_METADATA", "INVNUM", strIniPath)
   strVendor = GetSectionEntry("MFILES_METADATA", "VENDOR", strIniPath)
   strPONum = GetSectionEntry("MFILES_METADATA", "PONUM", strIniPath)
    
   mfileInt.OpenXMLFile "Insert", "VendorInv", "Vendor Invoices", mfileInt.gsVndInvClassID, "Scan", lblInv
   mfileInt.AddXMLMetaData "Vendor", strVendor, lblVnd
   mfileInt.AddXMLMetaData "InvoiceNumber", strInvNum, lblInv
   mfileInt.AddXMLMetaData "PONumber", strPONum, lblPon
      
   mfileInt.CloseXMLFile
'   If mfileInt.SendXMLFileToMFile Then SysMsg "Record Successfully Indexed", True
   If Not mfileInt.SendXMLFileToMFile Then SysMsg "Record Indexing Failed", True

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

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Invoice Items (Single PO)"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdItemNoPo_Click()
   'UpdateTotals
   diaAPe01c.Show
   diaAPe01c.SetFocus
End Sub

Private Sub cmdPO_Click(Index As Integer)
   ViewItemComments CInt(vInvoice(iIndex + Index, VINV_ITEMNO)), _
                         CStr(vInvoice(iIndex + Index, VINV_ITEMREV))
End Sub

Private Sub cmdPst_Click()
   UpdateTotals
   PostInvoice
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
   lblAmt = Format(CCur(lblTot) + CCur(txtTax) + CCur(txtFrt), _
            "#,###,##0.00")
   
   If (Val(lblAmt) <> Val(cPrevAmt)) Then
      MsgBox "The initially Invoice Amount entered is not the same as the updated Invoice Amount." & vbCrLf _
                  & "The Initial entered Invoice Amount =" & Format(cPrevAmt, CURRENCYMASK) & ".", _
                  vbInformation, Caption
   End If
   
   diaAPe01a.txtAmt = lblAmt
   CheckIfBalanced
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      HideBoxes
      b = GetDBAccounts(sApAcct, sFrAcct, sTxAcct)
      If b = 0 Then
         MsgBox "One Or More Of The AP Accounts Required Is Not Installed." & vbCr _
            & "Please Install All Accounts In The AP Tab Of Company Setup." & vbCr _
            & "Also Install All Accounts For Parts, Product Codes And Types.", _
            vbInformation, Caption
         Unload Me
         Exit Sub
      End If
      FillMOs
      If Val(lblPon) = 0 Then
         diaAPe01c.Show
         DoEvents
      Else
         GetPoItems
      End If
      If DocumentLokEnabled Then cmdDockLok.Visible = True Else cmdDockLok.Visible = False
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_GotFocus()
   Unload Me
End Sub

Private Sub Form_Load()
   FormLoad Me
   Move diaAPe01a.Left + 200, diaAPe01a.Top + 200
   
   txtFrt = Format(0, "0.00")
   txtTax = Format(0, "0.00")
   lblTot = Format(0, "0.00")
   lblRel = "0"
   
   lblVnd = diaAPe01a.cmbVnd
   lblNme = diaAPe01a.lblNme
   lblInv = diaAPe01a.txtInv
   lblAmt = diaAPe01a.txtAmt
   
   cPrevAmt = diaAPe01a.txtAmt
   lblPdt = CheckDate(diaAPe01a.txtPdt) ' Invoice Post
   lblIdt = CheckDate(diaAPe01a.txtIdt) ' Invoice Received
   lblDue = CheckDate(diaAPe01a.txtDue) ' Invoice Due
   
   lblPon = diaAPe01a.cmbPon
   
   lblLst = "1"
   lblPge = "1"
   iCurrPage = 1
   
   sJournalID = diaAPe01a.sJournalID
   
   'sAccount = GetSetting("Esi2000", "Fina", "LastAccount", sAccount)
   bOnLoad = True
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bResponse As Byte
   
   diaAPe01a.enabled = True
   If bPosted Then diaAPe01a.optPst = vbChecked
   
   If Not bPosted And Val(lblTot) > 0 Then
      sMsg = "Do You Really Want To Close This " & vbCrLf _
             & "Invoice Without Posting It?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then Cancel = True
   End If
   'SaveSetting "Esi2000", "Fina", "LastAccount", sAccount
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set diaAPe01b = Nothing
End Sub

Private Sub UpdateTotals()
   Dim i As Integer
   Dim cTotal As Currency
   On Error GoTo DiaErr1
   For i = 1 To iTotalItems
      If CInt(vInvoice(i, VINV_INCLUDE)) = 1 Then
         cTotal = cTotal + CCur(vInvoice(i, VINV_INVOICEDCOST))
      End If
   Next
   lblTot = Format(cTotal, CURRENCYMASK)
   CheckIfBalanced
   Exit Sub
DiaErr1:
   sProcName = "updatetot"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtAdd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
End Sub

Private Sub txtAdd_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
End Sub

Private Sub txtDsc_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub optPst_Click(Index As Integer)
   If optPst(Index).Value = vbChecked Then
      'vInvoice(Index + iIndex, VINV_POCOST) = CCur((txtCst(Index)))
      vInvoice(Index + iIndex, VINV_INCLUDE) = 1
      'ComputeLabelExtension Index
   Else
      'lblExt(Index) = "0.00"
      'vInvoice(Index + iIndex, VINV_INVOICEDCOST) = 0
      vInvoice(Index + iIndex, VINV_INCLUDE) = 0
   End If
   UpdateTotals
End Sub

Private Sub optPst_KeyPress(Index As Integer, KeyAscii As Integer)
   'KeyLock KeyAscii
End Sub

Private Sub optPst_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
End Sub

Private Sub txtAdd_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtAdd_LostFocus(Index As Integer)
   txtAdd(Index) = CheckLen(txtAdd(Index), VINV_RUNNO)
   If txtAdd(Index) = "" Then txtAdd(Index) = "0.00"
   
   txtAdd(Index) = Format(txtAdd(Index), "#,###,##0.00##")
   
   If CCur(txtAdd(Index)) <> 0 Then
      vInvoice(Index + iIndex, VINV_ADJUSTMENTS) = CCur(txtAdd(Index))
   Else
      vInvoice(Index + iIndex, VINV_ADJUSTMENTS) = 0
   End If
   
   ComputeLabelExtension Index
   
   UpdateTotals
End Sub

Private Sub txtCst_GotFocus(Index As Integer)
   SelectFormat Me
End Sub


Private Sub txtCst_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
End Sub

Private Sub txtCst_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
End Sub

Private Sub txtCst_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
End Sub

Private Sub txtCst_LostFocus(Index As Integer)
   
   txtCst(Index) = CheckLen(txtCst(Index), VINV_RUNNO)
   If txtCst(Index) = "" Then
      txtCst(Index) = "0.0000"
   End If
   If CheckCurrency(txtCst(Index), False) = "*" Then
    txtCst(Index).SetFocus
    Exit Sub
   End If
   
   txtCst(Index) = Format(txtCst(Index), "#,###,##0.0000")
   
   vInvoice(Index + iIndex, VINV_POCOST) = txtCst(Index)
   ComputeLabelExtension Index
   
   UpdateTotals
End Sub

Private Sub txtDsc_LostFocus(Index As Integer)
   If Val(vInvoice(iIndex + Index, VINV_GLDISTRIBUTION)) = 1 Then
      txtDsc(Index) = CheckLen(txtDsc(Index), 30)
      vInvoice(iIndex + Index, VINV_PARTDESC) = txtDsc(Index)
   End If
End Sub

Private Sub txtFrt_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtFrt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
End Sub


Private Sub txtFrt_KeyPress(KeyAscii As Integer)
   KeyValue KeyAscii
End Sub

Private Sub txtFrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
End Sub

Private Sub txtFrt_LostFocus()
   txtFrt = CheckLen(txtFrt, 8)
   txtFrt = Format(Abs(Val(txtFrt)), "#,###,##0.00")
   UpdateTotals
End Sub

Private Sub txtTax_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtTax_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
End Sub

Private Sub txtTax_KeyPress(KeyAscii As Integer)
   KeyValue KeyAscii
End Sub

Private Sub GetPoItems()
   Dim RdoPoi As ADODB.Recordset
   Dim b As Byte
   Dim i As Integer
   Dim sInvAcct As String
   
   On Error GoTo DiaErr1
   
   ' 1/4/2019 determine whether to use PIACCOUNT for inventory acct if it exists
   Dim rs As ADODB.Recordset
   Dim usePIACCOUNTS As Boolean
   sSql = "select purchaseaccount from preferences where PreRecord = 1"
   If clsADOCon.GetDataSet(sSql, rs, ES_FORWARD) <> 0 Then
      usePIACCOUNTS = IIf(CInt("0" + rs.Fields(0)) = 0, False, True)
   End If
   Set rs = Nothing
   
   ' Added 11/20/2009 'AND PIAQTY <> 0' for all rejected partnumbers
   ' may be later we can add 'PIPQTY <> PIONDOCKQTYREJ'
   sSql = "SELECT PONUMBER,POVENDOR,PINUMBER,PIITEM,PIREV,PIRUNPART,PIRUNNO," _
          & "PIPART,PITYPE,PIAQTY,PIAMT, PIESTUNIT,PILOT,PILOTNUMBER,PARTREF,PARTNUM,PADESC," _
          & "PALEVEL,PAPRODCODE,RTRIM(PIACCOUNT) AS PIACCOUNT FROM PohdTable,PoitTable,PartTable WHERE PONUMBER=" _
          & "PINUMBER AND (PINUMBER=" & Val(lblPon) & " AND " _
          & "PITYPE=15) AND PIAQTY <> 0 AND PIPART=PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPoi, ES_FORWARD)
   If bSqlRows Then
      With RdoPoi
         Do Until .EOF
            i = i + 1
            vInvoice(i, VINV_ITEMNO) = !PIITEM
            vInvoice(i, VINV_ITEMREV) = Trim(!PIREV)
            vInvoice(i, VINV_PARTNO) = Trim(!PARTNUM)
            vInvoice(i, VINV_PARTDESC) = Trim(!PADESC)
            vInvoice(i, VINV_QTY) = Format(!PIAQTY, "####0.000")
            vInvoice(i, VINV_POCOST) = Format(!PIAMT, "#,###,##0.0000")
            'vInvoice(i, VINV_POCOST) = Format(!PIESTUNIT, "#,###,##0.0000")
            vInvoice(i, VINV_INVOICEDCOST) = 0
            vInvoice(i, VINV_INCLUDE) = 0
            vInvoice(i, VINV_GLDISTRIBUTION) = 0
            vInvoice(i, VINV_ADJUSTMENTS) = "0.00"
            vInvoice(i, 11) = "" & Trim(!PILOTNUMBER)
            If usePIACCOUNTS Then
               sInvAcct = !PIACCOUNT
            Else
               sInvAcct = ""
            End If
            If sInvAcct = "" Then
               b = GetInvAccount(Trim(!PARTNUM), !PALEVEL, !PAPRODCODE, sInvAcct)
            End If
            vInvoice(i, VINV_ACCOUNT) = sInvAcct
            
            vInvoice(i, VINV_MO) = "" & Trim(!PIRUNPART)
            vInvoice(i, VINV_RUNNO) = !PIRUNNO
            vInvoice(i, VINV_LOTQTY) = IIf(!PILOT = 0, 1, CCur(!PILOT))
            
            ComputeArrayExtension (i)
            
            .MoveNext
         Loop
         .Cancel
         iPOItems = i
      End With
   Else
      If Me.Caption <> "Vendor Invoice (No PO)" Then
         MsgBox "No Uninvoiced, Received Items Found.", _
            vbInformation, Caption
      End If
   End If
   
   If i > 0 Then
      iTotalItems = i
      
      lblLst = CInt((((iTotalItems * 100) / 100) / NEXTPAGE) + 0.5)
''      For i = 1 To 5
''         If i > iTotalItems Then Exit For
''         lblItm(i).Visible = True
''         lblPrt(i).Visible = True
''
''         txtDsc(i).Visible = True
''         txtDsc(i).BackColor = Me.BackColor
''         txtDsc(i).Locked = True
''
''         lblQty(i).Visible = True
''         txtCst(i).Visible = True
''         optPst(i).Visible = True
''         cmdPO(i).Visible = True
''         txtAdd(i).Visible = True
''         lblExt(i).Visible = True
''
''         cmbMon(i).Visible = True
''         cmbRun(i).Visible = True
''         lblRun(i).Visible = True
''         lblMon(i).Visible = True
''
''         lblItm(i) = vInvoice(i, VINV_ITEMNO) & vInvoice(i, VINV_ITEMREV)
''         lblPrt(i) = vInvoice(i, VINV_PARTNO)
''         txtDsc(i) = vInvoice(i, VINV_PARTDESC)
''         lblQty(i) = vInvoice(i, VINV_QTY)
''         txtCst(i) = vInvoice(i, VINV_POCOST)
''         txtAdd(i) = vInvoice(i, VINV_ADJUSTMENTS)
''
''         '                lblExt(I) = Format(vInvoice(I + iIndex, VINV_QTY) _
''         '                    * vInvoice(I + iIndex, VINV_POCOST) _
''         '                    + vInvoice(I + iIndex, VINV_ADJUSTMENTS), "#,###,##0.00#")
''         'ComputeLabelExtension i
''         lblExt(i) = vInvoice(i, VINV_INVOICEDCOST)
''         cmbMon(i) = vInvoice(i, VINV_MO)
''         cmbRun(i) = vInvoice(i, VINV_RUNNO)
''         If vInvoice(i, VINV_GLDISTRIBUTION) = 0 Then
''            cmbMon(i).enabled = False
''            cmbRun(i).enabled = False
''         End If
''
''
''      Next
      If Val(lblLst) > 1 Then
         cmdDn.enabled = True
         cmdDn.Picture = Endn
      End If
      iIndex = 0
   End If
   
   GetNextGroup
   
   Set RdoPoi = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpoitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub ComputeLabelExtension(nSubscript As Integer)
   '    Dim cAmount As Currency
   '    cAmount = Round(vInvoice(nSubscript + iIndex, VINV_QTY) _
   '        * vInvoice(nSubscript + iIndex, VINV_POCOST) _
   '        + vInvoice(nSubscript + iIndex, VINV_ADJUSTMENTS), VINV_PARTNO)
   '    lblExt(nSubscript) = Format(cAmount, "#,###,##0.00")
   '    vInvoice(nSubscript + iIndex, VINV_INVOICEDCOST) = cAmount
   ComputeArrayExtension nSubscript + iIndex
   lblExt(nSubscript) = vInvoice(nSubscript + iIndex, VINV_INVOICEDCOST)
End Sub

Sub ComputeArrayExtension(nSubscript As Integer)
   'nSubscript = label index
   Dim cAmount As Currency
   
'   cAmount = SARound(vInvoice(nSubscript, VINV_QTY) _
'      * vInvoice(nSubscript, VINV_POCOST) / vInvoice(nSubscript, VINV_LOTQTY) _
'      + vInvoice(nSubscript, VINV_ADJUSTMENTS), 2)

' 2/21/2010 - Removed the LOTQTY as the PIAMT/PIESTUNIT is acual unit cost now
   cAmount = SARound(vInvoice(nSubscript, VINV_QTY) _
      * vInvoice(nSubscript, VINV_POCOST) _
      + vInvoice(nSubscript, VINV_ADJUSTMENTS), 2)
   vInvoice(nSubscript, VINV_INVOICEDCOST) = Format(cAmount, "###,###,###.00")
End Sub


Private Sub txtTax_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
End Sub

Private Sub GetNextGroup()
   Dim i As Integer
   iIndex = (Val(lblPge) - 1) * 5
   HideBoxes
   On Error GoTo DiaErr1
   For i = 1 To 5
      If i + iIndex > iTotalItems Then Exit For
      
      lblPrt(i).Visible = True
      txtDsc(i).Visible = True
      
      
      
      If vInvoice(i + iIndex, VINV_GLDISTRIBUTION) = 1 Then
         ' Hide if GL account distribution.
         lblItm(i).Visible = False
         lblQty(i).Visible = False
         txtAdd(i).Visible = False
         txtDsc(i).Locked = False
         txtDsc(i).BackColor = txtTax.BackColor
         cmdPO(i).Visible = False
         
      Else
         lblItm(i).Visible = True
         lblQty(i).Visible = True
         txtAdd(i).Visible = True
         txtDsc(i).Locked = True
         txtDsc(i).BackColor = Me.BackColor
         cmdPO(i).Visible = True
      End If
      
      lblExt(i).Visible = True
      txtCst(i).Visible = True
      optPst(i).Visible = True
      cmbMon(i).Visible = True
      cmbRun(i).Visible = True
      lblRun(i).Visible = True
      lblMon(i).Visible = True
      
      lblItm(i) = vInvoice(i + iIndex, VINV_ITEMNO) & vInvoice(i + iIndex, VINV_ITEMREV)
      lblPrt(i) = vInvoice(i + iIndex, VINV_PARTNO)
      txtDsc(i) = vInvoice(i + iIndex, VINV_PARTDESC)
      lblQty(i) = vInvoice(i + iIndex, VINV_QTY)
      If vInvoice(i + iIndex, VINV_LOTQTY) <> 1 Then
         lblQty(i) = lblQty(i) & " / " & vInvoice(i + iIndex, VINV_LOTQTY)
      End If
      txtCst(i) = Format(vInvoice(i + iIndex, VINV_POCOST), "#,###,##0.0000")
      txtAdd(i) = Format(vInvoice(i + iIndex, VINV_ADJUSTMENTS), "#,###,##0.00#")
      optPst(i).Value = Val(vInvoice(i + iIndex, VINV_INCLUDE))
      
      '            lblExt(I) = Format(vInvoice(I + iIndex, VINV_QTY) _
      '                * vInvoice(I + iIndex, VINV_POCOST) _
      '                + vInvoice(I + iIndex, VINV_ADJUSTMENTS), "#,###,##0.00#")
      ComputeLabelExtension i
      
      cmbMon(i) = vInvoice(i + iIndex, VINV_MO)
      cmbRun(i) = vInvoice(i + iIndex, VINV_RUNNO)
      
      If vInvoice(i + iIndex, VINV_GLDISTRIBUTION) = 0 Then
         cmbMon(i).enabled = False
         cmbRun(i).enabled = False
      End If
   Next
   Exit Sub
   
DiaErr1:
   sProcName = "GetNextGroup"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub HideBoxes()
   Dim i As Integer
   For i = 1 To 5
      lblItm(i).Visible = False
      lblPrt(i).Visible = False
      txtDsc(i).Visible = False
      lblQty(i).Visible = False
      txtCst(i).Visible = False
      optPst(i).Visible = False
      cmdPO(i).Visible = False
      txtAdd(i).Visible = False
      lblExt(i).Visible = False
      cmbMon(i).Visible = False
      cmbRun(i).Visible = False
      lblRun(i).Visible = False
      lblMon(i).Visible = False
   Next
End Sub

Private Sub txtTax_LostFocus()
   txtTax = CheckLen(txtTax, 8)
   txtTax = Format(Abs(Val(txtTax)), "####0.00")
   UpdateTotals
End Sub

Private Sub PostInvoice()
   Dim bByte As Boolean
   Dim bResponse As Byte
   
   Dim i As Integer
   Dim K As Integer
   Dim iTrans As Integer
   Dim iRef As Integer
   
   Dim cAmt As Currency
   Dim cTot As Currency
   Dim cQty As Currency
   Dim cFREIGHT As Currency
   Dim cItemTot As Currency
   Dim cLineTot As Currency
   Dim cTax As Currency
   
   'Dim sPon      As String
   Dim sVendor As String
   Dim sAcctNo As String
   Dim sTxtCmt As String
   
   Dim sCredit As String
   Dim sDebit As String
   Dim sNow As String
   Dim sPostDate As String
   
   On Error GoTo DiaErr1
   
   cTot = CCur(lblAmt)
   cAmt = CCur(lblAmt)
   
   ' Make sure something is selected
   For i = 1 To iTotalItems
      If vInvoice(i, VINV_INCLUDE) = 1 Then bByte = True
   Next
   If Not bByte Then
      sMsg = "There Are No Items Selected To Be Invoiced."
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   
   'sNow = Format(ES_SYSDATE, "mm/dd/yy")
   sPostDate = lblPostDate
   
   If sJournalID <> "" Then
      iTrans = GetNextTransaction(sJournalID)
   End If
   
   ' Everybody have an account?
   bByte = False
   If iTrans > 0 Then
      For i = 1 To iTotalItems
         If vInvoice(i, VINV_INCLUDE) = 1 Then
            If Trim(vInvoice(i, VINV_ACCOUNT)) = "" Then bByte = True
         End If
      Next
   End If
   If bByte Then
      sMsg = "Each Selected Item Must Have An Account."
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   
   ' looks OK, lets go..
   sVendor = Compress(lblVnd)
   sMsg = "Are Ready To Post This Invoice?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   
   If bResponse = vbYes Then
      'On Error Resume Next
      
      Dim success As Boolean
      success = True
      cFREIGHT = CCur(txtFrt)
      cTax = CCur(txtTax)
      
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      sTxtCmt = diaAPe01a.txtCmt
      sTxtCmt = CheckComments(sTxtCmt)
      
      If success Then
         sSql = "INSERT INTO VihdTable (VINO,VIVENDOR," _
                & "VIDATE,VIDUE,VIDTRECD,VIDUEDATE,VIFREIGHT,VITAX,VICOMT) " _
                & "VALUES('" & lblInv & "','" & sVendor & "','" _
                & lblIdt & "'," & cTot & ",'" & lblPdt & "','" _
                & lblDue & "'," & cFREIGHT & "," _
                & cTax & ",'" & diaAPe01a.txtCmt & "')"
                
         success = clsADOCon.ExecuteSql(sSql)
      End If
      
      'accumulate total part quantity
      Dim partTotalQty As Currency
      partTotalQty = 0
      For i = 1 To iTotalItems
         If Val(vInvoice(i, VINV_INCLUDE)) = 1 And Val(vInvoice(i, VINV_GLDISTRIBUTION)) = 0 Then
            partTotalQty = partTotalQty + CCur(vInvoice(i, VINV_QTY))
         End If
      Next
      
      For i = 1 To iTotalItems
         
         If Val(vInvoice(i, VINV_INCLUDE)) = 1 Then
            
            sAcctNo = vInvoice(i, VINV_ACCOUNT)
            sAcctNo = Compress(sAcctNo)
            
            K = K + 1
            
            cQty = CCur(vInvoice(i, VINV_QTY))
            If cQty = 0 Then cQty = 1
            
            cItemTot = CCur(vInvoice(i, VINV_POCOST))
            cLineTot = Format((cQty * cItemTot) + CCur(vInvoice(i, VINV_ADJUSTMENTS)), CURRENCYMASK)
            
            
            If vInvoice(i, VINV_GLDISTRIBUTION) = 1 Then
               ' NO PO
               If success Then
                  sSql = "INSERT INTO ViitTable (VITNO,VITVENDOR,VITITEM," _
                         & "VITQTY,VITCOST,VITMO,VITMORUN,VITACCOUNT,VITNOTE) " _
                         & "VALUES('" _
                         & lblInv & "','" _
                         & sVendor & "'," _
                         & K & "," _
                         & cQty & "," _
                         & cItemTot & ",'" _
                         & Compress(vInvoice(i, VINV_MO)) & "'," _
                         & Val(vInvoice(i, VINV_RUNNO)) & ",'" _
                         & sAcctNo & "','" _
                         & vInvoice(i, VINV_PARTDESC) & "')"
                  
                  success = clsADOCon.ExecuteSql(sSql)
               End If
               
               If iTrans > 0 Then
                  
                  ' If line item is negative then flip debits and credits.
                  If cLineTot < 0 Then
                     sDebit = "DCCREDIT"
                     sCredit = "DCDEBIT"
                  Else
                     sDebit = "DCDEBIT"
                     sCredit = "DCCREDIT"
                  End If
                  
                  ' Credit
                  iRef = iRef + 1
                  If success Then
                     sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                            & sCredit & ",DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) " _
                            & "VALUES('" _
                            & Trim(sJournalID) & "'," _
                            & iTrans & "," _
                            & iRef & "," _
                            & Abs(cLineTot) & ",'" _
                            & sApAcct & "','" _
                            & sPostDate & "','" _
                            & sVendor & "','" _
                            & Trim(lblInv) & "')"
                     
                     success = clsADOCon.ExecuteSql(sSql)
                  End If
                  
                  ' Debit
                  iRef = iRef + 1
                  If success Then
                     sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                            & sDebit & ",DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) " _
                            & "VALUES('" _
                            & Trim(sJournalID) & "'," _
                            & iTrans & "," _
                            & iRef & "," _
                            & Abs(cLineTot) & ",'" _
                            & sAcctNo & "','" _
                            & sPostDate & "','" _
                            & sVendor & "','" _
                            & Trim(lblInv) & "')"
                     success = clsADOCon.ExecuteSql(sSql)
                   End If
               End If
               
            Else
               
               ' PO
               If success Then
                  sSql = "INSERT INTO ViitTable (VITNO,VITVENDOR,VITITEM," _
                         & "VITPO,VITPORELEASE,VITPOITEM,VITPOITEMREV," _
                         & "VITQTY,VITCOST,VITMO,VITMORUN,VITACCOUNT,VITADDERS) " _
                         & "VALUES('" & lblInv & "','" & sVendor & "'," _
                         & K & "," & Val(lblPon) & "," & Val(lblRel) & "," _
                         & Val(vInvoice(i, VINV_ITEMNO)) & ",'" & vInvoice(i, VINV_ITEMREV) & "'," _
                         & cQty & "," & cItemTot & ",'" _
                         & Compress(vInvoice(i, VINV_MO)) & "'," & Val(vInvoice(i, VINV_RUNNO)) & ",'" & sAcctNo & "'," & CCur(vInvoice(i, VINV_ADJUSTMENTS)) & ")"
                  
                  success = clsADOCon.ExecuteSql(sSql)
               End If
               
               ' Removed the VAL function to calculate PIAMT
               If success Then
                  sSql = "UPDATE PoitTable SET PITYPE=17," _
                         & "PIAMT=" & vInvoice(i, VINV_INVOICEDCOST) / cQty & " " _
                         & "WHERE PINUMBER=" & Val(lblPon) & " " _
                         & "AND PIRELEASE=" & Val(lblRel) & " " _
                         & "AND (PIITEM=" & Val(vInvoice(i, VINV_ITEMNO)) & " " _
                         & "AND PIREV='" & vInvoice(i, VINV_ITEMREV) & "')"
                  
                  success = clsADOCon.ExecuteSql(sSql)
               End If
               
               'calculate unit cost including freight
               Dim totalCost As Currency, unitCost As Currency
               unitCost = 0
               If cQty > 0 Then unitCost = CCur((cLineTot / cQty)) 'Prevents Division by Zero
               If partTotalQty > 0 Then unitCost = unitCost + CCur((cFREIGHT / partTotalQty))   'Prevents Division by zero
               
               totalCost = cQty * unitCost
               
               '                    sPon = "PO " & Trim(lblPon) & "-" _
               '                        & "ITEM " & Trim(vInvoice(i, VINV_ITEMNO)) & vInvoice(i, VINV_ITEMREV)
               '
               '                    ssql = "UPDATE InvaTable SET INAMT=" & unitCost & " " _
               '                        & "WHERE INTYPE=15 AND INREF2='" & sPon & "' "
               
               'if standard cost is being used, update to that now
               'update costs of lot and allocated pick items.
               'success = clsADOCon.ExecuteSql(sSql)
               Dim ia As New ClassInventoryActivity
               Dim sReturnMsg As String
               sReturnMsg = ia.UpdateReceiptCosts(CStr(vInvoice(i, VINV_PARTNO)), CLng(Trim(lblPon)), _
                            CInt(Trim(vInvoice(i, VINV_ITEMNO))), CStr(vInvoice(i, VINV_ITEMREV)), _
                            cQty, ByVal unitCost, True)
               
               ' 10/12/2010 Update the MO cost for MO costing
               Dim rdoMoPart As ADODB.Recordset
               Dim strINMoPart As String
               Dim strINMoRun As String
               
               If success Then
                  sSql = " SELECT DISTINCT INMOPART, INMORUN FROM InvaTable " _
                           & " WHERE INTYPE = " & IATYPE_PickedItem & " AND  INPART ='" & CStr(Compress(vInvoice(i, VINV_PARTNO))) & "'" _
                           & " AND INLOTNUMBER IN " _
                        & " (SELECT DISTINCT INLOTNUMBER FROM InvaTable b, partTable" _
                           & " WHERE b.INTYPE = " & IATYPE_PoReceipt & " AND  b.INPART ='" & CStr(Compress(vInvoice(i, VINV_PARTNO))) & "'" _
                              & " AND b.INPART = partTable.PartRef" _
                           & " AND b.INPONUMBER ='" & CStr(Trim(lblPon)) & "' AND  b.INPOITEM ='" & CStr(Trim(vInvoice(i, VINV_ITEMNO))) & "' AND PALEVEL IN (4,7))"
                  bSqlRows = clsADOCon.GetDataSet(sSql, rdoMoPart, ES_STATIC)
                  If bSqlRows Then
                  
                     With rdoMoPart
                        Do Until .EOF
                           strINMoPart = "" & Trim(!INMOPART)
                           strINMoRun = "" & Trim(!INMORUN)
                        
                           If (strINMoPart <> "" And strINMoRun <> "") Then
                              Dim mo As New ClassMO
                              Dim bRet As Boolean
                              mo.PartNumber = strINMoPart
                              mo.RunNumber = strINMoRun
                              ' Recalculate the cost for MO part
                              bRet = mo.RecalculateCosts()
                           End If
                           .MoveNext
                        Loop
                     End With
                     ClearResultSet rdoMoPart
                  End If
                  Set rdoMoPart = Nothing
               End If
               
               'Update Lot Header
               AverageCost Format$(vInvoice(i, VINV_PARTNO))
               
               If iTrans > 0 Then
                  
                  ' Credit
                  iRef = iRef + 1
                  If success Then
                     sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                            & "DCCREDIT,DCACCTNO,DCPARTNO,DCPONUMBER,DCPOITEM," _
                            & "DCPOITREV,DCDATE,DCVENDOR,DCVENDORINV) " _
                            & "VALUES('" _
                            & Trim(sJournalID) & "'," _
                            & iTrans & "," _
                            & iRef & "," _
                            & cLineTot & ",'" _
                            & sApAcct & "','" _
                            & Compress(vInvoice(i, VINV_PARTNO)) & "'," _
                            & Val(lblPon) & "," _
                            & Val(vInvoice(i, VINV_ITEMNO)) & ",'" _
                            & vInvoice(i, VINV_ITEMREV) & "','" _
                            & sPostDate & "','" _
                            & sVendor & "','" _
                            & Trim(lblInv) & "')"
                  
                     success = clsADOCon.ExecuteSql(sSql)
                  End If
                  
                  ' Debit
                  iRef = iRef + 1
                  If success Then
                     sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                            & "DCDEBIT,DCACCTNO,DCPARTNO,DCPONUMBER,DCPOITEM," _
                            & "DCPOITREV,DCDATE,DCVENDOR,DCVENDORINV) " _
                            & "VALUES('" _
                            & Trim(sJournalID) & "'," _
                            & iTrans & "," _
                            & iRef & "," _
                            & cLineTot & ",'" _
                            & sAcctNo & "','" _
                            & Compress(vInvoice(i, VINV_PARTNO)) & "'," _
                            & Val(lblPon) & "," _
                            & Val(vInvoice(i, VINV_ITEMNO)) & ",'" _
                            & vInvoice(i, VINV_ITEMREV) & "','" _
                            & sPostDate & "','" _
                            & sVendor & "','" _
                            & Trim(lblInv) & "')"
                     success = clsADOCon.ExecuteSql(sSql)
                  End If
               End If
            End If
         End If
      Next
      
      ' Tax
      If Val(txtTax) > 0 Then
         ' Credit
         iRef = iRef + 1
         If success Then
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                   & "DCCREDIT,DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) " _
                   & "VALUES('" _
                   & Trim(sJournalID) & "'," _
                   & iTrans & "," _
                   & iRef & "," _
                   & CCur(txtTax) & ",'" _
                   & sApAcct & "','" _
                   & sPostDate & "','" _
                   & sVendor & "','" _
                   & Trim(lblInv) & "')"
               success = clsADOCon.ExecuteSql(sSql)
         End If
         
         ' Debit
         iRef = iRef + 1
         If success Then
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                   & "DCDEBIT,DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) " _
                   & "VALUES('" _
                   & Trim(sJournalID) & "'," _
                   & iTrans & "," _
                   & iRef & "," _
                   & CCur(txtTax) & ",'" _
                   & sTxAcct & "','" _
                   & sPostDate & "','" _
                   & sVendor & "','" _
                   & Trim(lblInv) & "')"
            success = clsADOCon.ExecuteSql(sSql)
         End If
      End If
      
      'Freight
      If Val(txtFrt) > 0 Then
         ' Credit
         iRef = iRef + 1
         If success Then
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                   & "DCCREDIT,DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) " _
                   & "VALUES('" _
                   & Trim(sJournalID) & "'," _
                   & iTrans & "," _
                   & iRef & "," _
                   & CCur(txtFrt) & ",'" _
                   & sApAcct & "','" _
                   & sPostDate & "','" _
                   & sVendor & "','" _
                   & Trim(lblInv) & "')"
            success = clsADOCon.ExecuteSql(sSql)
         End If
         
         ' Debit
         iRef = iRef + 1
         If success Then
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                   & "DCDEBIT,DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) " _
                   & "VALUES('" _
                   & Trim(sJournalID) & "'," _
                   & iTrans & "," _
                   & iRef & "," _
                   & CCur(txtFrt) & ",'" _
                   & sFrAcct & "','" _
                   & sPostDate & "','" _
                   & sVendor & "','" _
                   & Trim(lblInv) & "')"
            success = clsADOCon.ExecuteSql(sSql)
         End If
      End If
      
      
      
      If success Then
         clsADOCon.CommitTrans
         bPosted = True
         sMsg = "Successfully Posted Invoice"
         SysMsg sMsg, 1, Me
         diaAPe01a.optPst.Value = vbChecked
         
         If (cmdDockLok.Visible = True) Then
            sMsg = "Do you want to Scan the Vendor Invoice ?"
            bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
            If bResponse = vbNo Then
               Unload Me
            Else
               cmdpst.enabled = False
               cmdDockLok_Click
            End If
         Else
            Unload Me
         End If
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         
         bPosted = False
         sMsg = "Could Not Post This Invoice." & vbCrLf _
                & "Transaction Canceled."
         MsgBox sMsg, vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "postinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function GetInvAccount(PartNo As String, Level As Integer, _
                               PCode As String, Account As String) As Byte
   Dim rdoAct As ADODB.Recordset
   Dim sPart As String
   'Check the part
   On Error GoTo DiaErr1
   sPart = Compress(PartNo)
   If Level = 5 Or Level = 6 Or Level = 7 Then
      sSql = "SELECT PAINVEXPACCT FROM PartTable WHERE PARTREF='" _
             & sPart & "' "
   Else
      sSql = "SELECT PAINVMATACCT FROM PartTable WHERE PARTREF='" _
             & sPart & "' "
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         Account = "" & Trim(.Fields(0))
         .Cancel
      End With
   Else
      Account = ""
   End If
   'rdoAct.Close
   'Product Code
   If Account = "" Then
      If Trim(PCode) <> "" Then
         PCode = Compress(PCode)
         If Level = 5 Or Level = 6 Or Level = 7 Then
            sSql = "SELECT PCINVEXPACCT FROM PcodTable WHERE " _
                   & "PCREF='" & PCode & "' "
         Else
            sSql = "SELECT PCINVMATACCT FROM PcodTable WHERE " _
                   & "PCREF='" & PCode & "' "
         End If
         bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
         If bSqlRows Then
            With rdoAct
               Account = "" & Trim(.Fields(0))
               .Cancel
            End With
         End If
      End If
   Else
      GetInvAccount = 1
      Set rdoAct = Nothing
      Exit Function
   End If
   
   'Okay, maybe still none
   If Account = "" Then
      If Level = 5 Or Level = 6 Or Level = 7 Then
         sSql = "SELECT COINVEXPACCT" & Trim(Level) & " FROM " _
                & "ComnTable WHERE COREF=1"
      Else
         sSql = "SELECT COINVMATACCT" & Trim(Level) & " FROM " _
                & "ComnTable WHERE COREF=1"
      End If
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            Account = "" & Trim(.Fields(0))
            .Cancel
         End With
      End If
   Else
      GetInvAccount = 1
      Set rdoAct = Nothing
      Exit Function
   End If
   'Last chance
   If Account = "" Then
      sSql = "SELECT COREVACCT" & Trim(Level) & " FROM " _
             & "ComnTable WHERE COREF=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            Account = "" & Trim(.Fields(0))
            .Cancel
         End With
      End If
   End If
   If Trim(Account) = "" Then GetInvAccount = 0
   Set rdoAct = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getinvacco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Public Sub AddGLDis()
   
   On Error GoTo DiaErr1
   
   ' vInvoice Documentation
   ' 0 = Item
   ' 1 = Item Rev
   ' 2 = Part Number
   ' 3 = Part Desc
   ' 4 = PO Qty
   ' 5 = PO Cost
   ' 6 = Invoiced Cost
   ' 7 = Account
   ' 8 = To Be Invoiced (0 or 1)
   ' 9 = Type PO(0) / NO PO(1)
   
   vInvoice(iTotalItems + 1, VINV_ITEMNO) = iTotalItems + 1
   vInvoice(iTotalItems + 1, VINV_ITEMREV) = ""
   vInvoice(iTotalItems + 1, VINV_PARTNO) = diaAPe01c.cmbAct & " - " & diaAPe01c.lblDsc
   vInvoice(iTotalItems + 1, VINV_PARTDESC) = diaAPe01c.txtCom
   vInvoice(iTotalItems + 1, VINV_QTY) = 1
   vInvoice(iTotalItems + 1, VINV_LOTQTY) = 1
   vInvoice(iTotalItems + 1, VINV_POCOST) = diaAPe01c.txtAmt
   vInvoice(iTotalItems + 1, VINV_INVOICEDCOST) = diaAPe01c.txtAmt
   vInvoice(iTotalItems + 1, VINV_ACCOUNT) = Compress(diaAPe01c.cmbAct)
   vInvoice(iTotalItems + 1, VINV_INCLUDE) = 1
   vInvoice(iTotalItems + 1, VINV_GLDISTRIBUTION) = 1
   vInvoice(iTotalItems + 1, VINV_ADJUSTMENTS) = 0
   
   ' 12 = MO
   ' 13 = Run
   vInvoice(iTotalItems + 1, VINV_MO) = Trim(diaAPe01c.cmbMon)
   If vInvoice(iTotalItems + 1, VINV_MO) = "<NONE>" Then
      vInvoice(iTotalItems + 1, VINV_MO) = ""
      vInvoice(iTotalItems + 1, VINV_RUNNO) = ""
   Else
      vInvoice(iTotalItems + 1, VINV_RUNNO) = diaAPe01c.cmbRun
   End If
   
   SysMsg "Added To Invoice.", 1, diaAPe01c
   iTotalItems = iTotalItems + 1
   
   lblLst = CInt((((iTotalItems * 100) / 100) / NEXTPAGE) + 0.5)
   
   ' Check if we should turn on the down arrow key.
   If Val(lblPge) < Val(lblLst) Then
      cmdDn.Picture = Endn
      cmdDn.enabled = True
   End If
   GetNextGroup
   UpdateTotals
   diaAPe01c.txtCom.SetFocus
   
   Exit Sub
DiaErr1:
   sProcName = "Addgldis"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub ViewItemComments(iItem As Integer, sRev As String)
   Dim lPO As Long: lPO = Val(lblPon)
   Dim iRel As Integer: iRel = Val(lblRel)
   Dim RdoCmt As ADODB.Recordset
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   sSql = "SELECT PICOMT,Convert(varchar(12), PIRECEIVED, 101) PIRECEIVED, " _
         & " Convert(varchar(12), PIODDELDATE, 101) PIODDELDATE,  " _
         & " POREQBY FROM PoitTable, PohdTable  WHERE PONUMBER = PINUMBER AND " _
          & " PINUMBER = " & lPO _
          & " AND PIRELEASE = " & iRel _
          & " AND PIITEM = " & iItem _
          & " AND PIREV = '" & sRev & "'"
        
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmt, ES_STATIC)
   vewPoItmCmt.Left = (Me.Left + Me.ActiveControl.Left) + MdiSect.Sidebar.Width
   If iBarOnTop = 0 Then
      vewPoItmCmt.Left = (Me.Left + Me.ActiveControl.Left) + MdiSect.Sidebar.Width
      vewPoItmCmt.Top = (Me.Top + (Me.ActiveControl.Top + Me.ActiveControl.Height)) + 1000
   Else
      vewPoItmCmt.Top = (Me.Top + (Me.ActiveControl.Top + Me.ActiveControl.Height)) + (1000 + MdiSect.TopBar.Height)
      vewPoItmCmt.Left = (Me.Left + Me.ActiveControl.Left) '+ MdiSect.SideBar.Width
   End If
   
   sMsg = "DELIVERY DATE : " & Trim(RdoCmt!PIODDELDATE) & vbCrLf & vbCrLf
   sMsg = sMsg & "REQUESTED BY : " & Trim(RdoCmt!POREQBY) & vbCrLf & vbCrLf
   
   vewPoItmCmt.txtCmt = sMsg & "PO COMMENT: " & vbCrLf & Trim(RdoCmt!PICOMT)
   
   'vewPoItmCmt.txtCmt = "" & Trim(RdoCmt!PICOMT)
   vewPoItmCmt.Show
   Set RdoCmt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "viewitemcomm"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function GetRuns(i As Integer) As Boolean
   Dim RdoRns As ADODB.Recordset
   Dim SPartRef As String
   
   cmbRun(i).Clear
   SPartRef = Compress(cmbMon(i))
   If Len(SPartRef) > 0 Then
      On Error GoTo DiaErr1
      sSql = "SELECT PARTREF,PARTNUM,PADESC,RUNREF,RUNSTATUS," _
             & "RUNNO FROM PartTable,RunsTable WHERE PARTREF='" _
             & SPartRef & "' AND PARTREF=RUNREF AND RUNSTATUS<>'CA'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns)
      If bSqlRows Then
         With RdoRns
            cmbRun(i) = Format(0 + !Runno, "####0")
            Do Until .EOF
               cmbRun(i).AddItem Format(0 + !Runno, "####0")
               .MoveNext
            Loop
            .Cancel
         End With
         GetRuns = True
      Else
         SPartRef = ""
         GetRuns = False
      End If
   End If
   Set RdoRns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub FillMOs()
   Dim RdoFrn As ADODB.Recordset
   Dim b As Byte
   On Error GoTo DiaErr1
   
   MouseCursor 13
   For b = 1 To 5
      cmbMon(b).Clear
   Next
   
   sSql = "Qry_RunsNotCanceled"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFrn, ES_FORWARD)
   If bSqlRows Then
      With RdoFrn
         While Not .EOF
            For b = 1 To 5
               AddComboStr cmbMon(b).hWnd, "" & Trim(!PARTNUM)
            Next
            .MoveNext
         Wend
         .Cancel
      End With
  End If

'            For b = 1 To 5
'               AddComboStr cmbMon(b).hWnd, ""   'empty field to type into
'            Next


' doesn't work -- datasource property not recognized
'      With RdoFrn
'         For b = 1 To 5
'            cmbMon(b).DataField = "PARTNUM"
'            cmbMon(b).DataSource = RdoFrn
'         Next
'         .Cancel
'      End With
'  End If

   On Error Resume Next
   Set RdoFrn = Nothing
   MouseCursor 0
   Exit Sub
DiaErr1:
   MouseCursor 0
   sProcName = "FillMOs"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
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


