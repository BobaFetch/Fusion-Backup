VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaARe01c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Invoice Items"
   ClientHeight    =   6615
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   11310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPrt 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox txtPrt 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtPrt 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtPrt 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtPrt 
      BackColor       =   &H8000000F&
      CausesValidation=   0   'False
      Height          =   285
      Index           =   1
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CheckBox optCan 
      Caption         =   "OptCan"
      Height          =   255
      Left            =   7080
      TabIndex        =   78
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox optItm 
      Height          =   255
      Left            =   120
      TabIndex        =   76
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   5
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New Item"
      Height          =   315
      Left            =   120
      TabIndex        =   65
      ToolTipText     =   "Add Item Not Found On Sales Order"
      Top             =   5040
      Width           =   875
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   4
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   3
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   2
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox txtAct 
      Height          =   315
      Index           =   5
      Left            =   7320
      TabIndex        =   15
      Tag             =   "3"
      Top             =   4320
      Width           =   1275
   End
   Begin VB.ComboBox txtAct 
      Height          =   315
      Index           =   4
      Left            =   7320
      TabIndex        =   12
      Tag             =   "3"
      Top             =   3600
      Width           =   1275
   End
   Begin VB.ComboBox txtAct 
      Height          =   315
      Index           =   3
      Left            =   7320
      TabIndex        =   9
      Tag             =   "3"
      Top             =   2880
      Width           =   1275
   End
   Begin VB.ComboBox txtAct 
      Height          =   315
      Index           =   2
      Left            =   7320
      TabIndex        =   6
      Tag             =   "3"
      Top             =   2160
      Width           =   1275
   End
   Begin VB.ComboBox txtAct 
      Height          =   315
      Index           =   1
      Left            =   7320
      TabIndex        =   3
      Tag             =   "3"
      Top             =   1440
      Width           =   1275
   End
   Begin VB.CheckBox optAll 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   6360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
   End
   Begin VB.CheckBox chkShip 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   10800
      TabIndex        =   14
      ToolTipText     =   "Ship As Complete"
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtShp 
      Height          =   285
      Index           =   5
      Left            =   5280
      TabIndex        =   13
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkShip 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   10800
      TabIndex        =   11
      ToolTipText     =   "Ship As Complete"
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtShp 
      Height          =   285
      Index           =   4
      Left            =   5280
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkShip 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   10800
      TabIndex        =   2
      ToolTipText     =   "Ship As Complete"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtShp 
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   1
      ToolTipText     =   "Quantity To Ship"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CheckBox chkShip 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   10800
      TabIndex        =   5
      ToolTipText     =   "Ship As Complete"
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtShp 
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtShp 
      Height          =   285
      Index           =   3
      Left            =   5280
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkShip 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   10800
      TabIndex        =   8
      ToolTipText     =   "Ship As Complete"
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Post"
      Height          =   315
      Left            =   10320
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Add Selections To Pack Slip (When Finished)"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   10320
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   17
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
      PictureUp       =   "diaARe01c.frx":0000
      PictureDn       =   "diaARe01c.frx":0146
   End
   Begin Threed.SSCommand cmdDn 
      Height          =   375
      Left            =   10920
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Next Page (Page Down)"
      Top             =   6060
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaARe01c.frx":028C
   End
   Begin Threed.SSCommand cmdUp 
      Height          =   375
      Left            =   10920
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Last Page (Page Up)"
      Top             =   5700
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaARe01c.frx":078E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8280
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6615
      FormDesignWidth =   11310
   End
   Begin VB.Label lblSchedDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   6360
      TabIndex        =   134
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblSchedDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   6360
      TabIndex        =   133
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblSchedDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   132
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblSchedDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   131
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblSchedDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   130
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sched Dte"
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
      Index           =   20
      Left            =   6360
      TabIndex        =   129
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblUnitPrice 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   8640
      TabIndex        =   128
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price        "
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
      Index           =   19
      Left            =   8640
      TabIndex        =   127
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblUnitPrice 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   8640
      TabIndex        =   126
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblUnitPrice 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   8640
      TabIndex        =   125
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblUnitPrice 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   8640
      TabIndex        =   124
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblUnitPrice 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   8640
      TabIndex        =   123
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblFedTaxAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   10320
      TabIndex        =   117
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   4620
      Width           =   720
   End
   Begin VB.Label lblFedTaxRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   9720
      TabIndex        =   116
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   4620
      Width           =   495
   End
   Begin VB.Label lblFedTaxCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   7320
      TabIndex        =   115
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   4620
      Width           =   1200
   End
   Begin VB.Label lblStateTaxAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   5880
      TabIndex        =   114
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   4620
      Width           =   720
   End
   Begin VB.Label lblStateTaxRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   5280
      TabIndex        =   113
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   4620
      Width           =   615
   End
   Begin VB.Label lblStateTaxCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   3840
      TabIndex        =   112
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   4620
      Width           =   1200
   End
   Begin VB.Label lblFedTaxAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   10320
      TabIndex        =   111
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   3900
      Width           =   720
   End
   Begin VB.Label lblFedTaxRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   9720
      TabIndex        =   110
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   3900
      Width           =   495
   End
   Begin VB.Label lblFedTaxCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   7320
      TabIndex        =   109
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   3900
      Width           =   1200
   End
   Begin VB.Label lblStateTaxAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   5880
      TabIndex        =   108
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   3900
      Width           =   720
   End
   Begin VB.Label lblStateTaxRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   5280
      TabIndex        =   107
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   3900
      Width           =   615
   End
   Begin VB.Label lblStateTaxCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   3840
      TabIndex        =   106
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   3900
      Width           =   1200
   End
   Begin VB.Label lblFedTaxAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   10320
      TabIndex        =   105
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   3240
      Width           =   720
   End
   Begin VB.Label lblFedTaxRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   9720
      TabIndex        =   104
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lblFedTaxCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   7320
      TabIndex        =   103
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   3180
      Width           =   1200
   End
   Begin VB.Label lblStateTaxAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   5880
      TabIndex        =   102
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   3180
      Width           =   720
   End
   Begin VB.Label lblStateTaxRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   5280
      TabIndex        =   101
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   3180
      Width           =   615
   End
   Begin VB.Label lblStateTaxCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   100
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   3180
      Width           =   1200
   End
   Begin VB.Label lblFedTaxAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   10320
      TabIndex        =   99
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   2460
      Width           =   720
   End
   Begin VB.Label lblFedTaxRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   9720
      TabIndex        =   98
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   2460
      Width           =   495
   End
   Begin VB.Label lblFedTaxCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   7320
      TabIndex        =   97
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   2460
      Width           =   1200
   End
   Begin VB.Label lblStateTaxAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   5880
      TabIndex        =   96
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   2460
      Width           =   720
   End
   Begin VB.Label lblStateTaxRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   95
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   2460
      Width           =   615
   End
   Begin VB.Label lblStateTaxCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   94
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   2460
      Width           =   1200
   End
   Begin VB.Label lblFedTaxAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   10320
      TabIndex        =   93
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   1740
      Width           =   720
   End
   Begin VB.Label lblFedTaxRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   9720
      TabIndex        =   92
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   1740
      Width           =   495
   End
   Begin VB.Label lblFedTaxCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   7320
      TabIndex        =   91
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label lblStateTaxAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   5880
      TabIndex        =   90
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   1740
      Width           =   720
   End
   Begin VB.Label lblStateTaxRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   89
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   1740
      Width           =   615
   End
   Begin VB.Label lblStateTaxCode 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   88
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   1740
      Width           =   1200
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Federal Tax"
      Height          =   255
      Index           =   18
      Left            =   8880
      TabIndex        =   87
      Top             =   5340
      Width           =   975
   End
   Begin VB.Label lblFedTax 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   9960
      TabIndex        =   86
      Top             =   5340
      Width           =   855
   End
   Begin VB.Label lblLotTracked 
      Caption         =   "Ship lot-tracked items with packing slips"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   5
      Left            =   3840
      TabIndex        =   85
      Top             =   4680
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Label lblLotTracked 
      Caption         =   "Ship lot-tracked items with packing slips"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   3840
      TabIndex        =   84
      Top             =   3960
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Label lblLotTracked 
      Caption         =   "Ship lot-tracked items with packing slips"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   3840
      TabIndex        =   83
      Top             =   3240
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Label lblLotTracked 
      Caption         =   "Ship lot-tracked items with packing slips"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   3840
      TabIndex        =   82
      Top             =   2520
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Label lblLotTracked 
      Caption         =   "Ship lot-tracked items with packing slips"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   3840
      TabIndex        =   81
      Top             =   1800
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   80
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblitm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   79
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "UOM"
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
      Index           =   17
      Left            =   4800
      TabIndex        =   77
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblFrt 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   9960
      TabIndex        =   75
      Top             =   5580
      Width           =   855
   End
   Begin VB.Label lblTax 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   9960
      TabIndex        =   74
      Top             =   5100
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax"
      Height          =   255
      Index           =   16
      Left            =   8880
      TabIndex        =   73
      Top             =   5100
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight"
      Height          =   255
      Index           =   15
      Left            =   8880
      TabIndex        =   72
      Top             =   5580
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ship All"
      Height          =   255
      Index           =   14
      Left            =   5640
      TabIndex        =   71
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   9960
      TabIndex        =   70
      Top             =   5820
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Total"
      Height          =   255
      Index           =   13
      Left            =   8880
      TabIndex        =   69
      Top             =   5880
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   8640
      X2              =   11280
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev Account               "
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
      Left            =   7320
      TabIndex        =   67
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount         "
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
      Index           =   11
      Left            =   9720
      TabIndex        =   66
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblOrd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   3840
      TabIndex        =   60
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   600
      TabIndex        =   59
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblitm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   58
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   57
      Top             =   4635
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   4800
      TabIndex        =   56
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblOrd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   3840
      TabIndex        =   55
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   54
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblitm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   53
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   52
      Top             =   3915
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   4800
      TabIndex        =   51
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblSon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3600
      TabIndex        =   50
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   49
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   48
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4800
      TabIndex        =   47
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   46
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   45
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item   "
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
      Index           =   2
      Left            =   120
      TabIndex        =   44
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev  "
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
      Index           =   3
      Left            =   600
      TabIndex        =   43
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                               "
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
      Left            =   1080
      TabIndex        =   42
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order Qty    "
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
      Left            =   3840
      TabIndex        =   41
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ship Qty         "
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
      Left            =   5280
      TabIndex        =   40
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ship    "
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
      Index           =   7
      Left            =   10800
      TabIndex        =   39
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   38
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   1755
      Width           =   2655
   End
   Begin VB.Label lblOrd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   37
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   36
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblitm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   35
      ToolTipText     =   "Right Click Mouse For Comments"
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblOrd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   34
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   33
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblitm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   32
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   31
      Top             =   2475
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   30
      Top             =   3195
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblOrd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   29
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   8280
      Picture         =   "diaARe01c.frx":0C90
      Top             =   6180
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   7320
      Picture         =   "diaARe01c.frx":1182
      Top             =   6180
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   7680
      Picture         =   "diaARe01c.frx":1674
      Top             =   6180
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   7920
      Picture         =   "diaARe01c.frx":1B66
      Top             =   6180
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
      Left            =   9480
      TabIndex        =   28
      Top             =   6180
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
      Left            =   10320
      TabIndex        =   27
      Top             =   6180
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   255
      Index           =   9
      Left            =   8880
      TabIndex        =   26
      Top             =   6180
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      Height          =   255
      Index           =   8
      Left            =   9960
      TabIndex        =   25
      Top             =   6180
      Width           =   375
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2520
      TabIndex        =   21
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   20
      Top             =   585
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   585
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   18
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "diaARe01c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*************************************************************************************
' diaARe01c - Add remove sales order items to invoice
'
' Created: (cjs)
' Revisions:
'   07/26/02 (nth) Added item amount (Qty * ITDOLLARS).
'   07/31/02 (nth) Fixed error with page up and down keys.vItems(i, 8) * vItems(i, 13)e
'   12/18/02 (nth) Added tax code logic to invoice creation.
'   10/14/03 (nth) Added new sales tax logic per WCK uses PATAXEXEMPT.
'   10/14/03 (nth) Added GetCustomerType to support B&O taxes.
'   03/02/04 (nth) Added lot select
'   03/24/04 (nth) Fixed error with LoitTable and InvaTable activity counter not
'                  auto incrementing
'   08/03/04 (nth) Cache current standards for part being invoiced in InvaTable
'   07/19/05 TEL Made it impossible to ship a lot-tracked item in this transaction
'
'*************************************************************************************

Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim bOnLoad As Byte
Dim bInvSaved As Byte
Dim bGoodSO As Byte
Dim sMsg As String
Dim iCurrPage As Integer
Dim iNewBookmark As Integer
Dim iTotalItems As Integer
Dim iMaxPages As Integer
Dim lInv As Long
Dim lSalesOrder As Long
Dim sSOCust As String
Dim sAccount As String
Dim sCreditAcct As String
Dim sDebitAcct As String

'Sales journal
Dim sJournal As String
Dim sCOSjARAcct As String
Dim sCOSjINVAcct As String
Dim sCOSjNFRTAcct As String
Dim sCOSjTFRTAcct As String
Dim sCOSjTaxAcct As String
Dim sCOSjFedTaxAcct As String

' Sales tax
Dim sTaxCode As String
Dim sTaxState As String
Dim nTaxRate As Currency
Dim sTaxAcct As String
Dim sType As String * 1

Dim sLots(50, 2) As String

Dim taxPerLineItem As Boolean

'Dim FusionToolTip As New clsToolTip



'array to hold results
'this approach sucks.  should be reengineered to hold results in a grid.
Const MAX_ITEMS = 2048
Dim vItems(MAX_ITEMS, 28) As Variant
'   vItems(iTotalitems, 0) = Right(lblSon, 5)
'   vItems(iTotalitems, 1) = Format(!ITNUMBER, "#0")
Const INV_ItemNumber = 1
'   vItems(iTotalitems, 2) = Trim(!ITREV)
Const INV_ItemRev = 2
'   vItems(iTotalitems, 3) = Trim(!PARTNUM)
'   vItems(iTotalitems, 4) = Trim(!PADESC)
'   vItems(iTotalitems, 5) = Trim(!ITQTY)
Const INV_OrderQty = 5
'   vItems(iTotalitems, 6) = Trim(!ITCOMMENTS)
'   vItems(iTotalitems, 7) = Trim(!PAUNITS)
'   vItems(iTotalitems, 8) = Ship Qty
Const INV_ShipQty = 8
'   vItems(iTotalitems, 9) = Ship? (T/F)
Const INV_ShipThisItem = 9
'   vItems(iTotalitems,10) = Compressed Part
Const INV_PartRef = 10
'   vItems(iTotalitems,11) = Standard Cost (or Actual later)
Const INV_Cost_NoLongerUsed = 11
'   vItems(iTotalitems,12) = Account
Const INV_PartRevAccount = 12
'   vItems(iTotalitems,13) = ITDOLLARS
Const INV_UnitPrice = 13

'   vItems(iTotalitems,14) = Part Type (Level)
'   vItems(iTotalitems,15) = Product Code
'   vItems(iTotalitems,16) = Taxable (!PATAXEXEMPT)
Const INV_TaxExempt = 16 '(!PATAXEXEMPT)
'   vItems(iTotalitems,17) = Lot Tracked
Const INV_LotTracked = 17

Const INV_SalesTaxCode = 18
Const INV_SalesTaxAcct = 19
Const INV_SalesTaxRate = 20
Const INV_SalesTaxAmt = 21

Const INV_FedTaxCode = 22
Const INV_FedTaxAcct = 23
Const INV_FedTaxRate = 24
Const INV_FedTaxAmt = 25
Const INV_SchedDate = 27
Const INV_UnitCost = 28

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'See if the Accounts are there

Public Sub GetSJAccounts()
   Dim rdoJrn As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   
   On Error GoTo DiaErr1
   sSql = "SELECT COREF,COSJARACCT,COSJINVACCT,COSJNFRTACCT," _
          & "COSJTFRTACCT,COSJTAXACCT,COFEDTAXACCT FROM ComnTable WHERE " _
          & "COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      With rdoJrn
         sCOSjARAcct = "" & Trim(.Fields(1))
         sCOSjINVAcct = "" & Trim(.Fields(2))
         sCOSjNFRTAcct = "" & Trim(.Fields(3))
         sCOSjTFRTAcct = "" & Trim(.Fields(4))
         sCOSjTaxAcct = "" & Trim(.Fields(5))
         sCOSjFedTaxAcct = "" & Trim(.Fields(6))
         .Cancel
      End With
   End If
   Set rdoJrn = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsjacco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdNew_Click()
   optItm = vbChecked
   diaARe01d.Show
End Sub




Private Sub lblSon_Click()
   CloseBoxes False
End Sub

Private Sub cmdAdd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "Are You Finished With Your Selections?" & vbCrLf _
          & "Do You Wish to Record The Items?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      CreateInvoice
   Else
      CancelTrans
      On Error Resume Next
      txtShp(1).SetFocus
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdDn_Click()
   iCurrPage = iCurrPage + 1
   GetNextGroup
End Sub

Private Sub cmdDn_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Customer Invoice (Sales Order)"
      MouseCursor 0
      cmdHlp = False
   End If
End Sub

Private Sub cmdUp_Click()
   iCurrPage = iCurrPage - 1
   GetNextGroup
End Sub

Private Sub cmdUp_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   lblFrt = diaARe01b.txtFrt
   If bOnLoad Then
      ' is tax calculated per line item?
      taxPerLineItem = IsTaxCalculatedPerLineItem()
      
      GetSJAccounts
      FillAccounts
      lSalesOrder = Val(Right(lblSon, SO_NUM_SIZE))
      sSOCust = Compress(lblCst)
      GetSalesTaxInfo sSOCust, nTaxRate, sTaxCode, sTaxState, sTaxAcct
      bGoodSO = GetItems()
      UpdateTotals
      bOnLoad = False
      
   End If
   UpdateTotals
   MouseCursor 0
   Call HookToolTips
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
   
   ' 10/19/06 if a new item is added, it's fairly complicated to display it (maybe)
   ' and add it to the vItems array (always) without wiping out any information typed into
   ' the form and vItems array.  So we're just going to take away this feature
   cmdNew.Visible = False
   
   SetFormSize Me
   Move diaARe01b.Top + 200, diaARe01b.Left + 200
   lblInv = diaARe01b.lblInv
   lInv = Val(Right(lblInv, 6))
   
   sCurrForm = ""
   sAccount = GetSetting("Esi2000", "Fina", "LastRAccount", sAccount)
   
   sSql = "SELECT ITSO,ITNUMBER,ITREV,ITQTY,ITPART," _
          & "ITCOMMENTS,ITDOLLARS,PARTREF,PARTNUM,PADESC,PAUNITS,PASTDCOST,PAEXTDESC, " _
          & "PALEVEL,PAPRICE,PAPRODCODE,PATAXEXEMPT,PALOTTRACK," _
          & "ITTAXCODE,ITSLSTXACCT,ITTAXRATE,ITTAXAMT," _
          & "ITFEDTAXCODE,ITFEDTAXACCT,ITFEDTAXRATE,ITFEDTAXAMT, ITDOLLORIG, ITSCHED" & vbCrLf _
          & "FROM SoitTable,PartTable " _
          & "WHERE ITSO= ? AND (ITPART=PARTREF AND ITACTUAL IS NULL AND " _
          & "ITCANCELED=0 AND ITQTY>0 AND ITPSNUMBER='' AND ITINV='')"
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adInteger
   AdoQry.parameters.Append AdoParameter
   
   
   CloseBoxes False
   
   bInvSaved = False
   bOnLoad = True
   
   'With FusionToolTip
   '   Call .Create(Me)
   '   .MaxTipWidth = 240
   '   .DelayTime(ttDelayShow) = 20000
   '    For i = 1 To 5
   '         Call .AddTool(txtPrt(i))
   '         txtPrt(i).BackColor = &H8000000F
   '         txtPrt(i).Locked = True
   '
   '    Next i
   'End With
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim i As Integer
   Dim sMsg As String
   On Error Resume Next
   
   ' Check if diaARe01b is closing this form
   If diaARe01b.optCan = vbUnchecked Then
      If Not bInvSaved Then
         sMsg = "You Have Not Recorded Any Items." & vbCrLf _
                & "Do You Really Want to Close?"
         i = MsgBox(sMsg, ES_NOQUESTION, Caption)
         If i <> vbYes Then
            Cancel = True
            On Error Resume Next
            SetFocus
            Exit Sub
         End If
      Else
         diaARe01b.optSav = vbChecked
      End If
   End If
   diaARe01b.optItm.Value = vbUnchecked
   
   ' Refresh invoice cursor
   Set diaARe01b.RdoInv = Nothing
   diaARe01b.bGoodInvoice = diaARe01b.GetInvoice()
   
   If Len(Trim(sAccount)) Then
      SaveSetting "Esi2000", "Fina", "LastRAccount", sAccount
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   'MdiSect.BotPanel = diaARe01a.Caption
   Erase vItems
   On Error Resume Next
   UnhookToolTips
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set diaARe01c = Nothing
   
End Sub

Public Function GetItems() As Byte
   Dim RdoItm As ADODB.Recordset
   Dim bSoIncluded
   Dim i As Integer
   'Dim a As Variant
   Dim sPartRevAcct As String
   Dim bByte As Byte
   Dim sItemCmts, sPartExtDesc As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   cmdUp.Picture = Dsup
   cmdDn.Picture = Dsdn
   cmdUp.enabled = False
   cmdDn.enabled = False
   lblLst = 1
   lblPge = 1
   bSoIncluded = False
   If bSoIncluded Then
      GetItems = True
      Exit Function
   End If
   i = 0
   AdoQry.parameters(0).Value = Val(Right(lblSon, SO_NUM_SIZE))
   bSqlRows = clsADOCon.GetQuerySet(RdoItm, AdoQry, ES_STATIC) 'ES_STATIC or ES_KEYSET required for blob-like ITCOMMENTS
   If bSqlRows Then
      
      With RdoItm
         iTotalItems = -1
         Do Until .EOF
            i = i + 1
            If i < 6 Then
               lblitm(i).Visible = True
               lblRev(i).Visible = True
               txtPrt(i).Visible = True
               lblDsc(i).Visible = True
               lblOrd(i).Visible = True
               txtShp(i).Visible = True
               chkShip(i).Visible = True
               lblUom(i).Visible = True
               txtAct(i).Visible = True
               txtAmt(i).Visible = True
               lblUnitPrice(i).Visible = True
               lblSchedDate(i).Visible = True
               
               lblitm(i) = Format(!ITNUMBER, "#0")
               lblRev(i) = "" & Trim(!itrev)
               txtPrt(i) = "" & Trim(!PARTNUM)
               lblUnitPrice(i) = Format("" & Trim(!ITDOLLORIG), "0.000")
               lblSchedDate(i) = Format("" & Trim(!ITSCHED), "mm/dd/yy")
               
               sItemCmts = "" & Trim(RdoItm!ITCOMMENTS)
               sPartExtDesc = "" & Trim(RdoItm!PAEXTDESC)
               If Len(Trim(sItemCmts)) = 0 Then sItemCmts = "Sales Order Item Comments"
               If Len(Trim(sPartExtDesc)) = 0 Then sPartExtDesc = "Extended Part Description"
            
               txtPrt(i).ToolTipText = sItemCmts
               lblDsc(i).ToolTipText = sPartExtDesc
               
               lblDsc(i) = "" & Trim(!PADESC)
               'lblOrd(i) = Format(!ITQTY, "####0.000")
               lblOrd(i) = !ITQTY
               txtShp(i) = "0.000"
               lblUom(i) = "" & Trim(!PAUNITS)
               txtAmt(i) = "0.0000"
            

               
               'don't allow this transaction to ship lot-tracked items
               txtShp(i).enabled = !PALOTTRACK = 0
               txtAct(i).enabled = !PALOTTRACK = 0
               txtAmt(i).enabled = !PALOTTRACK = 0
               chkShip(i).enabled = !PALOTTRACK = 0
               lblLotTracked(i).Visible = !PALOTTRACK <> 0
               
               ' if tax per line item is enabled show these controls
               lblStateTaxCode(i).Visible = taxPerLineItem
               lblStateTaxRate(i).Visible = taxPerLineItem
               lblStateTaxAmt(i).Visible = taxPerLineItem
               lblFedTaxCode(i).Visible = taxPerLineItem
               lblFedTaxRate(i).Visible = taxPerLineItem
               lblFedTaxAmt(i).Visible = taxPerLineItem
               If taxPerLineItem Then
                  lblStateTaxCode(i) = "" & Trim(!ITTAXCODE) & " " & Trim(!ITSLSTXACCT)
                  lblStateTaxRate(i) = Format(!ITTAXRATE, "0.0")
                  If lblStateTaxRate(i).Caption = "0.0" Then
                     lblStateTaxRate(i).Caption = ""
                  End If
                  lblStateTaxAmt(i) = Format(!ITTAXAMT, "0.00")
                  If lblStateTaxAmt(i).Caption = "0.00" Then
                     lblStateTaxAmt(i).Caption = ""
                  End If
                  lblFedTaxCode(i) = "" & Trim(!ITFEDTAXCODE) & " " & Trim(!ITFEDTAXACCT)
                  lblFedTaxRate(i) = Format(!ITFEDTAXRATE, "0.0")
                  lblFedTaxAmt(i) = Format(!ITFEDTAXAMT, "0.00")
               End If
            End If
            
            'If i = 1 Then a = iTotalItems
            
            iTotalItems = iTotalItems + 1
            
            
            vItems(iTotalItems, INV_ItemNumber) = Format(!ITNUMBER, "#0")
            vItems(iTotalItems, INV_ItemRev) = "" & !itrev
            vItems(iTotalItems, 3) = "" & !PARTNUM
            vItems(iTotalItems, 4) = "" & Trim(!PADESC)
            vItems(iTotalItems, INV_OrderQty) = Format(!ITQTY, "0.000")
            vItems(iTotalItems, INV_SchedDate) = Format(!ITSCHED, "mm/dd/yy")
            vItems(iTotalItems, INV_UnitCost) = Format(!ITDOLLORIG, "0.000")
            
            'If !ITCOMMENTS.ChunkRequired Then
            '   vItems(iTotalItems, 6) = "" & Trim(!ITCOMMENTS.GetChunk(3072))
            'Else
               vItems(iTotalItems, 6) = "" & Trim(!ITCOMMENTS)
            'End If
            vItems(iTotalItems, 7) = "" & Trim(!PAUNITS)
            vItems(iTotalItems, INV_ShipQty) = "0.000"
            vItems(iTotalItems, INV_ShipThisItem) = vbUnchecked
            vItems(iTotalItems, INV_PartRef) = "" & Trim(!PartRef)
            ''no longer used: vItems(iTotalItems, INV_Cost) = Format(!PASTDCOST, "0.000")
            
            'vItems(iTotalItems, INV_UnitPrice) = Format(!ITDOLLARS, CURRENCYMASK)
            vItems(iTotalItems, INV_UnitPrice) = Format(!ITDOLLARS, "0.000")
            
            vItems(iTotalItems, INV_FedTaxCode) = "" & Trim(!ITFEDTAXCODE)
            vItems(iTotalItems, INV_FedTaxAcct) = "" & Trim(!ITFEDTAXACCT)
            vItems(iTotalItems, INV_FedTaxRate) = Format(!ITFEDTAXRATE, "0.0")
            vItems(iTotalItems, INV_FedTaxAmt) = Format(!ITFEDTAXAMT, "0.00")
            
            vItems(iTotalItems, INV_SalesTaxCode) = "" & Trim(!ITTAXCODE)
            vItems(iTotalItems, INV_SalesTaxAcct) = "" & Trim(!ITSLSTXACCT)
            vItems(iTotalItems, INV_SalesTaxRate) = Format(!ITTAXRATE, "0.0")
            
            vItems(iTotalItems, INV_SalesTaxAmt) = Format(!ITTAXAMT, "0.00")
            
            vItems(iTotalItems, 14) = Format(!PALEVEL, "0")
            vItems(iTotalItems, 15) = "" & Trim(!PAPRODCODE)
            vItems(iTotalItems, INV_TaxExempt) = !PATAXEXEMPT
            vItems(iTotalItems, INV_LotTracked) = !PALOTTRACK
            
            sPartRevAcct = ""
            
            bByte = GetPartInvoiceAccounts(CStr(vItems(iTotalItems, INV_PartRef)), CInt(vItems(iTotalItems, 14)), _
                    CStr(vItems(iTotalItems, 15)), sPartRevAcct)
            
            vItems(iTotalItems, INV_PartRevAccount) = sPartRevAcct
            
            ' sloppy will fix later (nth)
            If i < 6 Then
               txtAct(i) = sPartRevAcct
            End If
            .MoveNext
            
         Loop
         .Cancel
         GetItems = True
      End With
      iCurrPage = 1
      iMaxPages = ((iTotalItems + 1) / 5) + 0.45
      
      lblPge = iCurrPage
      lblLst = iMaxPages
      'If iTotalItems > 5 Then
      If (iTotalItems + 1) > 5 Then
         cmdDn.Picture = Endn
         cmdDn.enabled = True
      End If
      iNewBookmark = 0
      On Error Resume Next
      optAll.SetFocus
   Else
      GetItems = False
   End If
   MouseCursor 0
   Set RdoItm = Nothing
   Exit Function
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Public Sub CloseBoxes(bNotFirstBox As Byte)
   Dim i As Integer
   Dim A As Integer
   On Error GoTo DiaErr1
   If bNotFirstBox Then A = 1 Else A = 2
   '        For i = a To 4
   '            lblitm(i).Visible = False
   '            lblRev(i).Visible = False
   '            lblPrt(i).Visible = False
   '            lblDsc(i).Visible = False
   '            lblOrd(i).Visible = False
   '            txtShp(i).Visible = False
   '            lblUom(i).Visible = False
   '            chkShip(i).Visible = False
   '            txtAmt(i).Visible = False
   '            txtAct(i).Visible = False
   '
   '
   '
   '        Next
   
   '        why do the last iteration separately?
   '        lblitm(i).Visible = False
   '        lblRev(i).Visible = False
   '        lblPrt(i).Visible = False
   '        lblDsc(i).Visible = False
   '        lblOrd(i).Visible = False
   '        txtShp(i).Visible = False
   '        lblUom(i).Visible = False
   '        chkShip(i).Visible = False
   '        txtAmt(i).Visible = False
   '        txtAct(i).Visible = False
   
   For i = A To 5
      lblitm(i).Visible = False
      lblRev(i).Visible = False
      txtPrt(i).Visible = False
      lblDsc(i).Visible = False
      lblOrd(i).Visible = False
      txtShp(i).Visible = False
      lblUom(i).Visible = False
      chkShip(i).Visible = False
      txtAmt(i).Visible = False
      lblUnitPrice(i).Visible = False
      lblSchedDate(i).Visible = False
      txtAct(i).Visible = False
      
      lblStateTaxCode(i).Visible = False
      lblStateTaxRate(i).Visible = False
      lblStateTaxAmt(i).Visible = False
      lblFedTaxCode(i).Visible = False
      lblFedTaxRate(i).Visible = False
      lblFedTaxAmt(i).Visible = False
      
   Next
   
   For i = 1 To 5
      lblitm(i) = ""
      lblRev(i) = ""
      txtPrt(i) = ""
      lblDsc(i) = ""
      lblOrd(i) = ""
      txtShp(i) = ""
      lblUom(i) = ""
      'chkShip(i) = vbUnchecked
      txtAmt(i) = ""
      lblUnitPrice(i) = ""
      lblSchedDate(i) = ""
      txtAct(i) = ""
      
      lblStateTaxCode(i) = ""
      lblStateTaxRate(i) = ""
      lblStateTaxAmt(i) = ""
      lblFedTaxCode(i) = ""
      lblFedTaxRate(i) = ""
      lblFedTaxAmt(i) = ""
   Next
   
   ' why?
   '        lblitm(i) = ""
   '        lblRev(i) = ""
   '        lblPrt(i) = ""
   '        lblDsc(i) = ""
   '        lblOrd(i) = ""
   '        txtShp(i) = ""
   '        lblUom(i) = ""
   '        'chkShip(i) = vbUnchecked
   '        txtAmt(i) = ""
   '        txtAct(i) = ""
   Exit Sub
DiaErr1:
   sProcName = "closeboxes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Public Sub GetNextGroup()
   Dim i As Integer
   Dim A As Integer
   CloseBoxes False
   
   On Error GoTo DiaErr1
   If iCurrPage < 1 Then iCurrPage = 1
   
   If iCurrPage > 1 Then
      cmdUp.Picture = Enup
      cmdUp.enabled = True
   Else
      cmdUp.Picture = Dsup
      cmdUp.enabled = False
   End If
   
   If iCurrPage >= iMaxPages Then
      iCurrPage = iMaxPages
      cmdDn.Picture = Dsdn
      cmdDn.enabled = False
   Else
      cmdDn.Picture = Endn
      cmdDn.enabled = True
   End If
   
   
   
   lblPge = iCurrPage
   A = 0
   iNewBookmark = (5 * iCurrPage) - 5
   For i = iNewBookmark To iTotalItems
      A = A + 1
      If A > 5 Then Exit For
      lblitm(A).Visible = True
      lblRev(A).Visible = True
      txtPrt(A).Visible = True
      lblDsc(A).Visible = True
      lblOrd(A).Visible = True
      txtShp(A).Visible = True
      chkShip(A).Visible = True
      lblUom(A).Visible = True
      txtAmt(A).Visible = True
      lblUnitPrice(A).Visible = True
      lblSchedDate(A).Visible = True
      txtAct(A).Visible = True
      
      lblStateTaxCode(A).Visible = taxPerLineItem
      lblStateTaxRate(A).Visible = taxPerLineItem
      lblStateTaxAmt(A).Visible = taxPerLineItem
      lblFedTaxCode(A).Visible = taxPerLineItem
      lblFedTaxRate(A).Visible = taxPerLineItem
      lblFedTaxAmt(A).Visible = taxPerLineItem
      
      lblitm(A) = Format(vItems(i, INV_ItemNumber), "#0")
      lblRev(A) = vItems(i, INV_ItemRev)
      txtPrt(A) = vItems(i, 3)
      lblDsc(A) = vItems(i, 4)
      lblOrd(A) = vItems(i, INV_OrderQty)
      lblSchedDate(A) = vItems(i, INV_SchedDate)
      Me.lblUnitPrice(A) = vItems(i, INV_UnitCost)
      
      txtShp(A) = vItems(i, INV_ShipQty)
      lblUom(A) = vItems(i, 7)
      chkShip(A).Value = vItems(i, INV_ShipThisItem)
      txtAct(A) = vItems(i, INV_PartRevAccount)
      txtAmt(A) = Format(vItems(i, INV_UnitPrice) * vItems(i, INV_ShipQty), "####0.0000")
      'here
      
      If taxPerLineItem Then
         lblStateTaxCode(A) = vItems(i, INV_SalesTaxCode) & " " & vItems(i, INV_SalesTaxAcct)
         lblStateTaxRate(A) = vItems(i, INV_SalesTaxRate)
         If lblStateTaxRate(A).Caption = "0.0" Then
            lblStateTaxRate(A).Caption = ""
         End If
         lblStateTaxAmt(A) = vItems(i, INV_SalesTaxAmt)
         If lblStateTaxAmt(A).Caption = "0.00" Then
            lblStateTaxAmt(A).Caption = ""
         End If
         lblFedTaxCode(A) = vItems(i, INV_FedTaxCode) & " " & vItems(i, INV_FedTaxAcct)
         lblFedTaxRate(A) = vItems(i, INV_FedTaxRate)
         lblFedTaxAmt(A) = vItems(i, INV_FedTaxAmt)
      End If
      
      'don't allow this transaction to ship lot-tracked items
      txtShp(A).enabled = vItems(i, INV_LotTracked) = 0
      txtAct(A).enabled = vItems(i, INV_LotTracked) = 0
      txtAmt(A).enabled = vItems(i, INV_LotTracked) = 0
      chkShip(A).enabled = vItems(i, INV_LotTracked) = 0
      lblLotTracked(A).Visible = vItems(i, INV_LotTracked) = 1
      
   Next
   Exit Sub
DiaErr1:
   sProcName = "getnextgroup"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub optAll_Click()
   Dim i As Integer
   If optAll.Value = vbChecked Then
      For i = 0 To iTotalItems
         vItems(i, INV_ShipQty) = vItems(i, INV_OrderQty)
         If vItems(i, INV_LotTracked) = 0 Then
            vItems(i, INV_ShipThisItem) = vbChecked
         End If
      Next
      For i = 1 To 5
         If i + iNewBookmark <= iTotalItems + 1 Then
            If vItems(i - 1, INV_LotTracked) = 0 And txtShp(i).Visible Then
               txtShp(i) = vItems(iNewBookmark + (i - 1), INV_ShipQty)
               chkShip(i) = vbChecked
            End If
            UpdateLineItemInfo iNewBookmark + i - 1, i
         End If
      Next
   Else
      
      For i = 0 To iTotalItems
         vItems(i, INV_ShipQty) = "0.000"
         vItems(i, INV_ShipThisItem) = vbUnchecked
      Next
      For i = 1 To 5
         If i + iNewBookmark <= iTotalItems + 1 Then
            txtShp(i) = vItems(iNewBookmark + (i - 1), INV_ShipQty)
            chkShip(i) = vbUnchecked
            UpdateLineItemInfo iNewBookmark + i - 1, i
         End If
      Next
   End If
   UpdateTotals
End Sub

Private Sub optAll_LostFocus()
   '
End Sub


Private Sub chkShip_Click(Index As Integer)
   If chkShip(Index) = vbChecked Then
      If Val(txtShp(Index)) = 0 Then txtShp(Index) = lblOrd(Index)
   Else
      txtShp(Index) = "0.000"
   End If
   
   vItems(iNewBookmark + (Index - 1), INV_ShipThisItem) = chkShip(Index).Value
   vItems(iNewBookmark + (Index - 1), INV_ShipQty) = txtShp(Index).Text
   
   'txtAmt(Index) = Format((vItems(iNewBookmark + (Index - 1), INV_UnitPrice) * Val(txtShp(Index))), "0.00")
   UpdateLineItemInfo iNewBookmark + Index - 1, Index
   UpdateTotals
End Sub

Private Sub chkShip_KeyPress(Index As Integer, KeyAscii As Integer)
   'KeyLock KeyAscii
   
End Sub

Private Sub chkShip_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub chkShip_LostFocus(Index As Integer)
   'vItems(iNewBookmark + (Index - 1), INV_ShipThisItem) = chkShip(Index).Value
End Sub

Private Sub txtAct_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtAct_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtAct_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub txtAct_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtAct_LostFocus(Index As Integer)
   Dim b As Byte
   Dim i As Integer
   txtAct(Index) = CheckLen(txtAct(Index), 12)
   If txtAct(Index).ListCount > 0 Then
      For i = 0 To txtAct(Index).ListCount - 1
         If txtAct(Index).List(i) = txtAct(Index) Then b = 1
      Next
      If b = 0 Then
         If Len(sAccount) > 0 Then
            txtAct(Index) = sAccount
         Else
            txtAct(Index) = txtAct(Index) = txtAct(Index).List(0)
         End If
      End If
   End If
   vItems(iNewBookmark + (Index - 1), INV_PartRevAccount) = txtAct(Index)
   If Len(txtAct(Index)) Then sAccount = txtAct(Index)
   
End Sub


Private Sub txtShp_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtShp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtShp_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtShp_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub txtShp_LostFocus(Index As Integer)
   
   Dim Subscript As Integer
   Subscript = iNewBookmark + (Index - 1)
   txtShp(Index) = CheckLen(txtShp(Index), 9)
   txtShp(Index) = Format(Abs(Val(txtShp(Index))), "####0.0000")
   
   ' move down: txtAmt(Index) = Format((vItems(iNewBookmark + (Index - 1), INV_UnitPrice) * Val(txtShp(Index))), "#####0.00")
   
   If Val(txtShp(Index)) > Val(lblOrd(Index)) Then txtShp(Index) = lblOrd(Index)
   
   '    Dim price As Currency
   '    price = CCur(vItems(subscript, INV_UnitPrice)) * CCur(txtShp(Index))
   '    txtAmt(Index) = Format(price, "#####0.00")
   '
   '    If taxPerLineItem Then
   '        If Val(vItems(subscript, INV_SalesTaxRate)) > 0 Then
   '            vItems(subscript, INV_SalesTaxAmt) _
   '                = Format(price * Val(vItems(subscript, INV_SalesTaxRate)) / 100, "0.00")
   '            lblStateTaxAmt(Index) = vItems(subscript, INV_SalesTaxAmt)
   '        End If
   '
   '        If Val(vItems(subscript, INV_FedTaxRate)) > 0 Then
   '            vItems(subscript, INV_FedTaxAmt) _
   '                = Format(price * Val(vItems(subscript, INV_FedTaxRate)) / 100, "0.00")
   '            lblFedTaxAmt(Index) = vItems(subscript, INV_FedTaxAmt)
   '        End If
   '
   '    End If
   
   
   vItems(Subscript, INV_ShipQty) = txtShp(Index)
   
   If Val(txtShp(Index)) > 0 Then
      chkShip(Index) = vbChecked
   Else
      chkShip(Index) = vbUnchecked
   End If
   
   vItems(Subscript, INV_ShipThisItem) = chkShip(Index).Value
   
   UpdateLineItemInfo Subscript, Index
   
   UpdateTotals
End Sub

Private Sub UpdateLineItemInfo(arraySub As Integer, controlSub As Integer)
   Dim shipQty As Currency, price As Currency
   shipQty = CCur("0" & txtShp(controlSub))
   price = CCur(vItems(arraySub, INV_UnitPrice)) * shipQty
   txtAmt(controlSub) = Format(price, "#####0.0000")
   
   'calc taxes per line item if required
   'note addition of .0001 to make rounding work.  Otherwise, .5 rounds down (unbelievable)
   If taxPerLineItem Then
      If vItems(arraySub, INV_ShipThisItem) = vbChecked Then
         If Val(vItems(arraySub, INV_SalesTaxRate)) > 0 Then
            vItems(arraySub, INV_SalesTaxAmt) _
                   = Format(Round((price * Val(vItems(arraySub, INV_SalesTaxRate)) + 0.0001) / 100, 2), "0.00")
         End If
         
         If Val(vItems(arraySub, INV_FedTaxRate)) > 0 Then
            vItems(arraySub, INV_FedTaxAmt) _
                   = Format(Round((price * Val(vItems(arraySub, INV_FedTaxRate)) + 0.0001) / 100, 2), "0.00")
         End If
         
      Else
         If Val(vItems(arraySub, INV_SalesTaxRate)) > 0 Then
            vItems(arraySub, INV_SalesTaxAmt) = "0.00"
         End If
         If Val(vItems(arraySub, INV_FedTaxRate)) > 0 Then
            vItems(arraySub, INV_FedTaxAmt) = "0.00"
         End If
      End If
      
      lblStateTaxAmt(controlSub) = vItems(arraySub, INV_SalesTaxAmt)
      lblFedTaxAmt(controlSub) = vItems(arraySub, INV_FedTaxAmt)
      
   End If
End Sub

Public Sub CreateInvoice()
   Dim bByte As Byte
   Dim i As Integer
   Dim A As Integer
   Dim K As Integer
   
   Dim iItemNumber As Integer
   
   'Dim iTrans As Integer
   'Dim iRef As Integer
   
   Dim nLDollars As Single
   Dim nTdollars As Single
   Dim sComments As String
   Dim sDate As String
   
   Dim sPart As String
   Dim sProd As String
   Dim iLevel As Integer
   Dim nItemNo As Integer
   Dim sItemRev As String
   
   
   Dim sRev As String
   Dim sPartCgsAcct As String
   Dim sPartRevAcct As String
   Dim sPartDisAcct As String
   Dim cSalesTax As Currency
   Dim cFREIGHT As Currency
   
   ' BnO Taxes
   Dim nRate As Single
   Dim sType As String
   Dim sState As String
   Dim sCode As String
   
   Dim lLastActivity As Long
   
   Dim rdoCst As ADODB.Recordset ' Standard Cost
   
   On Error GoTo DiaErr1
   ' Any selected?
   For i = 0 To iTotalItems
      If Val(vItems(i, INV_ShipQty)) > 0 Then
         bByte = True
      End If
   Next
   If Not bByte Then
      MouseCursor 0
      MsgBox "There Are No Items Selected.", vbInformation, Caption
      On Error Resume Next
      txtShp(1).SetFocus
      Exit Sub
   End If
   diaARe01b.optSav.Value = vbChecked
   
   ' Accounts ?
   bByte = True
   For i = 1 To iTotalItems
      If Val(vItems(i, INV_ShipQty)) > 0 Then
         sPart = vItems(i, INV_PartRef)
         iLevel = Val(vItems(i, 14))
         sProd = vItems(i, 15)
         bByte = GetPartInvoiceAccounts(sPart, iLevel, sProd, _
                 sPartRevAcct, sPartCgsAcct)
      End If
      If Not bByte Then Exit For
   Next
   
   ' None or not all acounts found.
   If Not bByte Then
      sMsg = "Could Not Find Either A Revenue Or A CGS Account " & vbCr _
             & "For One More Of The Selected Part Numbers."
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   
   MouseCursor 13
   
   'get invoice date to use with debits and credits
   'sDate = Format(GetServerDateTime(), "mm/dd/yy")
   sSql = "SELECT INVDATE from CihdTable" & vbCrLf _
          & "where INVNO=" & lCurrInvoice
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst)
   If bSqlRows Then
      With rdoCst
         sDate = Format(!INVDATE, "mm/dd/yy")
         .Cancel
      End With
   End If
   Set rdoCst = Nothing
   
   'On Error Resume Next
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   lLastActivity = GetLastActivity() + 1
   
   ' get ready to create debits and credits
   Dim gl As New GLTransaction
   gl.JournalID = Trim(diaARe01b.sJournalID)
   gl.InvoiceDate = CDate(sDate)
   gl.InvoiceNumber = lInv
   
   ' Add all selected items to invoice.
   For i = 0 To iTotalItems
      If Val(vItems(i, INV_ShipQty)) > 0 And Val(vItems(i, INV_ShipThisItem)) = 1 Then
         iItemNumber = iItemNumber + 1
         nLDollars = (CCur(vItems(i, INV_UnitPrice)) * Val(vItems(i, INV_ShipQty)))
         nTdollars = nTdollars + nLDollars
         
         ' If not invoicing full quantity, split SO item
         Dim cRemainingQty As Currency
         cRemainingQty = Val(vItems(i, INV_OrderQty)) - Val(vItems(i, INV_ShipQty))
         Dim cRemainingTotal As Currency
         cRemainingTotal = cRemainingQty * Val(vItems(i, INV_UnitPrice))
         
         If cRemainingQty > 0 Then
            sRev = Trim(vItems(i, INV_ItemRev))
            If sRev = "" Then
               sRev = "A"
            Else
               A = Asc(sRev)
               sRev = Chr(A + 1)
            End If
            
            'create a new open item
            sSql = "insert into SoitTable" & vbCrLf _
                   & "(" & vbCrLf _
                   & "ITREV,ITQTY," & vbCrLf _
                   & "ITSO,ITNUMBER,ITPART,ITCUSTREQ,ITCOMMISSION,ITCOST,ITINV,ITDOLLARS,ITSCHED," & vbCrLf _
                   & "ITACTUAL,ITINTSPAR1,ITADJUST,ITORIGQTY,ITDISCRATE,ITDISCAMOUNT,ITCGSACCT,ITINVACCT," & vbCrLf _
                   & "ITDISACCT,ITTFRPRICE, ITTAXCODE,ITSLSTXACCT,ITTAXRATE,ITTAXAMT, ITSTATE,ITQTYORIG," & vbCrLf _
                   & "ITDOLLORIG,ITDISCRORIG,ITDISCAORIG,ITTAXRORIG,ITTAXAORIG,ITBOSTATIC,ITBOSTATE,ITBOCODE," & vbCrLf _
                   & "ITCOSTED,ITQUOTEPROB,ITPSNUMBER,ITPSITEM,ITPSCARTON,ITPSSHIPNO,ITFRTALLOW,ITSCHEDDEL," & vbCrLf _
                   & "ITBOOKDATE,ITCREATED,ITREVISED,ITCANCELDATE,ITCANCELED,ITUSER,ITCANCELEDBY,ITCOMMENTS," & vbCrLf _
                   & "ITINVOICE , ITSERNO, ITREVACCT, ITLOTNUMBER, ITCUSTITEMNO, ITPSSHIPPED, ITFEDTAXCODE, ITFEDTAXACCT, ITFEDTAXRATE, ITFEDTAXAMT" & vbCrLf _
                   & ")" & vbCrLf _
                   & "select" & vbCrLf _
                   & "'" & sRev & "'," & cRemainingQty & "," & vbCrLf _
                   & "ITSO,ITNUMBER,ITPART,ITCUSTREQ,ITCOMMISSION,ITCOST,ITINV,ITDOLLARS,ITSCHED," & vbCrLf _
                   & "ITACTUAL,ITINTSPAR1,ITADJUST,ITORIGQTY,ITDISCRATE,ITDISCAMOUNT,ITCGSACCT,ITINVACCT," & vbCrLf _
                   & "ITDISACCT,ITTFRPRICE, ITTAXCODE,ITSLSTXACCT,ITTAXRATE," & cRemainingTotal & " * ITTAXRATE / 100, ITSTATE,ITQTYORIG," & vbCrLf _
                   & "ITDOLLORIG,ITDISCRORIG,ITDISCAORIG,ITTAXRORIG,ITTAXAORIG,ITBOSTATIC,ITBOSTATE,ITBOCODE," & vbCrLf _
                   & "ITCOSTED,ITQUOTEPROB,ITPSNUMBER,ITPSITEM,ITPSCARTON,ITPSSHIPNO,ITFRTALLOW,ITSCHEDDEL," & vbCrLf _
                   & "ITBOOKDATE,ITCREATED,ITREVISED,ITCANCELDATE,ITCANCELED,ITUSER,ITCANCELEDBY,ITCOMMENTS," & vbCrLf _
                   & "ITINVOICE , ITSERNO, ITREVACCT, ITLOTNUMBER, ITCUSTITEMNO, ITPSSHIPPED, ITFEDTAXCODE, ITFEDTAXACCT, ITFEDTAXRATE, " & vbCrLf _
                   & cRemainingTotal & " * ITFEDTAXRATE / 100 " & vbCrLf _
                   & "From SoitTable" & vbCrLf _
                   & "WHERE " _
                   & "ITSO=" & lSalesOrder & " AND ITNUMBER=" & vItems(i, INV_ItemNumber) & " AND " _
                   & "ITREV='" & vItems(i, INV_ItemRev) & "' "
            clsADOCon.ExecuteSQL sSql
            
            'Dim rdo As ADODB.RecordSet
            
         End If
         
         ' Get part accounts
         sPart = vItems(i, INV_PartRef)
         iLevel = Val(vItems(i, 14))
         sProd = vItems(i, 15)
         bByte = GetPartInvoiceAccounts(sPart, iLevel, sProd, sPartRevAcct, sPartDisAcct, sPartCgsAcct)
         nItemNo = CInt(vItems(i, INV_ItemNumber))
         sItemRev = Trim(CStr(vItems(i, INV_ItemRev)))
         
         ' BnO tax
         sCode = ""
         nRate = 0
         sState = ""
         sType = ""
         GetPartBnO vItems(i, 3), nRate, sCode, sState, sType
         If sCode = "" Then
            GetCustBnO sSOCust, nRate, sCode, sState, sType
         End If
         
         ' if tax is calculated per line item, use existing imported amounts
         Dim salesTaxAmt As Currency
         
         Dim fedTaxCode As String
         Dim fedTaxAcct As String
         Dim fedTaxRate As Currency
         Dim fedTaxAmt As Currency
         
         sTaxCode = Compress(sTaxCode)
         sTaxAcct = Compress(sTaxAcct)
         
         If taxPerLineItem Then
            sTaxCode = vItems(i, INV_SalesTaxCode)
            sTaxAcct = vItems(i, INV_SalesTaxAcct)
            nTaxRate = vItems(i, INV_SalesTaxRate)
            
            fedTaxCode = vItems(i, INV_FedTaxCode)
            fedTaxAcct = vItems(i, INV_FedTaxAcct)
            fedTaxRate = vItems(i, INV_FedTaxRate)
            salesTaxAmt = vItems(i, INV_SalesTaxAmt)
            fedTaxAmt = vItems(i, INV_FedTaxAmt)
         Else
            fedTaxCode = ""
            fedTaxAcct = ""
            fedTaxRate = 0
            salesTaxAmt = Round((nTaxRate / 100) * nLDollars, 2)
            fedTaxAmt = Round((fedTaxRate * nLDollars) / 100, 2)
         End If
         
         ' update shipped sales order item
         sSql = "UPDATE SoitTable SET " _
                & "ITINVOICE=" & lInv & "," _
                & "ITQTY=" & vItems(i, INV_ShipQty) & "," _
                & "ITCGSACCT='" & Compress(sPartCgsAcct) & "'," _
                & "ITREVACCT='" & Compress(sPartRevAcct) & "'," _
                & "ITACTUAL='" & sDate & "'," _
                & "ITBOSTATE='" & sState & "'," _
                & "ITBOCODE='" & sCode & "'," _
                & "ITTAXCODE='" & sTaxCode & "'," _
                & "ITSLSTXACCT='" & sTaxAcct & "'," _
                & "ITTAXRATE=" & nTaxRate & "," _
                & "ITTAXAMT=" & salesTaxAmt & "," & vbCrLf _
                & "ITSTATE='" & sTaxState & "'," _
                & "ITFEDTAXCODE='" & fedTaxCode & "'," _
                & "ITFEDTAXACCT='" & fedTaxAcct & "'," _
                & "ITFEDTAXRATE=" & fedTaxRate & "," _
                & "ITFEDTAXAMT=" & fedTaxAmt _
                & " WHERE ITSO=" & lSalesOrder _
                & " AND ITNUMBER=" & vItems(i, INV_ItemNumber) _
                & " AND ITREV='" & Trim(sItemRev) & "' "
         'Debug.Print sSql
         clsADOCon.ExecuteSQL sSql
         
         ' Update part QOH
         sSql = "UPDATE PartTable SET PAQOH=PAQOH-" & Val(vItems(i, INV_ShipQty)) & " " _
                & "WHERE PARTREF='" & sPart & "' "
         clsADOCon.ExecuteSQL sSql
         
         AverageCost Trim(sPart)
         GetAccounts i
         
         Dim iLots As Integer
         Dim lCOUNTER As Long
         Dim lLOTRECORD As Long
         
         Dim cItmLot As Currency
         Dim cLotQty As Currency
         Dim cPartCost As Currency
         Dim cPckQty As Currency
         Dim cQuantity As Currency
         Dim cRemPQty As Currency
         
         Dim sLot As String
         
         '0 = Lot Number
         '1 = Lot Quantity
         
         cRemPQty = Val(vItems(i, INV_ShipQty))
         ' If part is lot tracked allow user to select lots.
         If vItems(i, INV_LotTracked) = 1 Then
         
'            '***** real lots
'            cLotQty = GetRemainingLotQty(sPart, True)
'            If cLotQty < cQuantity Then
'               MsgBox "The Lot Quantity of The Item Is Less Than Required.", _
'                  vbInformation, Caption
'               bLotFailed = 1
'            Else
'               MsgBox "Lot Tracking Is Required For This Part Number." & vbCrLf _
'                  & "You Must Select The Lot(s) To Use For Shipping.", _
'                  vbInformation, Caption
'               'Get The lots
'               LotSelect.lblPart = sPart
'               LotSelect.lblRequired = cQuantity
'               LotSelect.Show vbModal
'               If Es_TotalLots > 0 Then
'                  For A = 1 To UBound(lots)
'                     'insert lot transaction here
'                     lLOTRECORD = GetNextLotRecord(lots(A).LotSysId)
'                     sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
'                            & "LOITYPE,LOIPARTREF,LOIADATE,LOIQUANTITY," _
'                            & "LOICUSTINVNO,LOICUST,LOIACTIVITY,LOICOMMENT) " _
'                            & "VALUES('" & lots(A).LotSysId & "'," & lLOTRECORD _
'                            & "," & IATYPE_ShippedSoItem & ",'" & sPart & "','" & vAdate & "',-" _
'                            & lots(A).LotSelQty & ",'" & lInv & "','" _
'                            & Compress(lblCst) & "'," & lCOUNTER & ",'Shipped')"
'                     clsAdoCon.ExecuteSQL sSql
'
'                     sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY" _
'                            & "-" & lots(A).LotSelQty & " WHERE LOTNUMBER='" _
'                            & lots(A).LotSysId & "'"
'                     clsAdoCon.ExecuteSQL sSql
'                     cItmLot = cItmLot + lots(A).LotSelQty
'
'                     Dim cost As New CostInfo
'                     cost.PartRef = sPart
'                     sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
'                            & "INADATE,INAQTY," & vbCrLf _
'                            & "INAMT,INTOTMATL,INTOTLABOR,INTOTEXP,INTOTOH,INTOTHRS," & vbCrLf _
'                            & "INSONUMBER,INSOITEM,INSOREV," _
'                            & "INCREDITACCT,INDEBITACCT,INLOTNUMBER, INNUMBER) " _
'                            & "VALUES(" & IATYPE_ShippedSoItem & ",'" & sPart & "','SHIPMENT','SO " _
'                            & lblSon & " INV " & lInv & "','" & sDate _
'                            & "',-" & Abs(cItmLot) & "," _
'                            & cost.Amt & "," & cost.Matl & "," & cost.Labor & "," & vbCrLf _
'                            & cost.Exp & "," & cost.OH & "," & cost.Hrs & "," & vbCrLf _
'                            & lSalesOrder & "," _
'                            & Val(vItems(i, INV_ItemNumber)) & ",'" & sItemRev & "','" _
'                            & Compress(sCreditAcct) & "','" & Compress(sDebitAcct) & "','" & lots(A).LotSysId & "'," & lLastActivity & ")"
'
'                     Debug.Print sSql
'
'                     clsAdoCon.ExecuteSQL sSQL
'                     lLastActivity = lLastActivity + 1
'
'                  Next
'               Else
'                  MsgBox "Lot Quantity Failed And Resetting.", _
'                     vbInformation, Caption
'                  bLotFailed = 1
'               End If
'            End If
         
         Else
            
            ' Update the LohdTable and LoitTable
            iLots = GetPartLots(sPart)
            cItmLot = 0
            If iLots > 0 Then
               For A = 1 To iLots
                  cLotQty = Val(sLots(A, 1))
                  If cLotQty >= cRemPQty Then
                     cPckQty = cRemPQty
                     cLotQty = cLotQty - cRemPQty
                     cRemPQty = 0
                  Else
                     cPckQty = cLotQty
                     cRemPQty = cRemPQty - cLotQty
                     cLotQty = 0
                  End If
                  cItmLot = cItmLot + cPckQty
                  If cItmLot > Val(sLots(A, 1)) Then cItmLot = Val(sLots(A, 1))
                  sLot = sLots(A, 0)
                  lLOTRECORD = GetNextLotRecord(sLot)
                  lInv = Val(Right(lblInv, 6))

                  'insert lot transaction here
                  sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                         & "LOITYPE,LOIPARTREF,LOIQUANTITY," _
                         & "LOICUSTINVNO,LOICUST," _
                         & "LOIACTIVITY,LOICOMMENT) " _
                         & "VALUES('" & sLots(A, 0) & "'," _
                         & lLOTRECORD & "," & IATYPE_ShippedSoItem & ",'" & sPart & "',-" _
                         & Abs(cItmLot) & ",'" & lInv & "','" _
                         & Compress(lblCst) & "'," & lCOUNTER & ",'Shipped')"
                  clsADOCon.ExecuteSQL sSql
                  
                  'Update Lot Header
                  sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY" _
                         & "-" & Abs(cItmLot) & " WHERE LOTNUMBER='" & sLots(A, 0) & "'"
                  clsADOCon.ExecuteSQL sSql
                  
                  ' No lots just add to activity
                  Dim cost As New CostInfo
                  cost.PartRef = sPart
                  sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                         & "INADATE,INAQTY," & vbCrLf _
                         & "INAMT,INTOTMATL,INTOTLABOR,INTOTEXP,INTOTOH,INTOTHRS," & vbCrLf _
                         & "INSONUMBER,INSOITEM,INSOREV," _
                         & "INCREDITACCT,INDEBITACCT,INLOTNUMBER, INNUMBER) " _
                         & "VALUES(" & IATYPE_ShippedSoItem & ",'" & sPart & "','SHIPMENT','SO " _
                         & lblSon & " INV " & lInv & "','" & sDate _
                         & "',-" & Abs(cItmLot) & "," _
                         & cost.Amt & "," & cost.Matl & "," & cost.Labor & "," & vbCrLf _
                         & cost.Exp & "," & cost.OH & "," & cost.Hrs & "," & vbCrLf _
                         & lSalesOrder & "," _
                         & Val(vItems(i, INV_ItemNumber)) & ",'" & sItemRev & "','" _
                         & Compress(sCreditAcct) & "','" & Compress(sDebitAcct) & "','" & sLots(A, 0) & "'," & lLastActivity & ")"
                  
                  Debug.Print sSql
                  
                  clsADOCon.ExecuteSQL sSql
                  lLastActivity = lLastActivity + 1
                  
                  
                  If cRemPQty <= 0 Then Exit For
               Next
            End If
            
         End If
         
         ' Debit A/R (+)
         gl.AddDebitCredit CCur(nLDollars), 0, Compress(sCOSjARAcct), sPart, _
                                lSalesOrder, nItemNo, sItemRev, sSOCust
         
         ' Credit Revenue (-)
         gl.AddDebitCredit 0, CCur(nLDollars), Compress(Trim(vItems(i, INV_PartRevAccount))), _
                                   sPart, lSalesOrder, nItemNo, sItemRev, sSOCust
         
      End If
   Next
   
   cFREIGHT = CCur(lblFrt)
   cSalesTax = CCur(lblTax)
   Dim cFedTax As Currency
   cFedTax = CCur(lblFedTax)
   nTdollars = nTdollars + (cFREIGHT + cSalesTax + cFedTax)
   
   ' Update customer invoice.
   sSql = "UPDATE CihdTable SET " _
          & "INVTOTAL=INVTOTAL + " & nTdollars & "," _
          & "INVTAX=INVTAX + " & cSalesTax & "," _
          & "INVFREIGHT=" & cFREIGHT & "," _
          & "INVFEDTAX=INVFEDTAX + " & cFedTax _
          & " WHERE INVNO=" & lCurrInvoice & " "
   clsADOCon.ExecuteSQL sSql
   
   If cFREIGHT > 0 Then
      ' Debit A/R Freight
      gl.AddDebitCredit cFREIGHT, 0, Compress(sCOSjARAcct), "", _
                                              lSalesOrder, 0, 0, sSOCust
      
      ' Credit Freight
      gl.AddDebitCredit 0, cFREIGHT, Compress(sCOSjNFRTAcct), "", _
                                              lSalesOrder, 0, 0, sSOCust
   End If
   
   If cSalesTax > 0 Then
      ' Debit A/R Taxes
      gl.AddDebitCredit cSalesTax, 0, Compress(sCOSjARAcct), "", _
                                               lSalesOrder, 0, 0, sSOCust
      
      ' Credit Taxes
      gl.AddDebitCredit 0, cSalesTax, Compress(sCOSjTaxAcct), "", _
                                               lSalesOrder, 0, 0, sSOCust
   End If
   
   If cFedTax <> 0 Then
      ' Debit A/R Taxes
      gl.AddDebitCredit cFedTax, 0, Compress(sCOSjARAcct), "", _
                                             lSalesOrder, 0, 0, sSOCust
      
      ' Credit Fed Taxes
      gl.AddDebitCredit 0, cFedTax, Compress(sCOSjFedTaxAcct), "", _
                                             lSalesOrder, 0, 0, sSOCust
   End If
   
   MouseCursor 0
   
   'test that debits and credits balance
   Dim success As Boolean
   success = gl.Commit
   
   If success And clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      bInvSaved = True
      SysMsg "Invoiced Items Recorded.", 1
      Unload Me
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      bInvSaved = False
      MsgBox "Invoice Could Not Be Recorded.", vbExclamation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "createinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub FillAccounts()
   Dim rdoAct As ADODB.Recordset
   Dim i As Integer
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillLowAccounts"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct)
   'No Accounts?
   If Err = 0 Then
      If bSqlRows Then
         With rdoAct
            Do Until .EOF
               For i = 1 To 4
                  'txtAct(i).AddItem "" & Trim(!GLACCTNO)
                  AddComboStr txtAct(i).hWnd, "" & Trim(!GLACCTNO)
               Next
               'txtAct(i).AddItem "" & Trim(!GLACCTNO)
               AddComboStr txtAct(i).hWnd, "" & Trim(!GLACCTNO)
               .MoveNext
            Loop
         End With
         If txtAct(1).ListCount > 0 Then
            If sAccount = "" Then
               For i = 1 To 4
                  txtAct(i) = txtAct(i).List(0)
               Next
               txtAct(i) = txtAct(i).List(0)
            Else
               For i = 1 To 4
                  txtAct(i) = sAccount
               Next
               txtAct(i) = sAccount
            End If
         End If
      End If
   End If
   Set rdoAct = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "fillaccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub GetAccounts(iIndex As Integer)
   Dim rdoAct As ADODB.Recordset
   Dim bType As Byte
   Dim sPcode As String
   
   On Error GoTo DiaErr1
   'Use current Part
   sSql = "Qry_GetExtPartAccounts '" & vItems(iIndex, INV_PartRef) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         sPcode = "" & Trim(!PAPRODCODE)
         bType = Format(!PALEVEL, "0")
         If bType = 6 Or bType = 7 Then
            sDebitAcct = "" & Trim(!PACGSEXPACCT)
            sCreditAcct = "" & Trim(!PAINVEXPACCT)
         Else
            sDebitAcct = "" & Trim(!PACGSMATACCT)
            sCreditAcct = "" & Trim(!PAINVMATACCT)
         End If
         .Cancel
      End With
   Else
      sCreditAcct = ""
      sDebitAcct = ""
      Exit Sub
   End If
   If sDebitAcct = "" Or sCreditAcct = "" Then
      'None in one or both there, try Product code
      sSql = "Qry_GetPCodeAccounts '" & sPcode & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            If bType = 6 Or bType = 7 Then
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(!PCCGSEXPACCT)
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(!PCINVEXPACCT)
            Else
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(!PCCGSMATACCT)
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(!PCMATEXPACCT)
            End If
            .Cancel
         End With
      End If
      If sDebitAcct = "" Or sCreditAcct = "" Then
         'Still none, we'll check the common
         If bType = 6 Or bType = 7 Then
            sSql = "SELECT COREF,COCGSEXPACCT" & Trim(str(bType)) & "," _
                   & "COINVEXPACCT" & Trim(str(bType)) & " " _
                   & "FROM ComnTable WHERE COREF=1"
         Else
            sSql = "SELECT COREF,COCGSMATACCT" & Trim(str(bType)) & "," _
                   & "COINVMATACCT" & Trim(str(bType)) & " " _
                   & "FROM ComnTable WHERE COREF=1"
         End If
         bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
         If bSqlRows Then
            With rdoAct
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(.Fields(0))
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(.Fields(1))
               .Cancel
            End With
         End If
      End If
   End If
   Set rdoAct = Nothing
   Exit Sub
   
DiaErr1:
   'Just bail for now. May not have anything set
   'CurrError.Number = Err
   'CurrError.Description = Err.Description
   'DoModuleErrors Me
   On Error GoTo 0
End Sub

Private Sub UpdateTotals()
   Dim i As Integer
   Dim nTotal As Single
   Dim nTax As Single
   Dim cFedTax As Currency
   
   On Error Resume Next
   For i = 0 To iTotalItems
      If vItems(i, INV_ShipThisItem) = vbChecked Then
         nTotal = nTotal + (vItems(i, INV_ShipQty) * vItems(i, INV_UnitPrice))
         If vItems(i, INV_FedTaxRate) <> 0 Then
            'cFedTax = cFedTax + (vItems(i, INV_ShipQty) * vItems(i, INV_FedTaxRate) * vItems(i, INV_UnitPrice) / 100)
            cFedTax = cFedTax + vItems(i, INV_FedTaxAmt) 'TaxPerLineItem = true
         End If
         If vItems(i, INV_SalesTaxRate) <> 0 Then
            'nTax = nTax + ((vItems(i, INV_ShipQty) * vItems(i, INV_UnitPrice)) * (vItems(i, INV_SalesTaxRate) / 100))
            nTax = nTax + vItems(i, INV_SalesTaxAmt) 'TaxPerLineItem = true
         ElseIf vItems(i, INV_TaxExempt) = 0 Then
            ' Item is taxable (TaxPerLineItem = false)
            nTax = nTax + ((vItems(i, INV_ShipQty) * vItems(i, INV_UnitPrice)) * (nTaxRate / 100))
         End If
      End If
   Next
   nTax = Round(nTax + 0.0001, 2)
   nTotal = nTotal + (nTax + lblFrt + cFedTax)
   lblTax = Format(CCur(nTax), CURRENCYMASK)
   lblFedTax = Format(CCur(cFedTax), CURRENCYMASK)
   lblTot = Format(CCur(nTotal), CURRENCYMASK)
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Function IsTaxCalculatedPerLineItem() As Boolean
   Dim rdo As ADODB.Recordset
   Dim result As Boolean
   
   result = False
   
   On Error GoTo done
   sSql = "SELECT TaxPerItem from ComnTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If bSqlRows Then
      result = rdo.Fields(0)
   End If
   
done:
   IsTaxCalculatedPerLineItem = result
End Function

Private Function GetPartLots(sPartWithLot As String) As Integer
   Dim RdoLots As ADODB.Recordset
   Dim iList As Integer
   Dim bFIFO As Byte
   Erase sLots
   On Error GoTo DiaErr1
   
   bFIFO = GetInventoryMethod()
 
   sSql = "SELECT LOTNUMBER,LOTPARTREF,LOTREMAININGQTY,LOTAVAILABLE " _
          & "FROM LohdTable WHERE (LOTPARTREF='" & sPartWithLot & "' AND " _
          & "LOTREMAININGQTY>0 AND LOTAVAILABLE=1) "
   If bFIFO = 1 Then
      sSql = sSql & "ORDER BY LOTNUMBER ASC"
   Else
      sSql = sSql & "ORDER BY LOTNUMBER DESC"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
   If bSqlRows Then
      With RdoLots
         Do Until .EOF
            iList = iList + 1
            sLots(iList, 0) = "" & Trim(!lotNumber)
            sLots(iList, 1) = Format$(!LOTREMAININGQTY, ES_QuantityDataFormat)
            .MoveNext
         Loop
         ClearResultSet RdoLots
      End With
      GetPartLots = iList
   Else
      GetPartLots = 0
   End If
   Set RdoLots = Nothing
   Exit Function
   
DiaErr1:
   GetPartLots = 0
   
End Function

'Private Function BuildPartToolTip(sPartExtDesc As String, sItemComment As String) As String
'    BuildPartToolTip = ""
'    If Len(Trim(sItemComment)) > 0 Then BuildPartToolTip = "Item Comments : " & Trim(sItemComment)
'    If Len(Trim(sPartExtDesc)) > 0 Then
'        If Len(BuildPartToolTip) > 0 Then BuildPartToolTip = BuildPartToolTip & vbCrLf
'        BuildPartToolTip = BuildPartToolTip & "Ext Part Description: " & Trim(sPartExtDesc)
'    End If
'
'End Function

