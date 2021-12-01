VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaARe06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Credit or Debit Memo Against Invoice"
   ClientHeight    =   7665
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   10335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7665
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCmt 
      Height          =   975
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   109
      Tag             =   "9"
      Top             =   6480
      Width           =   5055
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "diaARe06a.frx":0000
      DownPicture     =   "diaARe06a.frx":0972
      Height          =   350
      Left            =   6600
      Picture         =   "diaARe06a.frx":12E4
      Style           =   1  'Graphical
      TabIndex        =   108
      ToolTipText     =   "Standard Comments"
      Top             =   6480
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox txtFedTaxPercent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Tag             =   "1"
      Text            =   "0"
      ToolTipText     =   "Tax Percentage (Positive Values)"
      Top             =   3120
      Width           =   675
   End
   Begin VB.CommandButton cmdComputeFedTax 
      Caption         =   "% ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2580
      TabIndex        =   15
      ToolTipText     =   "Calculate % tax of prior invoice total revenue"
      Top             =   3120
      Width           =   405
   End
   Begin VB.TextBox txtFedTax 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3060
      TabIndex        =   16
      Tag             =   "1"
      Text            =   "0.00"
      ToolTipText     =   " Freight (Positive Values)"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ComboBox cboFedTaxAccount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      TabIndex        =   17
      Tag             =   "3"
      Text            =   "cboFedTaxAccount"
      Top             =   3120
      Width           =   1635
   End
   Begin VB.ComboBox cboAcct 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   8640
      TabIndex        =   33
      Tag             =   "3"
      Text            =   "cboAcct"
      Top             =   5580
      Width           =   1635
   End
   Begin VB.ComboBox cboAcct 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   8640
      TabIndex        =   30
      Tag             =   "3"
      Text            =   "cboAcct"
      Top             =   5220
      Width           =   1635
   End
   Begin VB.ComboBox cboAcct 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   8640
      TabIndex        =   27
      Tag             =   "3"
      Text            =   "cboAcct"
      Top             =   4860
      Width           =   1635
   End
   Begin VB.ComboBox cboAcct 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   8640
      TabIndex        =   24
      Tag             =   "3"
      Text            =   "cboAcct"
      Top             =   4500
      Width           =   1635
   End
   Begin VB.ComboBox cboAcct 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   8640
      TabIndex        =   21
      Tag             =   "3"
      Text            =   "cboAcct"
      Top             =   4140
      Width           =   1635
   End
   Begin VB.ComboBox cboARAccount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      TabIndex        =   18
      Tag             =   "3"
      Text            =   "cboARAccount"
      Top             =   3480
      Width           =   1635
   End
   Begin VB.ComboBox cboTaxAccount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      TabIndex        =   13
      Tag             =   "3"
      Text            =   "cboTaxAccount"
      Top             =   2760
      Width           =   1635
   End
   Begin VB.ComboBox cboFreightAccount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      TabIndex        =   9
      Tag             =   "3"
      Text            =   "cboFreightAccount"
      Top             =   2400
      Width           =   1635
   End
   Begin VB.ComboBox cboRevenueAccount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      TabIndex        =   7
      Tag             =   "3"
      Text            =   "cboRevenueAccount"
      Top             =   2040
      Width           =   1635
   End
   Begin VB.ComboBox cmbTyp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "diaARe06a.frx":1C56
      Left            =   6120
      List            =   "diaARe06a.frx":1C58
      TabIndex        =   5
      Tag             =   "8"
      ToolTipText     =   "Select CM (Credit Memo) Or DM (Debit Memo) From List"
      Top             =   1260
      Width           =   735
   End
   Begin VB.TextBox txtRevenue 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3060
      TabIndex        =   6
      Tag             =   "1"
      Text            =   "0.00"
      ToolTipText     =   " Freight (Positive Values)"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtTax 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3060
      TabIndex        =   12
      Tag             =   "1"
      Text            =   "0.00"
      ToolTipText     =   " Freight (Positive Values)"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdComputeTax 
      Caption         =   "% ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2580
      TabIndex        =   11
      ToolTipText     =   "Calculate % tax of prior invoice total revenue"
      Top             =   2760
      Width           =   405
   End
   Begin VB.TextBox txtAdjust 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   7440
      TabIndex        =   32
      Tag             =   "1"
      Top             =   5580
      Width           =   1095
   End
   Begin VB.TextBox txtAdjust 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   7440
      TabIndex        =   29
      Tag             =   "1"
      Top             =   5220
      Width           =   1095
   End
   Begin VB.TextBox txtAdjust 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   7440
      TabIndex        =   26
      Tag             =   "1"
      Top             =   4860
      Width           =   1095
   End
   Begin VB.TextBox txtAdjust 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   7440
      TabIndex        =   23
      Tag             =   "1"
      Top             =   4500
      Width           =   1095
   End
   Begin VB.TextBox txtAdjust 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   7440
      TabIndex        =   20
      Tag             =   "1"
      Top             =   4140
      Width           =   1095
   End
   Begin VB.ComboBox cmbCst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   300
      Width           =   1485
   End
   Begin VB.TextBox txtTaxPercent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Tag             =   "1"
      Text            =   "0"
      ToolTipText     =   "Tax Percentage (Positive Values)"
      Top             =   2760
      Width           =   675
   End
   Begin VB.TextBox txtFrt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3060
      TabIndex        =   8
      Tag             =   "1"
      Text            =   "0.00"
      ToolTipText     =   " Freight (Positive Values)"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ComboBox txtDte 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3060
      TabIndex        =   4
      Tag             =   "4"
      Top             =   1260
      Width           =   1095
   End
   Begin VB.TextBox txtRet 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   5280
      TabIndex        =   31
      Tag             =   "1"
      Top             =   5580
      Width           =   1095
   End
   Begin VB.TextBox txtRet 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   5280
      TabIndex        =   28
      Tag             =   "1"
      Top             =   5220
      Width           =   1095
   End
   Begin VB.TextBox txtRet 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   5280
      TabIndex        =   25
      Tag             =   "1"
      Top             =   4860
      Width           =   1095
   End
   Begin VB.TextBox txtRet 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   22
      Tag             =   "1"
      Top             =   4500
      Width           =   1095
   End
   Begin VB.TextBox txtRet 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   19
      Tag             =   "1"
      Top             =   4140
      Width           =   1095
   End
   Begin VB.CommandButton cmdPst 
      Caption         =   "&Post"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8820
      TabIndex        =   34
      Top             =   1320
      Width           =   875
   End
   Begin VB.CommandButton cmdRes 
      Caption         =   "&Reselect"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8820
      TabIndex        =   35
      Top             =   1680
      Width           =   875
   End
   Begin VB.ComboBox cmbInv 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "diaARe06a.frx":1C5A
      Left            =   4950
      List            =   "diaARe06a.frx":1C5C
      TabIndex        =   2
      Tag             =   "3"
      Top             =   300
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   120
      TabIndex        =   39
      Top             =   1140
      Width           =   9555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8820
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   180
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   37
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
      PictureUp       =   "diaARe06a.frx":1C5E
      PictureDn       =   "diaARe06a.frx":1DA4
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   38
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaARe06a.frx":1EEA
      PictureDn       =   "diaARe06a.frx":2030
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   5940
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7665
      FormDesignWidth =   10335
   End
   Begin Threed.SSCommand cmdDn 
      Height          =   375
      Left            =   8940
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Next Page (Page Down)"
      Top             =   5940
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaARe06a.frx":2182
   End
   Begin Threed.SSCommand cmdUp 
      Height          =   375
      Left            =   9300
      TabIndex        =   49
      TabStop         =   0   'False
      ToolTipText     =   "Last Page (Page Up)"
      Top             =   5940
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaARe06a.frx":2684
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   285
      Index           =   27
      Left            =   360
      TabIndex        =   110
      Top             =   6480
      Width           =   840
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fed Tax Adjustment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   26
      Left            =   120
      TabIndex        =   107
      Top             =   3120
      Width           =   1800
   End
   Begin VB.Label lblPriorFedTax 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   106
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Federal Tax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   25
      Left            =   7380
      TabIndex        =   105
      Top             =   3180
      Width           =   960
   End
   Begin VB.Label lblPO 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   104
      Top             =   660
      Width           =   1335
   End
   Begin VB.Label lblSO 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7620
      TabIndex        =   103
      Top             =   300
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   4920
      TabIndex        =   102
      Top             =   1380
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bulk Revenue Adjustment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   120
      TabIndex        =   101
      Top             =   2040
      Width           =   2700
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   4560
      TabIndex        =   100
      Top             =   1740
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revenue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   22
      Left            =   7380
      TabIndex        =   99
      Top             =   2100
      Width           =   1005
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   21
      Left            =   7380
      TabIndex        =   98
      Top             =   3540
      Width           =   765
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   20
      Left            =   7380
      TabIndex        =   97
      Top             =   2820
      Width           =   960
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   19
      Left            =   7380
      TabIndex        =   96
      Top             =   2460
      Width           =   960
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Prior Invoice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   6120
      TabIndex        =   95
      Top             =   1740
      Width           =   1035
   End
   Begin VB.Label lblPriorFreight 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   94
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblPriorRevenue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   93
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblPriorTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   92
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblPriorSalesTax 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   91
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblRevenue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3060
      TabIndex        =   90
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Line Item Revenue Adjustment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   120
      TabIndex        =   89
      Top             =   1680
      Width           =   2805
   End
   Begin VB.Label lblUnitPrice 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   6420
      TabIndex        =   88
      Top             =   5580
      Width           =   975
   End
   Begin VB.Label lblUnitPrice 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   6420
      TabIndex        =   87
      Top             =   5220
      Width           =   975
   End
   Begin VB.Label lblUnitPrice 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   6420
      TabIndex        =   86
      Top             =   4860
      Width           =   975
   End
   Begin VB.Label lblUnitPrice 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   6420
      TabIndex        =   85
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label lblUnitPrice 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   6420
      TabIndex        =   84
      Top             =   4140
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev Account                "
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
      Index           =   16
      Left            =   8640
      TabIndex        =   83
      Top             =   3900
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price               "
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
      Index           =   15
      Left            =   6420
      TabIndex        =   82
      Top             =   3900
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price Adjust    "
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
      Left            =   7440
      TabIndex        =   81
      ToolTipText     =   "Positive to reduce, negative to increase"
      Top             =   3900
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   80
      Top             =   360
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight Adjustment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   79
      Top             =   2400
      Width           =   1320
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax Adjustment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   78
      Top             =   2760
      Width           =   1800
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   120
      TabIndex        =   77
      Top             =   3480
      Width           =   1605
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3060
      TabIndex        =   76
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   75
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   74
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Returned   "
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
      Index           =   5
      Left            =   5280
      TabIndex        =   73
      Top             =   3900
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty                "
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
      Index           =   4
      Left            =   4260
      TabIndex        =   72
      Top             =   3900
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                          "
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
      Index           =   3
      Left            =   1440
      TabIndex        =   71
      Top             =   3900
      Width           =   2865
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order "
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
      Index           =   2
      Left            =   540
      TabIndex        =   70
      Top             =   3900
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item  "
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
      Left            =   120
      TabIndex        =   69
      Top             =   3900
      Width           =   345
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4260
      TabIndex        =   68
      Top             =   5580
      Width           =   975
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   4260
      TabIndex        =   67
      Top             =   5220
      Width           =   975
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   4260
      TabIndex        =   66
      Top             =   4860
      Width           =   975
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4260
      TabIndex        =   65
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4260
      TabIndex        =   64
      Top             =   4140
      Width           =   975
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   1440
      TabIndex        =   63
      Top             =   5580
      Width           =   2775
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   62
      Top             =   5220
      Width           =   2775
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   61
      Top             =   4860
      Width           =   2775
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   60
      Top             =   4500
      Width           =   2775
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   59
      Top             =   4140
      Width           =   2775
   End
   Begin VB.Label lblSon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   540
      TabIndex        =   58
      Top             =   5580
      Width           =   855
   End
   Begin VB.Label lblSon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   540
      TabIndex        =   57
      Top             =   5220
      Width           =   855
   End
   Begin VB.Label lblSon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   540
      TabIndex        =   56
      Top             =   4860
      Width           =   855
   End
   Begin VB.Label lblSon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   540
      TabIndex        =   55
      Top             =   4500
      Width           =   855
   End
   Begin VB.Label lblSon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   540
      TabIndex        =   54
      Top             =   4140
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   8220
      TabIndex        =   53
      Top             =   6060
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   7020
      TabIndex        =   52
      Top             =   6060
      Width           =   615
   End
   Begin VB.Label lblLst 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8460
      TabIndex        =   51
      Top             =   6060
      Width           =   375
   End
   Begin VB.Label lblPge 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7740
      TabIndex        =   50
      Top             =   6060
      Width           =   375
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   3480
      Picture         =   "diaARe06a.frx":2B86
      Top             =   6060
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   2760
      Picture         =   "diaARe06a.frx":3078
      Top             =   6060
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   3000
      Picture         =   "diaARe06a.frx":356A
      Top             =   6060
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   3240
      Picture         =   "diaARe06a.frx":3A5C
      Top             =   6060
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblItm 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   47
      Top             =   5580
      Width           =   375
   End
   Begin VB.Label lblItm 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   46
      Top             =   5220
      Width           =   375
   End
   Begin VB.Label lblItm 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   45
      Top             =   4860
      Width           =   375
   End
   Begin VB.Label lblItm 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   44
      Top             =   4500
      Width           =   375
   End
   Begin VB.Label lblItm 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   36
      Top             =   4140
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice To Credit or Debit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   2760
      TabIndex        =   43
      Top             =   360
      Width           =   1845
   End
   Begin VB.Label lblPre 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4740
      TabIndex        =   42
      Top             =   300
      Width           =   255
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   41
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   40
      ToolTipText     =   "Invoice Source"
      Top             =   300
      Width           =   1335
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaARe06a"
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

'*********************************************************************************
' diaARe06a - Credit or Debit Memo Against Prior Invoice
'
' Notes:
'
' Created: (nth) 04/04/04
' Revisions:
'
'*********************************************************************************


Dim bOnLoad As Byte
Dim bCancel As Byte
Dim sMsg As String
Dim sJournalID As String
Dim iIndex As Integer
Dim iCurrPage As Integer
Dim iTotalItems As Integer

'- vItems Documentation
'0 = Item Number
'1 = Rev
'2 = SO
'3 = Part
'4 = Qty
'5 = Unit Price
'6 = Qty Return
'7 = Lot Tracked
'8 = Unit price adjustment
'9 = account

Private Const LINES_PER_PAGE = 5
Private Const MAX_PAGES = 40
Private Const IT_NUMBER = 0
Private Const IT_REV = 1
Private Const IT_SO = 2
Private Const IT_PART = 3
Private Const IT_QTY = 4
Private Const IT_UNITPRICE = 5
Private Const IT_RETURNQTY = 6
Private Const IT_LOTTRACKED = 7
Private Const IT_ADJUST = 8
Private Const IT_REVACCT = 9
Private Const IT_PSNUMBER = 10
Private Const IT_PSITEM = 11
Dim vItems(LINES_PER_PAGE * MAX_PAGES, 11) As Variant

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillInvoicesForCustomer(NickName As String, frm As Form, cbo As ComboBox)
   MouseCursor 13
   Dim CstRes As ADODB.Recordset
   On Error GoTo modErr1
   cbo.Clear
   ClearItems
   lblPre = ""
   lblTyp = ""
   lblNme = ""
   lblCst = ""
   sSql = "select INVNO from CihdTable where INVCUST = '" & NickName & "'" & vbCrLf _
          & "and INVTYPE <> 'CM' and INVTYPE <> 'DM'" & vbCrLf _
          & "order by INVNO desc"
   bSqlRows = clsADOCon.GetDataSet(sSql, CstRes, ES_FORWARD)
   If bSqlRows Then
      With CstRes
         Do Until .EOF
            AddComboStr cbo.hWnd, "" & Trim(.Fields(0))
            .MoveNext
         Loop
         .Cancel
      End With
      cbo.ListIndex = 0
   End If
   Set CstRes = Nothing
   MouseCursor 0
   Exit Sub
   
modErr1:
   MouseCursor 0
   sProcName = "FillInvoicesForCustomer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Set CstRes = Nothing
   DoModuleErrors frm
   
End Sub

Private Sub cmbCst_Click()
   FillInvoicesForCustomer cmbCst.Text, Me, cmbInv
End Sub

Private Sub cmbCst_LostFocus()
   'cmbCst_Click
End Sub

Private Sub cmbInv_Click()
   GetInvoice
End Sub

Private Sub cmbInv_GotFocus()
   SelectFormat Me
End Sub

Private Sub cmbInv_LostFocus()
   If Not bCancel Then
      GetInvoice
   End If
End Sub

Private Sub cmbTyp_Click()
   If cmbTyp.Text = "DM" Then
      ClearItems
   Else
      GetItems
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdComments_Click()
'bbs changes from the Comments form to the SysComments form on 6/28/2010 for Ticket #31511
   'Add one of these to the form and go
   If cmdComments Then
      'The Default is txtCmt and need not be included
      'Use Select Case cmdCopy to add your own
      SysComments.lblControl = "txtCmt"
      'See List For Index
      SysComments.lblListIndex = 4
      SysComments.Show
      cmdComments = False
   End If
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 2048)
End Sub

Private Sub cmdComputeTax_Click()
   Dim cTaxRate As Currency
   On Error Resume Next
   cTaxRate = CCur("0" & Trim(txtTaxPercent))
   If Err Then
      MsgBox "Tax rate is not valid"
      txtTaxPercent.SetFocus
      Exit Sub
   End If
   
   If cTaxRate > 20 Then
      MsgBox "Tax rate must be less than or equal to 20%"
      txtTaxPercent.SetFocus
      Exit Sub
   End If
   
   Dim tax As Currency
   tax = CCur(lblPriorRevenue) * cTaxRate / 100
   If tax > CCur(lblPriorSalesTax) And cmbTyp = "CM" Then
      MsgBox "Can't credit more than original tax amount: " & lblPriorSalesTax
      Exit Sub
   End If
   
   txtTax = Format(CCur(lblPriorRevenue) * cTaxRate / 100, "0.00")
   UpdateTotals
End Sub

Private Sub cmdComputeFedTax_Click()
   Dim cFedTaxRate As Currency
   On Error Resume Next
   cFedTaxRate = CCur("0" & Trim(txtFedTaxPercent))
   If Err Then
      MsgBox "Federal Tax rate is not valid"
      txtFedTaxPercent.SetFocus
      Exit Sub
   End If
   
   If cFedTaxRate > 20 Then
      MsgBox "Federal Tax rate must be less than or equal to 20%"
      txtFedTaxPercent.SetFocus
      Exit Sub
   End If
   
   Dim FedTax As Currency
   FedTax = CCur(lblPriorRevenue) * cFedTaxRate / 100
   If FedTax > CCur(lblPriorFedTax) And cmbTyp = "CM" Then
      MsgBox "Can't credit more than original FedTax amount: " & lblPriorFedTax
      Exit Sub
   End If
   
   txtFedTax = Format(CCur(lblPriorRevenue) * cFedTaxRate / 100, "0.00")
   UpdateTotals
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

Private Sub cmdPst_Click()
   
   If IsRevenueTotalOK Then
      CreateInvoice
      Unload Me
   End If
   
End Sub

Private Sub cmdRes_Click()
   ManageBoxs False
   cmbInv.SetFocus
End Sub

'Private Sub cmdSel_Click()
'    GetItems
'End Sub
'

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

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CurrentJournal "SJ", ES_SYSDATE, sJournalID
      'FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   
   Dim RevAcct As String
   
   FormLoad Me
   
   cmbTyp.AddItem "CM"
   cmbTyp.AddItem "DM"
   cmbTyp = "CM"
   
   'fill combo boxes with account numbers
   Dim gl As New GLTransaction
   gl.FillComboWithAccounts cboRevenueAccount, cboFreightAccount, cboTaxAccount, cboFedTaxAccount, _
      cboARAccount, cboAcct(1), cboAcct(2), cboAcct(3), cboAcct(4), cboAcct(5)
   
   
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
   FillCustomers Me
   If cUR.CurrentCustomer <> "" Then
      cmbCst = cUR.CurrentCustomer
      cmbCst_Click
   Else
      If cmbCst.ListCount > 0 Then cmbCst = cmbCst.List(0)
   End If
   txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   ManageBoxs False
   
   ' Get company management accout
   RevAcct = GetRevAccountFromCompany
   
   If (RevAcct <> "") Then
      cboRevenueAccount = RevAcct
   End If
   
   
   GetOptions
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   cUR.CurrentCustomer = cmbCst
   SaveCurrentSelections
   FormUnload
   Set diaARe06a = Nothing
End Sub

Private Sub Image1_Click()
   UnderCons.Show
   
End Sub

Private Sub Label1_Click()
   UnderCons.Show
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

'Public Sub FillCombo()
'    Dim RdoInv As ADODB.RecordSet
'    On Error GoTo DiaErr1
'    sSql = "SELECT INVNO FROM CihdTable WHERE INVCANCELED = 0 " _
'        & "AND INVTYPE IN ('SO','PS') ORDER BY INVNO DESC"
'    bSqlRows = clsAdoCon.GetDataSet(sSql,RdoInv)
'    If bSqlRows Then
'        With RdoInv
'            While Not .EOF
'                AddComboStr cmbInv.hWnd, Format(.Fields(0), "000000")
'                .MoveNext
'            Wend
'            .Cancel
'        End With
'        cmbInv.ListIndex = 0
'    End If
'    Set RdoInv = Nothing
'    Exit Sub
'DiaErr1:
'    sProcName = "fillcomb"
'    CurrError.Number = Err
'    CurrError.Description = Err.Description
'    DoModuleErrors Me
'End Sub
'

Private Sub SaveOptions()
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name _
                & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub GetInvoice()
   'returns false if invoice not found
   Dim RdoInv As ADODB.Recordset
   Dim sType As String
   On Error GoTo DiaErr1
   ClearItems
   
   sSql = "SELECT INVPRE,INVTYPE,CUNAME,CUNICKNAME,INVTOTAL,INVFREIGHT,INVTAX,INVFEDTAX," & vbCrLf _
          & "INVFRTACCT,INVTAXACCT,INVFEDTAXACCT,INVARACCT,INVSO,INVPACKSLIP,ISNULL(INVREASONS, '') INVREASONS " & vbCrLf _
          & "FROM CihdTable INNER JOIN CustTable ON INVCUST = CUREF " _
          & "AND INVCUST = '" & Compress(cmbCst) & "' " _
          & "WHERE (INVNO = " & Val(cmbInv) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
   If bSqlRows Then
      With RdoInv
         lblPre = "" & Trim(!INVPRE)
         lblTyp = GetInvoiceType(!INVTYPE)
         lblNme = "" & Trim(!CUNAME)
         lblCst = "" & Trim(!CUNICKNAME)
         lblPriorRevenue = Format(!INVTOTAL - !INVFREIGHT - !INVTAX, "0.00")
         lblPriorTotal = Format(!INVTOTAL, "0.00")
         lblPriorSalesTax = Format(!INVTAX, "0.00")
         lblPriorFedTax = Format(!INVFEDTAX, "0.00")
         lblPriorFreight = Format(!INVFREIGHT, "0.00")
         txtCmt = Trim(!INVREASONS)
         
         On Error Resume Next
         cboFreightAccount = Trim(!INVFRTACCT)
         cboTaxAccount = Trim(!INVTAXACCT)
         cboFedTaxAccount = Trim(!INVFEDTAXACCT)
         cboARAccount = Trim(!INVARACCT)
         On Error GoTo DiaErr1
         
         sType = Trim(!INVTYPE)
         lblSO = ""
         lblPO = ""
         If sType = "SO" Then
            lblSO = !INVSO
         ElseIf sType = "PS" Then
            lblSO = !INVPACKSLIP
         End If
         .Cancel
      End With
   ElseIf Len(Trim(cmbInv)) > 0 Then
      Set RdoInv = Nothing
      MsgBox "No such invoice number for this customer", vbInformation
      Exit Sub
      
   Else
      Set RdoInv = Nothing
      Exit Sub
      Exit Sub
      
   End If
   Set RdoInv = Nothing
   
   'if SO, get PO number
   If sType = "SO" Then
      sSql = "SELECT SOPO FROM SOHDTABLE " _
             & "WHERE SONUMBER = " & lblSO
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
      If bSqlRows Then
         With RdoInv
            If Trim(!SOPO) <> "" Then
               lblPO = "PO " & Trim(!SOPO)
            End If
            .Cancel
         End With
      End If
      Set RdoInv = Nothing
      
   ElseIf sType = "PS" Then
      sSql = "SELECT SOPO FROM SohdTable " _
             & "JOIN PshdTable on PSPRIMARYSO = SONUMBER " _
             & "WHERE PSNUMBER = '" & lblSO & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
      If bSqlRows Then
         With RdoInv
            If Trim(!SOPO) <> "" Then
               lblPO = "PO " & Trim(!SOPO)
            End If
            .Cancel
         End With
      End If
      Set RdoInv = Nothing
   End If
   
   'if accounts not in invoice, get from common
   'On Error Resume Next        'in case account #'s do not match
   Dim gl As New GLTransaction
   If cboARAccount = "" Then
      'cboARAccount = gl.SJARAccount
      SetCompressedComboBox cboARAccount, gl.SJARAccount
   End If
   If cboTaxAccount = "" Then
      'cboTaxAccount = gl.SJTaxAccount
      SetCompressedComboBox cboTaxAccount, gl.SJTaxAccount
   End If
   If cboFedTaxAccount = "" Then
      SetCompressedComboBox cboFedTaxAccount, gl.SJFedTaxAccount
   End If
   If cboFreightAccount = "" Then
      'cboFreightAccount = gl.SJFreightAccount
      SetCompressedComboBox cboFreightAccount, gl.SJFreightAccount
   End If
   On Error GoTo DiaErr1
   
   GetItems
   
   Exit Sub
DiaErr1:
   sProcName = "getinvoi"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub ClearItems()
   Dim i As Integer
   For i = 1 To LINES_PER_PAGE
      lblItm(i) = ""
      lblSon(i) = ""
      lblPrt(i) = ""
      lblQty(i) = ""
      txtRet(i) = ""
      lblUnitPrice(i) = ""
      txtAdjust(i) = ""
      txtRet(i).enabled = False
      txtAdjust(i).enabled = False
      cboAcct(i).enabled = False
   Next
End Sub

Public Sub GetItems()
   Dim b As Byte
   Dim RdoItm As ADODB.Recordset
   Dim RdoQty As ADODB.Recordset
   'Dim RdoRet As ADODB.Recordset
   Erase vItems
   ClearItems
   iTotalItems = 0
   On Error GoTo DiaErr1
   sSql = "SELECT ITNUMBER, ITREV, ITSO, PARTNUM, ITACTUAL, ITDOLLARS," & vbCrLf _
      & "SOTYPE, PALOTTRACK, ITREVACCT, ITCOST, ITPSNUMBER, ITPSITEM " & vbCrLf _
      & "FROM SoitTable" & vbCrLf _
      & "JOIN PartTable ON ITPART = PARTREF" & vbCrLf _
      & "JOIN SohdTable ON ITSO = SONUMBER" & vbCrLf _
      & "WHERE ITINVOICE = " & Val(cmbInv)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm)
   If bSqlRows Then
      With RdoItm
         While Not .EOF
            b = b + 1
            vItems(b, IT_NUMBER) = !ITNUMBER '.Fields(0)
            vItems(b, IT_REV) = !itrev '.Fields(1)
            vItems(b, IT_SO) = !SOTYPE & Format(!ITSO, SO_NUM_FORMAT) '.Fields(6) & Format(.Fields(2), "00000")
Debug.Print b & ": " & vItems(b, IT_SO)
            vItems(b, IT_PART) = Trim(!PARTNUM) '.Fields(3)
            vItems(b, IT_LOTTRACKED) = !PALOTTRACK '.Fields(7)
            vItems(b, IT_ADJUST) = 0
            vItems(b, IT_REVACCT) = Trim(!ITREVACCT)
            vItems(b, IT_UNITPRICE) = !ITDOLLARS
            vItems(b, IT_PSNUMBER) = !ITPSNUMBER
            vItems(b, IT_PSITEM) = !ITPSITEM
            
            'set default revenue account for bulk revenue adjustment to first line item acct
            If Trim(cboRevenueAccount.Text) = "" Then
               'cboRevenueAccount.Text = Trim(!ITREVACCT)
               SetCompressedComboBox cboRevenueAccount, Trim(!ITREVACCT)
            End If
            
'            'calculate quantity remaining
'            sSql = "SELECT ISNULL(SUM(INAQTY),0) FROM InvaTable " & vbCrLf _
'                   & "WHERE INSONUMBER = " & !ITSO & vbCrLf _
'                   & "AND INSOITEM = " & !ITNUMBER & vbCrLf _
'                   & "AND INSOREV = '" & Trim(!itrev) & "'"
'            Debug.Print sSql
'            bSqlRows = clsAdoCon.GetDataSet(sSql,RdoQty)
'            If bSqlRows Then
'               vItems(b, IT_QTY) = RdoQty.Fields(0) * -1
'            End If
            
            
            'calculate quantity remaining
            Dim outstandingQty As Currency
            outstandingQty = 0
            sSql = "SELECT ISNULL(SUM(INAQTY),0) FROM InvaTable " & vbCrLf _
                   & "WHERE INSONUMBER = " & !ITSO & vbCrLf _
                   & "AND INSOITEM = " & !ITNUMBER & vbCrLf _
                   & "AND INSOREV = '" & Trim(!itrev) & "'"
            bSqlRows = clsADOCon.GetDataSet(sSql, RdoQty)
            If bSqlRows Then
               outstandingQty = RdoQty.Fields(0) * -1
            End If
            
'            ' 1/27/08
'            'due to a bug, InvaTable records didn't always get created.
'            'so if qty = 0, try looking up quantity in LoitTable
'            'this should fix about 90% of the problem cases
'            If outstandingQty = 0 Then
'               sSql = "SELECT - ISNULL( SUM( LOIQUANTITY ), 0 ) FROM LoitTable" & vbCrLf _
'                  & "WHERE LOIPSNUMBER = '" & !ITPSNUMBER & "' AND LOIPSITEM = '" & !ITPSITEM & "'"
'               If GetDataSet(RdoQty) Then
'                  outstandingQty = RdoQty.Fields(0)
'               End If
'            End If
            
            vItems(b, IT_QTY) = outstandingQty
            
            RdoQty.Cancel
            Set RdoQty = Nothing
            .MoveNext
         Wend
         iTotalItems = b
         iCurrPage = 1
         lblLst = CInt((iTotalItems / 5) + 0.5)
         lblPge = 1
         ManageBoxs True
         GetNextGroup
         If Val(lblLst) > 1 Then
            cmdDn.enabled = True
            cmdDn.Picture = Endn
         End If
         UpdateTotals
         'txtRet(1).SetFocus
      End With
   End If
   Set RdoItm = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub GetNextGroup()
   Dim i As Integer
   Dim sType As String
   
   ClearItems
   iIndex = (Val(lblPge) - 1) * LINES_PER_PAGE
   On Error GoTo DiaErr1
   For i = 1 To LINES_PER_PAGE
      If i + iIndex > iTotalItems Then
         lblItm(i) = ""
         lblSon(i) = ""
         lblPrt(i) = ""
         lblQty(i) = ""
         txtRet(i) = ""
         txtRet(i).enabled = False
      Else
         '0 = Item Number
         '1 = Rev
         '2 = SO
         '3 = Part
         '4 = Qty
         '5 = Unit Price
         '6 = Qty Return
         lblItm(i) = vItems(i + iIndex, IT_NUMBER) & vItems(i, IT_REV)
         lblSon(i) = vItems(i + iIndex, IT_SO)
         lblPrt(i) = Trim(vItems(i + iIndex, IT_PART))
         lblQty(i) = Format(CSng(vItems(i + iIndex, IT_QTY)), "0.000")
         txtRet(i) = Format(CSng(vItems(i + iIndex, IT_RETURNQTY)), "0.000")
         txtRet(i).enabled = True
         lblUnitPrice(i) = Format(CSng(vItems(i + iIndex, IT_UNITPRICE)), "0.0000")
         txtAdjust(i) = Format(CSng(vItems(i + iIndex, IT_ADJUST)), "0.0000")
         txtAdjust(i).enabled = True
         'cboAcct(i) = vItems(i + iIndex, IT_REVACCT)
         SetCompressedComboBox cboAcct(i), CStr(vItems(i + iIndex, IT_REVACCT))
         cboAcct(i).enabled = True
         
      End If
   Next
   Exit Sub
DiaErr1:
   sProcName = "getnextgr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub ManageBoxs(b As Byte)
   Exit Sub
   Dim i As Byte
   If Not b Then
      cmdRes.enabled = False
      'cmdSel.Enabled = True
      cmbInv.enabled = True
      txtDte.enabled = True 'False
      lblTot.enabled = False
      lblTot = ""
      For i = 1 To 5
         txtRet(i) = ""
         lblItm(i) = ""
         lblSon(i) = ""
         lblPrt(i) = ""
         lblQty(i) = ""
         txtRet(i).enabled = True
      Next
      cmdPst.enabled = False
   Else
      cmdRes.enabled = True
      'cmdSel.Enabled = False
      cmbInv.enabled = False
      txtDte.enabled = True
      lblTot.enabled = True
      cmdPst.enabled = True
      For i = 1 To 5
         lblItm(i) = ""
         lblSon(i) = ""
         lblPrt(i) = ""
         lblQty(i) = ""
         txtRet(i).enabled = True
      Next
   End If
End Sub

Private Sub cboAcct_LostFocus(Index As Integer)
   
   'don't try to validate if leaving form
   If Me.ActiveControl.Name = "cmdCan" Then Exit Sub
   
   'validate account
   Dim S As String
   On Error Resume Next
   S = Trim(cboAcct(Index).Text)
   Dim gl As New GLTransaction
   If gl.IsValidRevenueAccount(S) Then
      vItems(iIndex + Index, IT_REVACCT) = S
   Else
      cboAcct(Index).SetFocus
   End If
End Sub

Private Sub txtAdjust_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtAdjust_LostFocus(Index As Integer)
   Dim S As Single
   On Error Resume Next
   S = CSng(txtAdjust(Index).Text)
   If Err Then
      MsgBox "invalid adjustment amount"
      txtAdjust(Index).SetFocus
      Exit Sub
   End If
   
   If S > vItems(iIndex + Index, IT_UNITPRICE) Then
      sMsg = "You cannot reduce price by more than original price.  " _
             & "If you are unsure about the amount, you can exit this field by entering a zero for the price reduction"
      MsgBox sMsg, vbInformation, Caption
      txtAdjust(Index).SetFocus
   Else
      vItems(iIndex + Index, IT_ADJUST) = S
      txtAdjust(Index) = Format(S, "0.0000")
   End If
   UpdateTotals
End Sub

Private Sub cboARAccount_LostFocus()
   'don't try to validate if leaving form
   If Me.ActiveControl.Name = "cmdCan" Then Exit Sub
   
   CheckAccountInComboBox cboARAccount
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDte_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
End Sub

Private Sub cboFreightAccount_LostFocus()
   'don't try to validate if leaving form
   If Me.ActiveControl.Name = "cmdCan" Then Exit Sub
   
   CheckAccountInComboBox cboFreightAccount
End Sub

Private Sub txtFrt_LostFocus()
   'make sure number is valid
   If Not CurrencyTextLostFocus(txtFrt) Then
      Exit Sub
   End If
   
   Dim cFrt As Currency
   On Error Resume Next
   cFrt = CCur(txtFrt)
   
   If cFrt > CCur(lblPriorFreight) And cmbTyp = "CM" Then
      MsgBox "Can't credit more than original freight amount: " & lblPriorFreight
      txtFrt.SetFocus
      Exit Sub
   End If
   
   UpdateTotals
End Sub

Private Sub txtRet_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtRet_LostFocus(Index As Integer)
   If Val(txtRet(Index)) > vItems(iIndex + Index, IT_QTY) Then
      sMsg = "Cannot Return More Than Original Quantity."
      MsgBox sMsg, vbInformation, Caption
      txtRet(Index).SetFocus
   Else
      vItems(iIndex + Index, IT_RETURNQTY) = Val(txtRet(Index))
      txtRet(Index) = Format(txtRet(Index), "0.000")
   End If
   UpdateTotals
End Sub

Public Function GetNextInvoice() As Long
'   Dim RdoInv As ADODB.RecordSet
'   On Error GoTo DiaErr1
'   sSql = "SELECT MAX(INVNO) FROM CihdTable"
'   bSqlRows = clsAdoCon.GetDataSet(sSql,RdoInv)
'   If bSqlRows Then
'      If IsNull(RdoInv.Fields(0)) Then
'         GetNextInvoice = 1
'      Else
'         GetNextInvoice = (RdoInv.Fields(0) + 1)
'      End If
'   Else
'      GetNextInvoice = 1
'   End If
'   Set RdoInv = Nothing
'   Exit Function
'
'DiaErr1:
'   sProcName = "getnextinvoice"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
   
   Dim inv As New ClassARInvoice
   GetNextInvoice = inv.GetNextInvoiceNumber

End Function

Public Sub CreateInvoice()
   Dim RdoInv As ADODB.Recordset
   Dim rdoAct As ADODB.Recordset
   Dim b As Byte
   Dim K As Byte
   Dim lInvoice As Long
   Dim lSo As Long
   Dim nItem As Integer
   Dim sRev As String
   Dim cTotal As Currency
   Dim bInType As Byte
   Dim sNow As String
   Dim lLastActivity As Long
   Dim sCust As String
   
   Dim lTrans As Long
   Dim iRef As Integer
   Dim sPart As String
   Dim sDebitAccount As String
   Dim sCreditAccount As String
   Dim sComments As String
   Dim sPre As String: sPre = "I"
   
   On Error GoTo DiaErr1
   
   ' Check for an open journal
   sJournalID = GetOpenJournal("SJ", txtDte)
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   If b = 0 Then
      MsgBox "There Is No Open Sales Journal For The Period.", _
         vbInformation, Caption
      Exit Sub
   End If
   lTrans = GetNextTransaction(sJournalID)
   MouseCursor 13
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   lInvoice = GetNextInvoice()
   lLastActivity = GetLastActivity()
   sNow = Format(GetServerDateTime(), "mm/dd/yy")
   sCust = Compress(lblCst)
   lTrans = GetNextTransaction(sJournalID)
   
   'create a general ledger transaction object for creation of debits and credits
   Dim gl As New GLTransaction
   gl.JournalID = sJournalID
   gl.InvoiceDate = CDate(txtDte.Text)
   gl.InvoiceNumber = lInvoice
   
   Dim cReturnQty As Currency
   Dim cQty As Currency
   Dim cAdjust As Currency
   Dim revenue As Currency
   Dim sUnits As String
   
   For b = 1 To iTotalItems
      lSo = CLng(Mid(vItems(b, IT_SO), 2, Len(vItems(b, IT_SO)) - 1))
      cReturnQty = Val(vItems(b, IT_RETURNQTY))
      cQty = vItems(b, IT_QTY)
      cAdjust = vItems(b, IT_ADJUST)
      
      'Don't allow if AR account number is empty
      If (cboARAccount = "") Then
        sMsg = "Please select Revenue Account"
        MsgBox sMsg, vbInformation, Caption
        Exit Sub
      End If
      
      'only do quantity stuff for credit memo
      If cQty > 0 And (cReturnQty > 0 Or cAdjust <> 0) And cmbTyp = "CM" Then
         sPart = Trim(Compress(vItems(b, IT_PART)))
         nItem = CInt(vItems(b, IT_NUMBER))
         sRev = Trim(CStr(vItems(b, IT_REV)))
         revenue = GetExtension(cQty, cReturnQty, CCur(vItems(b, IT_UNITPRICE)), cAdjust)
         
         'get units - InvaTable.INUNITS is not reliable
         sSql = "SELECT PAUNITS FROM PartTable WHERE PARTREF = '" & sPart & "'"
         bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct)
         sUnits = Trim(rdoAct!PAUNITS)
         rdoAct.Close
         Set rdoAct = Nothing
         
         Dim sLot As String
         sLot = ""
         If cReturnQty > 0 Then
            ' Adjust inventory
            lLastActivity = lLastActivity + 1
            sSql = "SELECT INAMT,INAQTY, ISNULL(INTOTMATL, 0)INTOTMATL," & vbCrLf _
                  & "ISNULL(INTOTLABOR, 0) INTOTLABOR ,ISNULL(INTOTEXP,0) INTOTEXP," & vbCrLf _
                  & "ISNULL(INTOTOH, 0) INTOTOH,ISNULL(INTOTHRS,0) INTOTHRS," & vbCrLf _
                   & "INTYPE,INCREDITACCT,INDEBITACCT," & vbCrLf _
                   & "INPSNUMBER,INPSITEM,INUNITS" & vbCrLf _
                   & "FROM InvaTable " & vbCrLf _
                   & "WHERE INSONUMBER = " & lSo _
                   & " AND INSOITEM = " & nItem _
                   & " AND INSOREV = '" & sRev & "'"
            bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_STATIC)
            With rdoAct
               If "" & Trim(!INPSNUMBER) <> "" Then
                  bInType = 4 ' returned so item
               Else
                  bInType = 33 ' returned ps item
               End If
            End With
            'Set rdoAct = Nothing
            
            
            Dim cInAmt As Currency
            Dim cUnitOvHd As Currency
            Dim cUnitLabor As Currency
            Dim cUnitMatl As Currency
            Dim cUnitExp As Currency
            Dim cUnitHrs As Currency
            Dim cInvQty As Currency
            
            Dim cUnitCost As Currency
            Dim cOvHd As Currency
            Dim cLabor As Currency
            Dim cMatl As Currency
            Dim cExp As Currency
            Dim cHours As Currency
            
            With rdoAct
               cInAmt = !INAMT
               cInvQty = Abs(!INAQTY)
               If cInvQty <> 0 Then
                  cUnitMatl = !INTOTMATL / cInvQty
                  cUnitLabor = !INTOTLABOR / cInvQty
                  cUnitExp = !INTOTEXP / cInvQty
                  cUnitOvHd = !INTOTOH / cInvQty
                  cUnitHrs = !INTOTHRS / cInvQty
               Else
                  cUnitMatl = !INTOTMATL
                  cUnitLabor = !INTOTLABOR
                  cUnitExp = !INTOTEXP
                  cUnitOvHd = !INTOTOH
                  cUnitHrs = !INTOTHRS
               End If
            End With
                                

                        ' Adjust lots (if lot tracked)
            If vItems(b, IT_LOTTRACKED) = 1 Then
               sMsg = "Lot Tracking Is Required For " _
                      & Trim(vItems(b, IT_PART)) & "." & vbCrLf _
                      & "You Must Select The Lot(s) Used For Return."
               MsgBox sMsg, vbInformation, Caption
               'Lots(0).LotSysId = "" ' clear lot array
               LotSelectReturn.lblPart = vItems(b, IT_PART)
               LotSelectReturn.lblRequired = cReturnQty
               LotSelectReturn.chkReturn = 1 'show lots with zero quantity
               LotSelectReturn.SqlFilter = "LOTNUMBER IN (SELECT LOINUMBER FROM LoitTable" & vbCrLf _
                  & "WHERE LOIPSNUMBER = '" & vItems(b, IT_PSNUMBER) & "'" & vbCrLf _
                  & "AND LOIPSITEM = " & vItems(b, IT_PSITEM) & ")"
               LotSelectReturn.Show 1
               
               ' No lots select then cancel invoice.
               'If Trim(Lots(0).LotSysId) <> "" Then     'creates subscript error if no lots
               If Es_TotalLots = 0 Then
                  sMsg = "Cannot Create Invoice." & vbCrLf _
                         & "If Lot Tracking Is Turned On Then Check Lots." & vbCrLf _
                         & "Make Certain That A Lot Was Selected."
                  MsgBox sMsg, vbExclamation, Caption
                  clsADOCon.RollbackTrans
                  Exit Sub
               End If
               
               For K = 1 To UBound(lots)
                  lLastActivity = lLastActivity + 1
                  ' We got lots update LoitTable
                  
                  sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD,LOITYPE," _
                         & "LOIPARTREF,LOIADATE,LOIPDATE,LOIQUANTITY,LOICUST," _
                         & "LOICUSTINVNO,LOIACTIVITY,LOICOMMENT,LOIUNITS) VALUES" _
                         & "('" & lots(K).LotSysId & "'," _
                         & GetNextLotRecord(lots(K).LotSysId) & "," _
                         & bInType & ",'" & lots(K).LotPartRef & "','" _
                         & sNow & "','" & txtDte & "'," _
                         & (lots(K).LotSelQty) & ",'" _
                         & sCust & "'," _
                         & lInvoice & "," _
                         & lLastActivity & "," _
                         & "'Returned Item','" _
                         & sUnits & "')"
                  clsADOCon.ExecuteSQL sSql
                  
                  If sLot = "" Then
                     sLot = "Lot: "
                  Else
                     sLot = sLot & ", "
                  End If
                  sLot = sLot & lots(K).LotUserId ' don't use LotSysId
                  
                  sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY+" _
                         & lots(K).LotSelQty & " WHERE LOTNUMBER = '" & lots(K).LotSysId & "'"
                  clsADOCon.ExecuteSQL sSql
            
                  GetInventoryCost lots(K).LotSysId, cUnitCost, cUnitMatl, cUnitExp, cUnitLabor, cUnitOvHd, cUnitHrs
                  
                  cMatl = cUnitMatl * lots(K).LotSelQty
                  cLabor = cUnitLabor * lots(K).LotSelQty
                  cExp = cUnitExp * lots(K).LotSelQty
                  cOvHd = cUnitOvHd * lots(K).LotSelQty
                  cHours = cUnitHrs * lots(K).LotSelQty
                        
                  With rdoAct
                     '                            sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                     '                                & "INADATE,INAQTY,INAMT,INSONUMBER,INSOITEM,INSOREV," _
                     '                                & "INCREDITACCT,INDEBITACCT,INNUMBER,INLOTNUMBER) " _
                     '                                & "VALUES(" & bInType & ",'" & Trim(vItems(b, IT_PART)) & "','RETURN','','" _
                     '                                & sNow & "'," & (Lots(k).LotSelQty) & "," & !INAMT & "," _
                     '                                & lSo & "," & nItem & ",'" & sRev & "','" _
                     '                                & Trim(!INDEBITACCT) & "','" & Trim(!INCREDITACCT) & "'," _
                     '                                & lLastActivity & ",'" & Lots(k).LotSysId & "')"
                     
                     sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                            & "INADATE,INAQTY," & vbCrLf _
                            & "INAMT,INTOTMATL,INTOTLABOR,INTOTEXP,INTOTOH,INTOTHRS," & vbCrLf _
                            & "INSONUMBER,INSOITEM,INSOREV," & vbCrLf _
                            & "INCREDITACCT,INDEBITACCT,INNUMBER,INLOTNUMBER,INUNITS) " & vbCrLf _
                            & "VALUES(" & bInType & ",'" & sPart & "','RETURN','','" _
                            & sNow & "'," & (lots(K).LotSelQty) & "," & vbCrLf _
                            & cUnitCost & "," & cMatl & "," & vbCrLf _
                            & cLabor & "," & cExp & "," & vbCrLf _
                            & cOvHd & "," & cHours & "," & vbCrLf _
                            & lSo & "," & nItem & ",'" & sRev & "','" _
                            & Trim(!INDEBITACCT) & "','" & Trim(!INCREDITACCT) & "'," _
                            & lLastActivity & ",'" & lots(K).LotSysId & "','" _
                            & sUnits & "')"
                     
                     Debug.Print sSql
                                          
                     clsADOCon.ExecuteSQL sSql
                  
                  End With
               Next
               
               'adjust inventory if not lot tracked
            Else
               lLastActivity = lLastActivity + 1
               With rdoAct
                  '                        sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                  '                            & "INADATE,INAQTY,INAMT,INSONUMBER,INSOITEM,INSOREV," _
                  '                            & "INCREDITACCT,INDEBITACCT,INNUMBER) " _
                  '                            & "VALUES(" & bInType & ",'" & sPart & "','RETURN','" & lInvoice & "','" _
                  '                            & sNow & "'," & cReturnQty & "," & !INAMT & "," _
                  '                            & lSo & "," & nItem & ",'" & sRev & "','" _
                  '                            & Trim(!INDEBITACCT) & "','" & Trim(!INCREDITACCT) & "'," & lLastActivity & ")"
                  
                  cMatl = cUnitMatl * cReturnQty
                  cLabor = cUnitLabor * cReturnQty
                  cExp = cUnitExp * cReturnQty
                  cOvHd = cUnitOvHd * cReturnQty
                  cHours = cUnitHrs * cReturnQty
                                                  
                  sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                         & "INADATE,INAQTY," & vbCrLf _
                         & "INAMT,INTOTMATL,INTOTLABOR,INTOTEXP,INTOTOH,INTOTHRS," & vbCrLf _
                         & "INSONUMBER,INSOITEM,INSOREV," _
                         & "INCREDITACCT,INDEBITACCT,INNUMBER,INUNITS) " _
                         & "VALUES(" & bInType & ",'" & sPart & "','RETURN','" & lInvoice & "','" _
                         & sNow & "'," & cReturnQty & "," & vbCrLf _
                         & !INAMT & "," & cMatl & "," & vbCrLf _
                         & cLabor & "," & cExp & "," & vbCrLf _
                         & cOvHd & "," & cHours & "," & vbCrLf _
                         & lSo & "," & nItem & ",'" & sRev & "','" _
                         & Trim(!INDEBITACCT) & "','" & Trim(!INCREDITACCT) & "'," & lLastActivity & ",'" _
                         & sUnits & "')"
                  
                  Debug.Print sSql
                  
                  clsADOCon.ExecuteSQL sSql
                  
               End With
               
               ' Update the Lot qty's
               Dim strLotNum As String
               Dim bLotEntry As Boolean
               
               bLotEntry = GetPartLots(sPart, strLotNum)
               
               If (bLotEntry = True) Then
                  
                  Dim lLOTRECORD As Long
                  Dim strInvno As Long
                  
                  lLOTRECORD = GetNextLotRecord(strLotNum)
                  'insert lot transaction here
                  sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                         & "LOITYPE,LOIPARTREF,LOIQUANTITY,LOICUSTINVNO, LOICUST," _
                         & "LOIACTIVITY,LOICOMMENT) " _
                         & "VALUES('" & strLotNum & "'," _
                         & lLOTRECORD & "," & bInType & ",'" & sPart & "'," _
                         & Abs(cReturnQty) & "," & Val(cmbInv) & ",'" & Compress(cmbCst) & "'," _
                         & lLastActivity & ",'RETURN - " & lInvoice & "')"
                  clsADOCon.ExecuteSQL sSql
                  
                  'Update Lot Header
                  sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY+" _
                         & Abs(cReturnQty) & " WHERE LOTNUMBER='" & strLotNum & "'"
                  clsADOCon.ExecuteSQL sSql
               End If
               
            End If
            Set rdoAct = Nothing
            If sLot <> "" Then
               sLot = sLot & " "
            End If
            
            ' Update part QOH
            sSql = "UPDATE PartTable SET PAQOH=PAQOH+" & Val(cReturnQty) & ", " _
                     & "PALOTQTYREMAINING = PALOTQTYREMAINING + " & Val(cReturnQty) & " " _
                   & "WHERE PARTREF='" & sPart & "' "
            clsADOCon.ExecuteSQL sSql
            
            AverageCost sPart
            
            'INVOICE ITEM PART QTY
            sComments = sComments _
                        & "Invoice " & lblPre & cmbInv & " " _
                        & "Item " & nItem & sRev & " " _
                        & Left(vItems(b, IT_PART), 30) & " " _
                        & sLot _
                        & "Qty " & Format(cReturnQty, "0.000") _
                        & " @ " & Format(vItems(b, IT_UNITPRICE), "0.0000") _
                        & " = " & Format(revenue, "$#,##0.00") _
                        & " " & lblPO & vbCrLf
            'MsgBox sComments
         End If
         
         ' Credit A/R (-) if CM, debit if DM
         ' Debit Revenue (+) if CM, credit if DM
         If cmbTyp = "CM" Then
            gl.AddDebitCredit 0, revenue, cboARAccount, sPart, lSo, nItem, sRev, sCust
            gl.AddDebitCredit revenue, 0, CStr(vItems(b, IT_REVACCT)), sPart, lSo, nItem, sRev, sCust
         Else
            gl.AddDebitCredit revenue, 0, cboARAccount, sPart, lSo, nItem, sRev, sCust
            gl.AddDebitCredit 0, revenue, CStr(vItems(b, IT_REVACCT)), sPart, lSo, nItem, sRev, sCust
         End If
      End If
   Next
   
   'create freight debit and credit if required
   Dim cFreightAmount As Currency
   Dim sAccount As String
   If IsNumeric(txtFrt) Then
      cFreightAmount = CCur(txtFrt)
   Else
      cFreightAmount = 0
   End If
   If cFreightAmount > 0 Then
      sAccount = cboFreightAccount
      ' credit A/R for freight amount if CM debit if DM
      ' debit freight if CM, credit if DM
      If cmbTyp = "CM" Then
         gl.AddDebitCredit 0, cFreightAmount, cboARAccount, "", lSo, 0, "", sCust
         gl.AddDebitCredit cFreightAmount, 0, sAccount, "", lSo, 0, "", sCust
      Else
         gl.AddDebitCredit cFreightAmount, 0, cboARAccount, "", lSo, 0, "", sCust
         gl.AddDebitCredit 0, cFreightAmount, sAccount, "", lSo, 0, "", sCust
      End If
   End If
   
   'create sales tax debit and credit if required
   Dim cTaxAmount As Currency
   If IsNumeric(txtTax) Then
      cTaxAmount = CCur(txtTax)
   Else
      cTaxAmount = 0
   End If
   If cTaxAmount > 0 Then
      sAccount = cboTaxAccount
      ' credit A/R for tax amount if CM, else debit
      ' debit taxes payable if CM, else credit
      If cmbTyp = "CM" Then
         gl.AddDebitCredit 0, cTaxAmount, cboARAccount, "", lSo, 0, "", sCust
         gl.AddDebitCredit cTaxAmount, 0, sAccount, "", lSo, 0, "", sCust
      Else
         gl.AddDebitCredit cTaxAmount, 0, cboARAccount, "", lSo, 0, "", sCust
         gl.AddDebitCredit 0, cTaxAmount, sAccount, "", lSo, 0, "", sCust
      End If
   End If
   
   'create federal tax debit and credit if required
   Dim cFedTaxAmount As Currency
   If IsNumeric(txtFedTax) Then
      cFedTaxAmount = CCur(txtFedTax)
   Else
      cFedTaxAmount = 0
   End If
   If cFedTaxAmount > 0 Then
      sAccount = cboFedTaxAccount
      ' credit A/R for fed tax amount if CM, else debit
      If cmbTyp = "CM" Then
         gl.AddDebitCredit 0, cFedTaxAmount, cboARAccount, "", lSo, 0, "", sCust
         gl.AddDebitCredit cFedTaxAmount, 0, sAccount, "", lSo, 0, "", sCust
      Else
         gl.AddDebitCredit cFedTaxAmount, 0, cboARAccount, "", lSo, 0, "", sCust
         gl.AddDebitCredit 0, cFedTaxAmount, sAccount, "", lSo, 0, "", sCust
      End If
   End If
   
   'create bulk revenue adjustment if required
   Dim cBulkRevenueAmount As Currency
   If IsNumeric(txtRevenue) Then
      cBulkRevenueAmount = CCur(txtRevenue)
   Else
      cBulkRevenueAmount = 0
   End If
   If cBulkRevenueAmount > 0 Then
      sAccount = cboRevenueAccount
      ' credit A/R for revenue amount if CM, else debit
      ' debit revenue if CM, else credit
      If cmbTyp = "CM" Then
         gl.AddDebitCredit 0, cBulkRevenueAmount, cboARAccount, "", lSo, 0, "", sCust
         gl.AddDebitCredit cBulkRevenueAmount, 0, sAccount, "", lSo, 0, "", sCust
      Else
         gl.AddDebitCredit cBulkRevenueAmount, 0, cboARAccount, "", lSo, 0, "", sCust
         gl.AddDebitCredit 0, cBulkRevenueAmount, sAccount, "", lSo, 0, "", sCust
      End If
      
      sComments = sComments & "Credit includes " & Format(cBulkRevenueAmount, "$#,##0.00") & " adjustment." & vbCrLf
   End If
   
   ' Create Invoice
   Dim nDebCred As Integer
   If cmbTyp.Text = "CM" Then
      nDebCred = -1
   Else
      nDebCred = 1
   End If
   
   Dim sARAcct As String
   sARAcct = cboARAccount.Text
   
   '    sSql = "INSERT INTO CihdTable (INVNO,INVPRE,INVTYPE,INVSHIPDATE," _
   '        & "INVDATE,INVTOTAL,INVTAX,INVFREIGHT,INVCUST,INVCOMMENTS,INVORIGIN) VALUES (" & lInvoice _
   '        & "," & "'" & sPre & "','" & cmbTyp.Text & "','" & sNow & "','" & txtDte & "'," & nDebCred * CCur(lblTot) _
   '        & "," & nDebCred * cTaxAmount & "," & nDebCred * cFreightAmount _
   '        & ",'" & sCust & "','" & sComments & "','" & cmbInv & "')"
   sSql = "INSERT INTO CihdTable (INVNO,INVPRE,INVTYPE,INVSHIPDATE," _
          & "INVDATE,INVTOTAL,INVTAX,INVFEDTAX,INVFREIGHT,INVCUST,INVCOMMENTS,INVPRIORINV, INVARACCT, INVREASONS) VALUES (" & lInvoice _
          & "," & "'" & sPre & "','" & cmbTyp.Text & "','" & sNow & "','" & txtDte & "'," & nDebCred * CCur(lblTot) _
          & "," & nDebCred * cTaxAmount & "," & nDebCred * cFedTaxAmount & "," & nDebCred * cFreightAmount _
          & ",'" & sCust & "','" & sComments & "'," & cmbInv & ",'" & sARAcct & "','" & Trim(txtCmt) & "')"
   
   Debug.Print sSql
   
   clsADOCon.ExecuteSQL sSql
   
   Dim inv As New ClassARInvoice
   inv.SaveLastInvoiceNumber lInvoice
      
   If gl.Commit() And clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      If cmbTyp.Text = "CM" Then
         sMsg = "Credit Memo "
      Else
         sMsg = "Debit Memo "
      End If
      sMsg = sMsg & sPre & Format(lInvoice, "000000") & " Created"
      SysMsg sMsg, True
      ManageBoxs False
      cmbInv.SetFocus
   Else
      clsADOCon.RollbackTrans
      If clsADOCon.ADOErrNum <> 0 Then
         clsADOCon.ADOErrNum = 0
         Err.Clear
         Resume DiaErr1
      End If
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "createin"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub UpdateTotals()
   Dim i As Integer
   Dim cTotal As Currency
   
   cTotal = 0
   On Error Resume Next
   For i = 1 To iTotalItems
      If vItems(i, IT_RETURNQTY) > 0 Then
         'adjust for returned items
         cTotal = cTotal + (vItems(i, IT_RETURNQTY) * vItems(i, IT_UNITPRICE))
      End If
      
      If vItems(i, IT_ADJUST) > 0 Then
         'adjust remaining items for price
         cTotal = cTotal + ((vItems(i, IT_QTY) - vItems(i, IT_RETURNQTY)) * vItems(i, IT_ADJUST))
      End If
   Next
   
   'display total revenue adjustment from line items
   lblRevenue = Format(cTotal, CURRENCYMASK)
   
   cTotal = cTotal + CCur("0" & txtRevenue)
   cTotal = cTotal + CCur("0" & txtTax)
   cTotal = cTotal + CCur("0" & txtFedTax)
   cTotal = cTotal + CCur("0" & txtFrt)
   
   lblTot = Format(cTotal, CURRENCYMASK)
   If cTotal > 0 Then
      cmdPst.enabled = True
   Else
      cmdPst.enabled = False
   End If
End Sub

Private Sub txtRevenue_LostFocus()
   'make sure number is valid
   If Not CurrencyTextLostFocus(txtRevenue) Then
      Exit Sub
   End If
   
   Dim cFromItems As Currency
   Dim cFromBulkAdjustment As Currency
   On Error Resume Next
   cFromItems = CCur(lblRevenue)
   cFromBulkAdjustment = CCur(txtRevenue)
   
   If IsRevenueTotalOK Then
      UpdateTotals
   Else
      txtRevenue.SetFocus
   End If
End Sub

Private Function IsRevenueTotalOK() As Boolean
   
   Dim cFromItems As Currency
   Dim cFromBulkAdjustment As Currency
   On Error Resume Next
   cFromItems = CCur(lblRevenue)
   cFromBulkAdjustment = CCur(txtRevenue)
   
   If cFromItems + cFromBulkAdjustment > CCur(lblPriorRevenue) Then
      MsgBox "Can't credit more ( " & Format(cFromItems + cFromBulkAdjustment, "0.00") _
         & " ) than the original revenue amount: " & lblPriorRevenue
      IsRevenueTotalOK = False
      Exit Function
   End If
   
   IsRevenueTotalOK = True
   
End Function

Private Sub cboRevenueAccount_LostFocus()
   'don't try to validate if leaving form
   If Me.ActiveControl.Name = "cmdCan" Then Exit Sub
   
   'validate account
   Dim S As String
   On Error Resume Next
   S = Trim(cboRevenueAccount.Text)
   Dim gl As New GLTransaction
   If Not gl.IsValidRevenueAccount(S) Then
      cboRevenueAccount.SetFocus
   End If
End Sub

Private Sub txtTax_LostFocus()
   If Not CurrencyTextLostFocus(txtTax) Then
      Exit Sub
   End If
   
   Dim cAmt As Currency
   On Error Resume Next
   cAmt = CCur(txtTax)
   
   If cAmt > CCur(lblPriorSalesTax) And cmbTyp = "CM" Then
      MsgBox "Can't credit more than original tax amount: " & lblPriorSalesTax
      txtTax.SetFocus
      Exit Sub
   End If
   
   UpdateTotals
End Sub

Private Sub txtFedTax_LostFocus()
   If Not CurrencyTextLostFocus(txtFedTax) Then
      Exit Sub
   End If
   
   Dim cAmt As Currency
   On Error Resume Next
   cAmt = CCur(txtFedTax)
   
   If cAmt > CCur(lblPriorFedTax) And cmbTyp = "CM" Then
      MsgBox "Can't credit more than original Federal Tax amount: " & lblPriorFedTax
      txtFedTax.SetFocus
      Exit Sub
   End If
   
   UpdateTotals
End Sub

Private Function CurrencyTextLostFocus(txt As TextBox) As Boolean
   'returns true if currency field is valid
   If Trim(txt) = "" Then
      txt = "0.00"
      CurrencyTextLostFocus = True
      Exit Function
   End If
   
   On Error Resume Next
   Dim C As Currency
   C = CCur(txt.Text)
   If Err Then
      MsgBox "invalid currency amount:" & C
      txt.SetFocus
      CurrencyTextLostFocus = False
      Exit Function
   Else
      txt = Format(C, "0.00")
      CurrencyTextLostFocus = True
   End If
End Function

Private Sub cboTaxAccount_LostFocus()
   'don't try to validate if leaving form
   If Me.ActiveControl.Name = "cmdCan" Then Exit Sub
   
   CheckAccountInComboBox cboTaxAccount
End Sub

Private Sub cboFedTaxAccount_LostFocus()
   'don't try to validate if leaving form
   If Me.ActiveControl.Name = "cmdCan" Then Exit Sub
   
   CheckAccountInComboBox cboFedTaxAccount
End Sub

Private Function GetExtension(orderqty As Currency, returnqty, unitprice As Currency, adjustment As Currency)
   Dim revenue As Currency
   revenue = SARound(returnqty * unitprice, 2) 'return
   revenue = revenue + SARound((orderqty - returnqty) * adjustment, 2) 'unit price adjust
   GetExtension = revenue
End Function

Private Function GetInventoryCost(ByVal strLotNum As String, ByRef cUnitCost As Currency, _
            ByRef cUnitMatl As Currency, ByRef cUnitExp As Currency, ByRef cUnitLabor As Currency, _
            ByRef cUnitOvHd As Currency, ByRef cUnitHrs As Currency)
   
   Dim RdoLotCost As ADODB.Recordset
   
   sSql = "SELECT LOTUNITCOST,cast(LOTTOTMATL / LOTORIGINALQTY as decimal(12,4)) as TotMatl," _
            & "cast ( LOTTOTLABOR / LOTORIGINALQTY as decimal(12,4)) as TotLab," _
            & "cast ( LOTTOTEXP / LOTORIGINALQTY as decimal(12,4)) as TotExp," _
            & "cast ( LOTTOTOH / LOTORIGINALQTY as decimal(12,4)) as TotOH," _
            & "cast ( LOTTOTHRS / LOTORIGINALQTY as decimal(12,4)) as TotHrs " _
      & "From LohdTable WHERE LOTNUMBER = '" & strLotNum & "' AND LOTORIGINALQTY <> 0"
   
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLotCost, ES_FORWARD)
   If (bSqlRows = True) Then
      With RdoLotCost
         cUnitCost = !LotUnitCost
         cUnitOvHd = !TotOH
         cUnitLabor = !TotLab
         cUnitMatl = !TotMatl
         cUnitExp = !TotExp
         cUnitHrs = !TotHrs
      End With
      RdoLotCost.Close
   Else
      cUnitCost = 0
      cUnitOvHd = 0
      cUnitLabor = 0
      cUnitMatl = 0
      cUnitExp = 0
      cUnitHrs = 0
   End If
   
   Set RdoLotCost = Nothing
   
End Function

Private Function GetRevAccountFromCompany() As String
   Dim RdoLots As ADODB.Recordset
   Dim iList As Integer
   On Error GoTo DiaErr1
   
   sSql = "select * from ComnTable"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_STATIC)
   If bSqlRows Then
      With RdoLots
         GetRevAccountFromCompany = "" & Trim(!COCRREVACCT)
         ClearResultSet RdoLots
      End With
   Else
      GetRevAccountFromCompany = ""
   End If
   Set RdoLots = Nothing
   Exit Function
   
DiaErr1:
   GetRevAccountFromCompany = ""
   Set RdoLots = Nothing
   
End Function

Private Function GetPartLots(strPartRef As String, ByRef strLotNumber As String) As Boolean
   Dim RdoLots As ADODB.Recordset
   Dim iList As Integer
   On Error GoTo DiaErr1
   
   sSql = "SELECT TOP(1) LOTNUMBER,LOTPARTREF,LOTREMAININGQTY,LOTAVAILABLE " _
          & "FROM LohdTable WHERE LOTPARTREF='" & strPartRef & "' AND " _
          & "LOTAVAILABLE=1 ORDER BY LOTNUMBER ASC"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_STATIC)
   If bSqlRows Then
      With RdoLots
         Do Until .EOF
            strLotNumber = "" & Trim(!lotNumber)
            .MoveNext
         Loop
         ClearResultSet RdoLots
      End With
      GetPartLots = True
   Else
      GetPartLots = False
   End If
   Set RdoLots = Nothing
   Exit Function
   
DiaErr1:
   GetPartLots = False
   Set RdoLots = Nothing
   
End Function


