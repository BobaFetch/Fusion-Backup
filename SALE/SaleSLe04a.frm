VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SaleSLe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Price Books"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPartClass 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "9"
      ToolTipText     =   "Select Product Class From List"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "SaleSLe04a.frx":0000
      DownPicture     =   "SaleSLe04a.frx":04F2
      Enabled         =   0   'False
      Height          =   372
      Left            =   7080
      MaskColor       =   &H00000000&
      Picture         =   "SaleSLe04a.frx":09E4
      Style           =   1  'Graphical
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   6000
      Width           =   400
   End
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "SaleSLe04a.frx":0ED6
      DownPicture     =   "SaleSLe04a.frx":13C8
      Enabled         =   0   'False
      Height          =   372
      Left            =   7080
      MaskColor       =   &H00000000&
      Picture         =   "SaleSLe04a.frx":18BA
      Style           =   1  'Graphical
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   6380
      Width           =   400
   End
   Begin VB.Frame z2 
      Height          =   40
      Left            =   240
      TabIndex        =   73
      Top             =   2760
      Width           =   7212
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLe04a.frx":1DAC
      Style           =   1  'Graphical
      TabIndex        =   72
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtMoq 
      Height          =   285
      Index           =   6
      Left            =   4320
      TabIndex        =   39
      Tag             =   "1"
      ToolTipText     =   "Minimum Order Quantity"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtMoq 
      Height          =   285
      Index           =   5
      Left            =   4320
      TabIndex        =   35
      Tag             =   "1"
      ToolTipText     =   "Minimum Order Quantity"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtMoq 
      Height          =   285
      Index           =   4
      Left            =   4320
      TabIndex        =   31
      Tag             =   "1"
      ToolTipText     =   "Minimum Order Quantity"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox txtMoq 
      Height          =   285
      Index           =   3
      Left            =   4320
      TabIndex        =   27
      Tag             =   "1"
      ToolTipText     =   "Minimum Order Quantity"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtMoq 
      Height          =   285
      Index           =   2
      Left            =   4320
      TabIndex        =   23
      Tag             =   "1"
      ToolTipText     =   "Minimum Order Quantity"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtMoq 
      Height          =   285
      Index           =   1
      Left            =   4320
      TabIndex        =   19
      Tag             =   "1"
      ToolTipText     =   "Minimum Order Quantity"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtMoq 
      Height          =   285
      Index           =   0
      Left            =   4320
      TabIndex        =   15
      Tag             =   "1"
      ToolTipText     =   "Minimum Order Quantity"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CheckBox opthlp 
      Caption         =   "Check1"
      Height          =   255
      Left            =   6240
      TabIndex        =   70
      Top             =   1620
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox optCst 
      Caption         =   "   "
      Height          =   255
      Left            =   6240
      TabIndex        =   69
      Top             =   1260
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox optSel 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      ToolTipText     =   "Show Only Price Book Items In Update List"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdCst 
      Caption         =   "C&ustomers"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6480
      TabIndex        =   67
      TabStop         =   0   'False
      ToolTipText     =   "Select Customers For This Price Book (Requires Valid Price Book)"
      Top             =   600
      Width           =   920
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Index           =   6
      Left            =   3120
      TabIndex        =   38
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Current Price"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtSell 
      Height          =   285
      Index           =   6
      Left            =   5520
      TabIndex        =   40
      Tag             =   "1"
      ToolTipText     =   "Book Price"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CheckBox optLst 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   6
      Left            =   6720
      TabIndex        =   41
      ToolTipText     =   "Add Or Keep On List"
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Index           =   5
      Left            =   3120
      TabIndex        =   34
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Current Price"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtSell 
      Height          =   285
      Index           =   5
      Left            =   5520
      TabIndex        =   36
      Tag             =   "1"
      ToolTipText     =   "Book Price"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CheckBox optLst 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   6720
      TabIndex        =   37
      ToolTipText     =   "Add Or Keep On List"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Index           =   4
      Left            =   3120
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Current Price"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox txtSell 
      Height          =   285
      Index           =   4
      Left            =   5520
      TabIndex        =   32
      Tag             =   "1"
      ToolTipText     =   "Book Price"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CheckBox optLst 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   6720
      TabIndex        =   33
      ToolTipText     =   "Add Or Keep On List"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Index           =   3
      Left            =   3120
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Current Price"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtSell 
      Height          =   285
      Index           =   3
      Left            =   5520
      TabIndex        =   28
      Tag             =   "1"
      ToolTipText     =   "Book Price"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CheckBox optLst 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   6720
      TabIndex        =   29
      ToolTipText     =   "Add Or Keep On List"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Index           =   2
      Left            =   3120
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Current Price"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtSell 
      Height          =   285
      Index           =   2
      Left            =   5520
      TabIndex        =   24
      Tag             =   "1"
      ToolTipText     =   "Book Price"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CheckBox optLst 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   6720
      TabIndex        =   25
      ToolTipText     =   "Add Or Keep On List"
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Current Price"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtSell 
      Height          =   285
      Index           =   1
      Left            =   5520
      TabIndex        =   20
      Tag             =   "1"
      ToolTipText     =   "Book Price"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CheckBox optLst 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   21
      ToolTipText     =   "Add Or Keep On List"
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5520
      TabIndex        =   13
      ToolTipText     =   "Cancel And Exit List"
      Top             =   2880
      Width           =   915
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6480
      TabIndex        =   12
      ToolTipText     =   "Update List And Apply Changes"
      Top             =   2880
      Width           =   915
   End
   Begin VB.CheckBox optLst 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   0
      Left            =   6720
      TabIndex        =   17
      ToolTipText     =   "Add Or Keep On List"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtSell 
      Height          =   285
      Index           =   0
      Left            =   5520
      TabIndex        =   16
      Tag             =   "1"
      ToolTipText     =   "Book Price"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Index           =   0
      Left            =   3120
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Current Price"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Height          =   315
      Left            =   4800
      TabIndex        =   10
      ToolTipText     =   "Select And Load Part Numbers"
      Top             =   2220
      Width           =   915
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Tag             =   "3"
      ToolTipText     =   "Select Parts (Leading Chars) Maximum 300 Part Numbers - Pattern Search >="
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CheckBox optInc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      ToolTipText     =   "Use The Discount To Calculate Book Price"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtDis 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Percentage"
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox txtEpr 
      Height          =   315
      Left            =   4200
      TabIndex        =   3
      Tag             =   "4"
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Tag             =   "4"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "2"
      ToolTipText     =   "(40) Char Max"
      Top             =   480
      Width           =   3475
   End
   Begin VB.ComboBox cmbPrb 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Price Book ID (12 Char Max)"
      Top             =   120
      Width           =   1800
   End
   Begin VB.ComboBox cmbProductCode 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "9"
      ToolTipText     =   "Select Product Class From List"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   920
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5280
      Top             =   6960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6810
      FormDesignWidth =   7545
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   240
      TabIndex        =   76
      Top             =   6360
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Class"
      Height          =   285
      Index           =   17
      Left            =   360
      TabIndex        =   77
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Line z3 
      Index           =   3
      X1              =   6720
      X2              =   7340
      Y1              =   3516
      Y2              =   3516
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   0
      Picture         =   "SaleSLe04a.frx":255A
      Top             =   6480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   0
      Picture         =   "SaleSLe04a.frx":2A4C
      Top             =   6120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   0
      Picture         =   "SaleSLe04a.frx":2F3E
      Top             =   6840
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   120
      Picture         =   "SaleSLe04a.frx":3430
      Top             =   7200
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Min Order Qty"
      Height          =   285
      Index           =   16
      Left            =   4320
      TabIndex        =   71
      Top             =   3240
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Selected Only"
      Height          =   285
      Index           =   14
      Left            =   3360
      TabIndex        =   68
      ToolTipText     =   "Show Only Price Book Items In Update List"
      Top             =   1560
      Width           =   1755
   End
   Begin VB.Label lblSelected 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6480
      TabIndex        =   66
      ToolTipText     =   "Number Of Items Selected"
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      Height          =   252
      Index           =   13
      Left            =   6120
      TabIndex        =   65
      Top             =   6360
      Width           =   372
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   252
      Index           =   12
      Left            =   5040
      TabIndex        =   64
      Top             =   6360
      Width           =   612
   End
   Begin VB.Label lblLst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   6480
      TabIndex        =   63
      Top             =   6360
      Width           =   372
   End
   Begin VB.Label lblPge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5520
      TabIndex        =   62
      Top             =   6360
      Width           =   372
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   61
      ToolTipText     =   "Part Number"
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   60
      ToolTipText     =   "Part Number"
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   59
      ToolTipText     =   "Part Number"
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   58
      ToolTipText     =   "Part Number"
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   57
      ToolTipText     =   "Part Number"
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   56
      ToolTipText     =   "Part Number"
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Line z3 
      Index           =   2
      X1              =   5520
      X2              =   6600
      Y1              =   3520
      Y2              =   3520
   End
   Begin VB.Line z3 
      Index           =   1
      X1              =   3120
      X2              =   5400
      Y1              =   3520
      Y2              =   3520
   End
   Begin VB.Line z3 
      Index           =   0
      X1              =   240
      X2              =   3000
      Y1              =   3520
      Y2              =   3520
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include"
      Height          =   285
      Index           =   9
      Left            =   6720
      TabIndex        =   55
      Top             =   3260
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Price"
      Height          =   285
      Index           =   11
      Left            =   5520
      TabIndex        =   54
      Top             =   3260
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   285
      Index           =   10
      Left            =   3120
      TabIndex        =   53
      Top             =   3260
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   52
      Top             =   3260
      Width           =   1515
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   51
      ToolTipText     =   "Part Number"
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   7
      Left            =   360
      TabIndex        =   50
      Top             =   2280
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Discount"
      Height          =   285
      Index           =   6
      Left            =   3360
      TabIndex        =   49
      ToolTipText     =   "Use The Discount To Calculate Book Price"
      Top             =   1200
      Width           =   1635
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   285
      Index           =   5
      Left            =   2640
      TabIndex        =   48
      Top             =   1200
      Width           =   435
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   285
      Index           =   4
      Left            =   360
      TabIndex        =   47
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Expires"
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   46
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Effective Date"
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   45
      Top             =   840
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   44
      Top             =   480
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price Book ID"
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   43
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   285
      Index           =   15
      Left            =   360
      TabIndex        =   42
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "SaleSLe04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/15/03 Added Minimum Order Quantity
'11/16/05 Removed sSql msgbox, reformatted form, added Get/Save settings
'5/9/06 Fixed Price Books query
'8/31/06 Locked cmdCst until cmbPrb.LostFocus
'9/8/06 Added Commit/RollBack to Update Entries
'9/26/06 Removed TabStops from txtPrice and changed BackColor
'        Added cPrices (currency array) for more accurate pricing(?)-PROPLA
'9/27/06 Added discount to optLst and txtMoq
'9/28/06 PROPLA - Added custom column to track changes (UpDateEntries)
Option Explicit
Dim RdoBok As ADODB.Recordset

Dim bCanceled As Byte
Dim bGoodBook As Byte
Dim bOnLoad As Byte
Dim bSelSaved As Byte

Dim iIndex As Integer
Dim iCurrPage As Integer
Dim iMaxPages As Integer
Dim iTotalItems As Integer

Dim cDiscount As Currency
Dim sClass As String

Dim vOldDate As Variant

Dim cPrices(300, 3) As Currency
Dim vPriceBook(300, 7) As Variant
'0 = Compressed Part
'1 = PartNumber
'2 = Description Used in ToolTipText
'3 = Part price (PAPRICE)
'4 = Book price
'5 = 0/1 value of checked
'6 = Minimun Order Quantity

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
'check for null dates

Private Sub CheckPriceBookDates()
   Dim vDate1 As Variant
   Dim vDate2 As Variant
   On Error Resume Next
   vDate1 = Format(ES_SYSDATE, "mm/dd/yyyy")
   vDate2 = Format(ES_SYSDATE + 365, "mm/dd/yyyy")
   sSql = "UPDATE PbhdTable SET PBHSTARTDATE='" & vDate1 & "'," _
          & "PBHENDDATE='" & vDate2 & "' WHERE (PBHSTARTDATE IS NULL " _
          & "OR PBHENDDATE IS NULL)"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub


Private Sub cmbPartClass_KeyPress(KeyAscii As Integer)
   
   'if backspace or space, go to first entry (ALL)
   If KeyAscii = 8 Or KeyAscii = 32 Then
      cmbPartClass.ListIndex = 0
   End If

End Sub

Private Sub cmbProductCode_KeyPress(KeyAscii As Integer)
   
   'if backspace or space, go to first entry (ALL)
   If KeyAscii = 8 Or KeyAscii = 32 Then
      cmbProductCode.ListIndex = 0
   End If

End Sub

Private Sub cmbProductCode_Validate(Cancel As Boolean)
   Dim b As Byte
   Dim iList As Integer
   
   If Trim(cmbProductCode) = "" Then
      'cmbProductCode = BRACKET_ALL
   End If
   For iList = 0 To cmbProductCode.ListCount - 1
      If cmbProductCode = cmbProductCode.List(iList) Then b = 1
   Next
   If b = 0 Then
      If cmbProductCode.ListCount > 0 Then
         Beep
         cmbProductCode = cmbProductCode.List(0)
      End If
   End If
   If bGoodBook Then
      On Error Resume Next
      'RdoBok.Edit
      RdoBok!PBHCLASS = cmbProductCode
      RdoBok!PBHREVDATE = Format(ES_SYSDATE, "mm/dd/yyyy hh:mm")
      RdoBok!PBHREVISEDBY = sInitials
      RdoBok.Update
   End If
   
End Sub



Private Sub cmbPrb_Click()
   bGoodBook = GetPriceBookId()
   'If bGoodBook Then cmdCst.Enabled = True
End Sub


Private Sub cmbPrb_LostFocus()
   cmbPrb = CheckLen(cmbPrb, 12)
   If bCanceled = 0 Then
      If Len(Trim(cmbPrb)) > 0 Then
         bGoodBook = GetPriceBookId()
         If bGoodBook = 0 Then
            AddPriceBookId
         Else
            'cmdCst.Enabled = True
         End If
      End If
   End If
   
End Sub

Private Sub cmdCan_Click()
   bCanceled = 1
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   bCanceled = 1
   
End Sub


Private Sub cmdCst_Click()
   If bGoodBook = 1 Then
      On Error Resume Next
      'RdoBok.Edit
      RdoBok!PBHREVDATE = Format(ES_SYSDATE, "mm/dd/yyyy hh:mm")
      RdoBok!PBHREVISEDBY = sInitials
      RdoBok!PBHCUSTOMERREVISEDBY = sInitials
      RdoBok.Update
      optCst.Value = vbChecked
      SaleSLe04b.Show
   Else
      MsgBox "Requires A Valid Price Book.", _
         vbInformation, Caption
   End If
   
End Sub

Private Sub cmdDn_Click()
   iCurrPage = iCurrPage + 1
   GetNextGroup
   
End Sub

Private Sub cmdDn_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub cmdEnd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If bSelSaved = 0 Then
      sMsg = "Any Changes Made Were Not Saved. Are You   " & vbCrLf _
             & "Certain That You Want To Cancel And Exit?...  "
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         bSelSaved = 1
         ManageUpper False
         cmbPrb.SetFocus
      End If
   Else
      bSelSaved = 1
      ManageUpper False
      cmbPrb.SetFocus
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2104
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdSel_Click()
   If Trim(cmbProductCode) = "" Then
      MsgBox "Requires A Valid Product Class.", _
         vbInformation, Caption
      Exit Sub
   End If
   If bGoodBook = 0 Then
      MsgBox "Requires A Valid Price Book.", _
         vbInformation, Caption
   Else
      GetPriceBook
   End If
   
   
   
End Sub

Private Sub cmdUp_Click()
   iCurrPage = iCurrPage - 1
   If iCurrPage < 1 Then iCurrPage = 1
   GetNextGroup
   
End Sub

Private Sub cmdUp_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub cmdUpd_Click()
   UpdateEntries
   
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CloseBoxes
'      FillProductClasses
'      If cmbProductCode.ListCount > 0 Then
'         cmbProductCode.AddItem "ALL"
'         For b = 0 To cmbProductCode.ListCount - 1
'            If cmbProductCode.List(b) = "NONE" Then cmbProductCode.RemoveItem b
'         Next
'         If cmbProductCode.ListCount > 0 Then cmbProductCode = cmbProductCode.List(0)
'      End If

      Dim prodcode As New ClassProductCode
      prodcode.PopulateProductCodeCombo cmbProductCode, True
      
      Dim partclass As New ClassPartClass
      partclass.PopulatePartClassCombo cmbPartClass, True
      
      FillCombo
      
      bOnLoad = 0
   End If
   If optCst.Value = vbChecked Then
      Unload SaleSLe04b
      optCst.Value = vbUnchecked
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetSettings
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   CheckPriceBookDates
   SaveSettings
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set SaleSLe04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtEpr = Format(Now + 365, "mm/dd/yyyy")
   txtDis = "0.000"
   cmdEnd.ToolTipText = "Cancel Work Not Updated And Return To Selection"
   cmdUp.ToolTipText = "Last Page (Page Up)"
   cmdDn.ToolTipText = "Next Page (Page Down)"
   For b = 0 To 6
      txtPrice(b).BackColor = BackColor
   Next
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillPriceBooks"
   LoadComboBox cmbPrb
   If cmbPrb.ListCount > 0 Then
        cmbPrb = cmbPrb.List(0)
        bGoodBook = GetPriceBookId()
        'If bGoodBook Then cmdCst.Enabled = True
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optCst_Click()
   'never visible. Customers on display
   
End Sub

Private Sub optInc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optLst_Click(Index As Integer)
   If optInc.Value = vbChecked Then
      If cDiscount > 0 And Val(txtSell(Index)) = 0 Then
         txtSell(Index) = Format(Val(txtPrice(Index)) * cDiscount, ES_QuantityDataFormat)
         vPriceBook(Index + iIndex, 4) = txtSell(Index)
         cPrices(Index + iIndex, 0) = Val(txtSell(Index))
      End If
   End If
   
End Sub

Private Sub optLst_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   vPriceBook(Index + iIndex, 5) = Val(optLst(Index).Value)
   txtSell(Index) = Format(Abs(Val(txtSell(Index))), ES_QuantityDataFormat)
   vPriceBook(Index + iIndex, 4) = txtSell(Index)
   cPrices(Index + iIndex, 0) = Val(txtSell(Index))
   
End Sub


Private Sub optLst_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub optLst_LostFocus(Index As Integer)
   vPriceBook(Index + iIndex, 5) = optLst(Index).Value
   
End Sub

Private Sub optLst_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   vPriceBook(Index + iIndex, 5) = Val(optLst(Index).Value)
   txtSell(Index) = Format(Abs(Val(txtSell(Index))), ES_QuantityDataFormat)
   vPriceBook(Index + iIndex, 4) = txtSell(Index)
   cPrices(Index + iIndex, 0) = Val(txtSell(Index))
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_Validate(Cancel As Boolean)
   txtBeg = CheckDateEx(txtBeg)
   If bGoodBook Then
      On Error Resume Next
      'RdoBok.Edit
      RdoBok!PBHSTARTDATE = txtBeg
      RdoBok!PBHREVDATE = Format(ES_SYSDATE, "mm/dd/yyyy hh:mm")
      RdoBok!PBHREVISEDBY = sInitials
      RdoBok.Update
   End If
   
End Sub


Private Sub txtDis_Validate(Cancel As Boolean)
   txtDis = CheckLen(txtDis, 6)
   txtDis = Format(Abs(Val(txtDis)), "#0.000")
   If bGoodBook Then
      On Error Resume Next
      'RdoBok.Edit
      RdoBok!PBHDISCOUNT = Val(txtDis)
      RdoBok.Update
   End If
   cDiscount = 1 - (Val(txtDis) / 100)
   If cDiscount = 1 Then optInc.Value = vbUnchecked
   
End Sub


Private Sub txtDsc_Validate(Cancel As Boolean)
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   If bGoodBook Then
      On Error Resume Next
      'RdoBok.Edit
      RdoBok!PBHDESC = txtDsc ' !PBHDESC
      RdoBok.Update
   End If
   
End Sub


Private Sub txtEpr_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEpr_Validate(Cancel As Boolean)
   txtEpr = CheckDateEx(txtEpr)
   If bGoodBook Then
      On Error Resume Next
      'RdoBok.Edit
      If Format(txtEpr, "mm/dd/yyyy") <> Format(vOldDate, "mm/dd/yyyy") Then
         RdoBok!PBHENDDATEREVISED = Format(ES_SYSDATE, "mm/dd/yyyy hh:mm")
         RdoBok!PBHENDDATEREVISEDBY = sInitials
      End If
      RdoBok!PBHENDDATE = Format(txtEpr, "mm/dd/yyyy")
      RdoBok!PBHREVDATE = Format(ES_SYSDATE, "mm/dd/yyyy hh:mm hh:mm")
      RdoBok!PBHREVISEDBY = sInitials
      RdoBok.Update
   End If
   
End Sub




Private Sub SaveSettings()
   Dim sOptions As String
   sOptions = str$(Trim(optInc.Value)) & str$(Trim(optSel.Value))
   SaveSetting "Esi2000", "EsiSale", "Cprbk", Trim(sOptions)
   
End Sub

Private Sub SetButtons()
   If iMaxPages = 1 Then
      cmdDn.Picture = Dsdn
      cmdDn.Enabled = False
      cmdUp.Picture = Dsup
      cmdUp.Enabled = False
   Else
      If iCurrPage > iMaxPages Then
         iCurrPage = iMaxPages
         cmdDn.Picture = Dsdn
         cmdDn.Enabled = False
      Else
         cmdDn.Picture = Endn
         cmdDn.Enabled = True
      End If
      If iCurrPage > 1 Then
         cmdUp.Picture = Enup
         cmdUp.Enabled = True
      Else
         cmdUp.Picture = Dsup
         cmdUp.Enabled = False
      End If
   End If
   
End Sub

Private Sub GetPriceBook()
   Dim RdoPbk As ADODB.Recordset
   Dim bLen As Byte
   Dim iRow As Integer
   Dim iList As Integer
   Dim C As Integer
   Dim sPartNumber As String
   
   sPartNumber = Compress(txtPrt)
   bLen = Len(sPartNumber)
   'If bLen = 0 Then bLen = 1
   Erase vPriceBook
   Erase cPrices
   iTotalItems = 0
   iIndex = 0
   CloseBoxes
   iRow = -1
   iList = -1
   'On Error GoTo 0
   
   'show all matching parts, not just those currently in the price book
   If optSel.Value = vbUnchecked Then
''      If cmbProductCode <> "ALL" Then
'''         sSql = "SELECT PARTREF,PARTNUM,PADESC,PAPRICE,PACLASS,PBIREF,PBIPARTREF,PBIPRICE," _
'''                & "PBIMINORDERQTY FROM PartTable,PbitTable WHERE (PARTREF*=PBIPARTREF) AND (PBIREF='" _
'''                & Compress(cmbPrb) & "') AND (LEFT(PARTREF," & bLen & ")>= '" & sPartNumber & "' AND " _
'''                & "PACLASS='" & Compress(cmbProductCode) & "')"
''         'no need for left join - where clause requires both rows to be non-null
''         sSql = "SELECT PARTREF,PARTNUM,PADESC,PAPRICE,PACLASS,PBIREF,PBIPARTREF,PBIPRICE," _
''                & "PBIMINORDERQTY" & vbCrLf _
''                & "FROM PartTable" & vbCrLf _
''                & "JOIN PbitTable ON PARTREF=PBIPARTREF" & vbCrLf _
''                & "WHERE PBIREF='" & Compress(cmbPrb) & "'" & vbCrLf _
''                & "AND PACLASS='" & Compress(cmbProductCode) & "'"
''      Else
'''         sSql = "SELECT PARTREF,PARTNUM,PADESC,PAPRICE,PACLASS,PBIREF,PBIPARTREF,PBIPRICE," _
'''                & "PBIMINORDERQTY FROM PartTable,PbitTable WHERE (PARTREF*=PBIPARTREF) AND (PBIREF='" _
'''                & Compress(cmbPrb) & "') AND (LEFT(PARTREF," & bLen & ")>= '" & sPartNumber & "')"
''         'no need for left join - where clause requires both rows to be non-null
''         sSql = "SELECT PARTREF,PARTNUM,PADESC,PAPRICE,PACLASS,PBIREF,PBIPARTREF,PBIPRICE," _
''                & "PBIMINORDERQTY" & vbCrLf _
''                & "FROM PartTable" & vbCrLf _
''                & "JOIN PbitTable ON PARTREF=PBIPARTREF" & vbCrLf _
''                & "WHERE PBIREF='" & Compress(cmbPrb) & "'" & vbCrLf
''                ' & "AND (LEFT(PARTREF," & bLen & ")>= '" & sPartNumber & "')"
''
''
''      End If
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PAPRICE,PACLASS,PBIREF,PBIPARTREF,PBIPRICE," _
         & "PBIMINORDERQTY" & vbCrLf _
         & "FROM PartTable" & vbCrLf _
         & "LEFT JOIN PbitTable ON PARTREF=PBIPARTREF AND PBIREF='" & Compress(cmbPrb) & "'" & vbCrLf _
         & "WHERE (PBIREF='" & Compress(cmbPrb) & "' OR PBIREF IS NULL)" & vbCrLf
'      If cmbProductCode <> BRACKET_ALL Then
'         sSql = sSql & "AND PACLASS='" & Compress(cmbProductCode) & "'"
'      End If
   
   
   'show current contents of price book only
   Else
'      sSql = "SELECT PARTREF,PARTNUM,PADESC,PAPRICE,PACLASS,PBIREF,PBIPARTREF,PBIPRICE," _
'             & "PBIMINORDERQTY FROM PartTable,PbitTable WHERE (PARTREF=PBIPARTREF) AND (PBIREF='" _
'             & Compress(cmbPrb) & "') AND (LEFT(PARTREF," & bLen & ")>= '" & sPartNumber & "')"
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PAPRICE,PACLASS,PBIREF,PBIPARTREF,PBIPRICE," _
             & "PBIMINORDERQTY" & vbCrLf _
             & "FROM PartTable" & vbCrLf _
             & "JOIN PbitTable ON PARTREF=PBIPARTREF" & vbCrLf _
             & "WHERE PBIREF='" & Compress(cmbPrb) & "'" & vbCrLf
                   
   End If
   
   'select matching product code if required
   If cmbProductCode <> BRACKET_ALL Then
      sSql = sSql & "AND PAPRODCODE='" & Compress(cmbProductCode) & "'" & vbCrLf
   End If
   
   'select matching part class if required
   If cmbPartClass <> BRACKET_ALL Then
      sSql = sSql & "AND PACLASS='" & Compress(cmbPartClass) & "'" & vbCrLf
   End If
   
   If bLen > 0 Then
      sSql = sSql & "AND PARTREF LIKE '" & sPartNumber & "%'" & vbCrLf
   End If
   sSql = sSql & "ORDER BY PARTREF"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPbk, ES_FORWARD)
   If bSqlRows Then
      With RdoPbk
         Do Until .EOF
            iRow = iRow + 1
            vPriceBook(iRow, 0) = "" & Trim(!PartRef)
            vPriceBook(iRow, 1) = "" & Trim(!PartNum)
            vPriceBook(iRow, 2) = "" & Trim(!PADESC)
            vPriceBook(iRow, 3) = Format(!PAPRICE, ES_QuantityDataFormat)
            cPrices(iRow, 2) = !PAPRICE
            If Not IsNull(!PBIPRICE) Then
               vPriceBook(iRow, 4) = Format(!PBIPRICE, ES_QuantityDataFormat)
               cPrices(iRow, 0) = !PBIPRICE
            Else
               vPriceBook(iRow, 4) = "0.000"
               cPrices(iRow, 0) = 0
            End If
            cPrices(iRow, 1) = cPrices(iRow, 0)
            If Not IsNull(!PBIPARTREF) Then
               vPriceBook(iRow, 5) = "1"
            Else
               vPriceBook(iRow, 5) = "0"
            End If
            
            If Not IsNull(!PBIMINORDERQTY) Then
               vPriceBook(iRow, 6) = Format(!PBIMINORDERQTY, ES_QuantityDataFormat)
            Else
               vPriceBook(iRow, 6) = "0.000"
            End If
            iList = iList + 1
            If iList < 7 Then
               lblPrt(iList).Enabled = True
               txtPrice(iList).Enabled = True
               txtSell(iList).Enabled = True
               optLst(iList).Enabled = True
               txtMoq(iList).Enabled = True
               txtPrice(iList) = vPriceBook(iRow, 3)
               txtSell(iList) = vPriceBook(iRow, 4)
               txtMoq(iList) = vPriceBook(iRow, 6)
               optLst(iList).Value = Val(vPriceBook(iRow, 5))
               lblPrt(iList) = vPriceBook(iRow, 1)
               lblPrt(iList).ToolTipText = vPriceBook(iRow, 2) & "     "
               lblPrt(iList).BackColor = Es_TextBackColor
               'txtPrice(iList).BackColor = Es_TextBackColor
               txtSell(iList).BackColor = Es_TextBackColor
               txtMoq(iList).BackColor = Es_TextBackColor
               optLst(iList).Caption = "____"
            End If
            C = C + 1
            If C > 0 Then
               lblSelected = iRow
               lblSelected.Refresh
               C = 0
            End If
            If iRow >= 299 Then Exit Do
            .MoveNext
         Loop
         ClearResultSet RdoPbk
      End With
   End If
   iIndex = 0
   iTotalItems = iRow + 1
   lblSelected = iTotalItems
   iCurrPage = 1
   iMaxPages = (iTotalItems / 7) + 0.498 'note rounding of int
   lblLst = iMaxPages
   SetButtons
   Set RdoPbk = Nothing
   'On Error Resume Next
   If iTotalItems > 0 Then
      bSelSaved = 0
      txtMoq(0).SetFocus
      ManageUpper True
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getpricebook"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetNextGroup()
   Dim bByte As Byte
   Dim iList As Integer
   SetButtons
   '4/17/02 Try to ensure capture of index
   For iList = 0 To 5
      vPriceBook(iList + iIndex, 5) = optLst(iList).Value
   Next
   vPriceBook(iList + iIndex, 5) = optLst(iList).Value
   
   iIndex = (iCurrPage - 1) * 7
   lblPge = iCurrPage
   'Clear it
   CloseBoxes
   On Error Resume Next
   For iList = 0 To 6
      If iList + iIndex >= iTotalItems Then Exit For
      lblPrt(iList) = vPriceBook(iList + iIndex, 1)
      If Trim(lblPrt(iList)) = "" Then Exit For
      lblPrt(iList).ToolTipText = vPriceBook(iList + iIndex, 2) & "     "
      txtPrice(iList) = vPriceBook(iList + iIndex, 3)
      'txtSell(iList) = vPriceBook(iList + iIndex, 4)
      txtSell(iList) = Format(cPrices(iList + iIndex, 0), ES_QuantityDataFormat)
      optLst(iList).Value = vPriceBook(iList + iIndex, 5)
      txtMoq(iList) = vPriceBook(iList + iIndex, 6)
      lblPrt(iList).BackColor = Es_TextBackColor
      'txtPrice(iList).BackColor = Es_TextBackColor
      txtMoq(iList).BackColor = Es_TextBackColor
      txtSell(iList).BackColor = Es_TextBackColor
      optLst(iList).Caption = "____"
      lblPrt(iList).Enabled = True
      txtPrice(iList).Enabled = True
      txtSell(iList).Enabled = True
      optLst(iList).Enabled = True
      txtMoq(iList).Enabled = True
   Next
   If txtMoq(0).Enabled Then txtMoq(0).SetFocus
   
End Sub

Private Sub txtMoq_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtMoq_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtMoq_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtMoq_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtMoq_LostFocus(Index As Integer)
   If optInc.Value = vbChecked Then
      If cDiscount > 0 And Val(txtSell(Index)) = 0 Then
         txtSell(Index) = Format(Val(txtPrice(Index)) * cDiscount, ES_QuantityDataFormat)
         vPriceBook(Index + iIndex, 4) = txtSell(Index)
         cPrices(Index + iIndex, 0) = Val(txtSell(Index))
      End If
   End If
   txtMoq(Index) = CheckLen(txtMoq(Index), 9)
   txtMoq(Index) = Format(Abs(Val(txtMoq(Index))), ES_QuantityDataFormat)
   vPriceBook(Index + iIndex, 6) = txtMoq(Index)
   
End Sub

Private Sub txtPrice_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtPrice_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtPrice_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtPrice_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub txtPrice_LostFocus(Index As Integer)
   If optInc.Value = vbChecked Then
      If cDiscount > 0 And Val(txtSell(Index)) = 0 Then
         txtSell(Index) = Format(Val(txtPrice(Index)) * cDiscount, ES_QuantityDataFormat)
         vPriceBook(Index + iIndex, 4) = txtSell(Index)
         cPrices(Index + iIndex, 0) = Val(txtSell(Index))
      End If
   End If
   
End Sub

Private Sub txtPrice_Validate(Index As Integer, Cancel As Boolean)
   txtPrice(Index) = CheckLen(txtPrice(Index), 9)
   txtPrice(Index) = Format(Abs(Val(txtPrice(Index))), ES_QuantityDataFormat)
   vPriceBook(Index + iIndex, 3) = txtPrice(Index)
   
End Sub

Private Sub txtPrt_Validate(Cancel As Boolean)
   txtPrt = CheckLen(txtPrt, 30)
   
End Sub


Private Sub txtSell_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtSell_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtSell_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtSell_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub



Private Sub CloseBoxes()
   Dim iList As Integer
   For iList = 0 To 6
      lblPrt(iList).Enabled = False
      txtPrice(iList).Enabled = False
      txtSell(iList).Enabled = False
      txtMoq(iList).Enabled = False
      optLst(iList).Enabled = False
      lblPrt(iList).BackColor = Es_FormBackColor
      txtPrice(iList).BackColor = Es_FormBackColor
      txtSell(iList).BackColor = Es_FormBackColor
      txtMoq(iList).BackColor = Es_FormBackColor
      lblPrt(iList) = ""
      txtPrice(iList) = ""
      txtSell(iList) = ""
      txtMoq(iList) = ""
      optLst(iList).Caption = ""
      optLst(iList).Value = vbUnchecked
   Next
   
End Sub

Private Sub txtSell_LostFocus(Index As Integer)
   txtSell(Index) = CheckLen(txtSell(Index), 9)
   txtSell(Index) = Format(Abs(Val(txtSell(Index))), ES_QuantityDataFormat)
   vPriceBook(Index + iIndex, 4) = txtSell(Index)
   cPrices(Index + iIndex, 0) = Val(txtSell(Index))
   
End Sub

Private Sub ManageUpper(bClose As Boolean)
   If Not bClose Then
      lblSelected = 0
      cmbPrb.Enabled = True
      cmbPrb.BackColor = Es_TextBackColor
      txtDsc.Enabled = True
      txtDsc.BackColor = Es_TextBackColor
      txtBeg.Enabled = True
      txtBeg.BackColor = Es_TextBackColor
      txtEpr.Enabled = True
      txtEpr.BackColor = Es_TextBackColor
      txtDis.Enabled = True
      txtDis.BackColor = Es_TextBackColor
      optInc.Enabled = True
      optInc.Caption = "____"
      optSel.Enabled = True
      optSel.Caption = "____"
      cmbProductCode.Enabled = True
      cmbProductCode.BackColor = Es_TextBackColor
      'cmbProductCode = sClass
      txtPrt.Enabled = True
      txtPrt.BackColor = Es_TextBackColor
      cmdSel.Enabled = True
      'cmdCst.Enabled = True
      cmdUpd.Enabled = False
      cmdEnd.Enabled = False
      CloseBoxes
   Else
      cmbPrb.Enabled = False
      cmbPrb.BackColor = Es_FormBackColor
      txtDsc.Enabled = False
      txtDsc.BackColor = Es_FormBackColor
      txtBeg.Enabled = False
      txtBeg.BackColor = Es_FormBackColor
      txtEpr.Enabled = False
      txtEpr.BackColor = Es_FormBackColor
      txtDis.Enabled = False
      txtDis.BackColor = Es_FormBackColor
      optInc.Enabled = False
      optInc.Caption = ""
      optSel.Enabled = False
      optSel.Caption = ""
      cmbProductCode.Enabled = False
      cmbProductCode.BackColor = Es_FormBackColor
      txtPrt.Enabled = False
      txtPrt.BackColor = Es_FormBackColor
      cmdSel.Enabled = False
      'cmdCst.Enabled = False
      cmdUpd.Enabled = True
      cmdEnd.Enabled = True
   End If
   
End Sub

Private Function GetPriceBookId() As Byte
   On Error GoTo DiaErr1
   bGoodBook = 0
   'cmdCst.Enabled = False
   sSql = "SELECT * FROM PbhdTable WHERE PBHREF='" & Compress(cmbPrb) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBok, ES_KEYSET)
   If bSqlRows Then
      With RdoBok
         cmbPrb = "" & Trim(!PBHID)
         txtDsc = "" & Trim(!PBHDESC)
         If Not IsNull(!PBHSTARTDATE) Then
            txtBeg = Format(!PBHSTARTDATE, "mm/dd/yyyy")
         Else
            txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         If Not IsNull(!PBHENDDATE) Then
            vOldDate = Format(!PBHENDDATE, "mm/dd/yyyy")
            txtEpr = "" & Format(!PBHENDDATE, "mm/dd/yyyy")
         Else
            vOldDate = "09/11/01"
            txtEpr = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         If Not IsNull(!PBHDISCOUNT) Then
            txtDis = "" & Format(!PBHDISCOUNT, "#0.000")
         Else
            txtDis = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         If Trim(!PBHCLASS) <> "" Then
            'cmbProductCode = "" & Trim(!PBHCLASS)
         Else
            'cmbProductCode = ""
         End If
         cDiscount = 1 - (!PBHDISCOUNT / 100)
         sClass = cmbProductCode
         GetPriceBookId = 1
      End With
      cmdCst.Enabled = True
   Else
      txtDsc = ""
      txtDis = "0.000"
      cmdCst.Enabled = False
      GetPriceBookId = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getpricebkid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddPriceBookId()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sNewId As String
   Dim sAddClass As String
   
   If cmbProductCode <> BRACKET_ALL Then sAddClass = Trim(cmbProductCode)
   bResponse = IllegalCharacters(cmbPrb)
   sNewId = Compress(cmbPrb)
   'cmdCst.Enabled = False
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   
   If bResponse > 0 Then
      MsgBox "Price Book ID Contains An Illegal " & Chr$(bResponse) & ".", _
         vbInformation, Caption
   Else
      sMsg = "Add The New Price Book " & cmbPrb & "?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         If sNewId = "ALL" Then
            Beep
            MsgBox "ALL Is An Illegal Price Book ID.", vbExclamation, Caption
            Exit Sub
         End If
         If Trim(cmbProductCode) = "" Then
            Beep
            MsgBox "Requires That You Have At Least One Product Class." & vbCrLf _
               & "Product Classes Are Added In The Administration Section.", _
               vbInformation, Caption
         Else
            txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
            txtEpr = Format(Now + 365, "mm/dd/yyyy")
            
            'Through with all of that so...
'            sSql = "INSERT INTO PbhdTable (PBHREF,PBHID,PBHSTARTDATE," _
'                   & "PBHENDDATE,PBHDISCOUNT,PBHCLASS) VALUES('" _
'                   & sNewId & "','" & cmbPrb & "','" & txtBeg & "','" _
'                   & txtEpr & "'," & Val(txtDis) & ",'" & sAddClass _
'                   & "') "
            sSql = "INSERT INTO PbhdTable (PBHREF,PBHID,PBHSTARTDATE," _
                   & "PBHENDDATE,PBHDISCOUNT) VALUES('" _
                   & sNewId & "','" & cmbPrb & "','" & txtBeg & "','" _
                   & txtEpr & "'," & Val(txtDis) & ") "
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
            If clsADOCon.ADOErrNum = 0 Then
               SysMsg "The Price Book Was Added.", True
               cmbPrb.AddItem cmbPrb
               bGoodBook = GetPriceBookId
            Else
               bGoodBook = 0
               MsgBox "Couldn't Add Price Book.", _
                  vbExclamation, Caption
            End If
            
         End If
         
      Else
         CancelTrans
      End If
   End If
   
End Sub

Private Sub UpdateEntries()
   Dim bResponse As Byte
   Dim sMsg As String
   
   Dim b As Integer
   Dim C As Integer
   Dim iRow As Integer
   
   Dim sPbRef As String
   Dim sPbId As String
   
   On Error Resume Next
   sMsg = "This Function While Add Or Update Book Items That Are" & vbCrLf _
          & "Checked. It Will Also Remove Items Have Been Unchecked." & vbCrLf _
          & "Do You Wish To Continue The Update?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      CancelTrans
      Exit Sub
   End If
   sPbRef = Compress(cmbPrb)
   sPbId = cmbPrb
   MouseCursor 13
   prg1.Visible = True
   C = 10
   prg1.Value = C
   If bGoodBook Then
      'RdoBok.Edit
      RdoBok!PBHREVDATE = Format(ES_SYSDATE, "mm/dd/yyyy hh:mm")
      RdoBok!PBHREVISEDBY = sInitials
      RdoBok.Update
   End If
   
   'Delete old entries to avoid duplicate errors
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   For iRow = 0 To iTotalItems - 1
      sSql = "DELETE FROM PbitTable WHERE (PBIREF='" _
             & sPbRef & "' AND PBIPARTREF='" & vPriceBook(iRow, 0) _
             & "')"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "DELETE FROM PbdtTable WHERE (PBDREF='" _
             & sPbRef & "' AND PBDPARTREF='" & vPriceBook(iRow, 0) _
             & "')"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      b = b + 1
      If b > 10 Then
         C = C + 10
         If C > 50 Then C = 50
         prg1.Value = C
         b = 0
      End If
   Next
   'Insert New Entries
   For iRow = 0 To iTotalItems - 1
      If Val(vPriceBook(iRow, 5)) = 1 Then
         sSql = "INSERT INTO PbitTable (PBIREF,PBIPARTREF," _
                & "PBIPRICE,PBIMINORDERQTY,PBIITEMSREVISEDBY) VALUES('" & sPbRef & "','" _
                & vPriceBook(iRow, 0) & "'," & cPrices(iRow, 0) _
                & "," & Val(vPriceBook(iRow, 6)) & ",'" & sInitials & "')"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
      End If
      b = b + 1
      If b > 10 Then
         C = C + 10
         If C > 90 Then C = 90
         prg1.Value = C
         b = 0
      End If
   Next
   '9/28/06 Propla only
   If ES_CUSTOM = "PROPLA" Or ES_CUSTOM = "ESINORTH" Then
      'Insert New Entries
      For iRow = 0 To iTotalItems - 1
         If Val(vPriceBook(iRow, 5)) = 1 Then
            If cPrices(iRow, 0) <> cPrices(iRow, 1) Then
               sSql = "INSERT INTO ProplaPBChanges (PBPART,PBBOOK," _
                      & "PBUNITPRICE,PBOLDPRICE,PBNEWPRICE,PBUSER) VALUES('" _
                      & vPriceBook(iRow, 0) & "','" & Compress(cmbPrb) & "'," _
                      & cPrices(iRow, 2) & "," & cPrices(iRow, 1) & "," _
                      & cPrices(iRow, 0) & ",'" & sInitials & "')"
               clsADOCon.ExecuteSQL sSql 'rdExecDirect
            End If
         End If
      Next
   End If
   prg1.Value = 100
   MouseCursor 0
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      SysMsg "Updated Successfully..", True
   Else
      clsADOCon.RollbackTrans
      MsgBox "The Price Book Could Not Be Updated and " & vbCrLf _
         & "And The Data Has Been Rolled Back.", _
         vbExclamation, Caption
   End If
   prg1.Visible = False
   bSelSaved = 1
   CloseBoxes
   ManageUpper False
   cmbPrb.SetFocus
   
End Sub


Public Sub GetSettings()
   Dim sOptions As String
   sOptions = GetSetting("Esi2000", "EsiSale", "Cprbk", Trim(sOptions))
   If Len(sOptions) Then
      optInc.Value = Val(Left(sOptions, 1))
      optSel.Value = Val(Right(sOptions, 1))
   End If
   
End Sub
