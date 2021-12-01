VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ShopSHe02d 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Order Allocations"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbSon 
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   72
      Top             =   4560
      Width           =   1395
   End
   Begin VB.TextBox txtQty 
      Height          =   315
      Index           =   5
      Left            =   5520
      TabIndex        =   71
      Top             =   4920
      Width           =   915
   End
   Begin VB.ComboBox cmbItm 
      Height          =   315
      Index           =   5
      Left            =   1560
      TabIndex        =   70
      Top             =   4560
      Width           =   735
   End
   Begin VB.ComboBox cmbSon 
      Height          =   315
      Index           =   6
      Left            =   120
      TabIndex        =   69
      Top             =   5280
      Width           =   1395
   End
   Begin VB.TextBox txtQty 
      Height          =   315
      Index           =   6
      Left            =   5520
      TabIndex        =   68
      Top             =   5640
      Width           =   915
   End
   Begin VB.ComboBox cmbItm 
      Height          =   315
      Index           =   6
      Left            =   1560
      TabIndex        =   67
      Top             =   5280
      Width           =   735
   End
   Begin VB.ComboBox cmbSon 
      Height          =   315
      Index           =   7
      Left            =   120
      TabIndex        =   66
      Top             =   6000
      Width           =   1395
   End
   Begin VB.TextBox txtQty 
      Height          =   315
      Index           =   7
      Left            =   5520
      TabIndex        =   65
      Top             =   6360
      Width           =   915
   End
   Begin VB.ComboBox cmbItm 
      Height          =   315
      Index           =   7
      Left            =   1560
      TabIndex        =   64
      Top             =   6000
      Width           =   735
   End
   Begin VB.ComboBox cmbSon 
      Height          =   315
      Index           =   8
      Left            =   120
      TabIndex        =   63
      Top             =   6720
      Width           =   1395
   End
   Begin VB.TextBox txtQty 
      Height          =   315
      Index           =   8
      Left            =   5520
      TabIndex        =   62
      Top             =   7080
      Width           =   915
   End
   Begin VB.ComboBox cmbItm 
      Height          =   315
      Index           =   8
      Left            =   1560
      TabIndex        =   61
      Top             =   6720
      Width           =   735
   End
   Begin VB.ComboBox cmbSon 
      Height          =   315
      Index           =   9
      Left            =   120
      TabIndex        =   60
      Top             =   7440
      Width           =   1395
   End
   Begin VB.TextBox txtQty 
      Height          =   315
      Index           =   9
      Left            =   5520
      TabIndex        =   59
      Top             =   7800
      Width           =   915
   End
   Begin VB.ComboBox cmbItm 
      Height          =   315
      Index           =   9
      Left            =   1560
      TabIndex        =   58
      Top             =   7440
      Width           =   735
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHe02d.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   56
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optOvr 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Override Quantity Test And Allow Allocations To Be Greater Than The SO Item Quantity"
      Top             =   360
      Width           =   735
   End
   Begin VB.ComboBox cmbSon 
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   1395
   End
   Begin VB.TextBox txtQty 
      Height          =   315
      Index           =   4
      Left            =   5520
      TabIndex        =   15
      Top             =   4200
      Width           =   915
   End
   Begin VB.ComboBox cmbItm 
      Height          =   315
      Index           =   4
      Left            =   1560
      TabIndex        =   14
      Top             =   3840
      Width           =   735
   End
   Begin VB.ComboBox cmbSon 
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   1395
   End
   Begin VB.TextBox txtQty 
      Height          =   315
      Index           =   3
      Left            =   5520
      TabIndex        =   12
      Top             =   3480
      Width           =   915
   End
   Begin VB.ComboBox cmbItm 
      Height          =   315
      Index           =   3
      Left            =   1560
      TabIndex        =   11
      Top             =   3120
      Width           =   735
   End
   Begin VB.ComboBox cmbSon 
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1395
   End
   Begin VB.TextBox txtQty 
      Height          =   315
      Index           =   2
      Left            =   5520
      TabIndex        =   9
      Top             =   2760
      Width           =   915
   End
   Begin VB.ComboBox cmbItm 
      Height          =   315
      Index           =   2
      Left            =   1560
      TabIndex        =   8
      Top             =   2400
      Width           =   735
   End
   Begin VB.ComboBox cmbSon 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1395
   End
   Begin VB.TextBox txtQty 
      Height          =   315
      Index           =   1
      Left            =   5520
      TabIndex        =   6
      Top             =   2040
      Width           =   915
   End
   Begin VB.ComboBox cmbItm 
      Height          =   315
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "&Apply"
      Height          =   315
      Index           =   0
      Left            =   7800
      TabIndex        =   17
      ToolTipText     =   "Allocate Current Selections And Apply Changes"
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox cmbItm 
      Height          =   315
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Select Item From List"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtQty 
      Height          =   315
      Index           =   0
      Left            =   5520
      TabIndex        =   3
      Top             =   1320
      Width           =   915
   End
   Begin VB.ComboBox cmbSon 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Enter Or Select Sales Order "
      Top             =   960
      Width           =   1395
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   8160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8745
      FormDesignWidth =   8880
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   2520
      TabIndex        =   57
      Top             =   8280
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblAQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   6480
      TabIndex        =   108
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label lblAQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   6480
      TabIndex        =   107
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label lblAQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   6480
      TabIndex        =   106
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label lblAQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   6480
      TabIndex        =   105
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label lblAQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   6480
      TabIndex        =   104
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label lblAQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   6480
      TabIndex        =   103
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblAQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   6480
      TabIndex        =   102
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblAQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   6480
      TabIndex        =   101
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblAQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   6480
      TabIndex        =   100
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblAQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   0
      Left            =   6480
      TabIndex        =   99
      Top             =   960
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Allocated Qty   "
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
      Index           =   9
      Left            =   6480
      TabIndex        =   98
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblRev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   97
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lblItm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   96
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label LblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   2280
      TabIndex        =   95
      Top             =   4560
      Width           =   3195
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   2280
      TabIndex        =   94
      Top             =   4920
      Width           =   3195
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   5520
      TabIndex        =   93
      Top             =   4560
      Width           =   915
   End
   Begin VB.Label lblRev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   92
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label lblItm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   91
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label LblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   2280
      TabIndex        =   90
      Top             =   5280
      Width           =   3195
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   2280
      TabIndex        =   89
      Top             =   5640
      Width           =   3195
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   5520
      TabIndex        =   88
      Top             =   5280
      Width           =   915
   End
   Begin VB.Label lblRev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   1560
      TabIndex        =   87
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label lblItm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   86
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label LblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   2280
      TabIndex        =   85
      Top             =   6000
      Width           =   3195
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   2280
      TabIndex        =   84
      Top             =   6360
      Width           =   3195
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   5520
      TabIndex        =   83
      Top             =   6000
      Width           =   915
   End
   Begin VB.Label lblRev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   1560
      TabIndex        =   82
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label lblItm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   81
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label LblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   2280
      TabIndex        =   80
      Top             =   6720
      Width           =   3195
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   2280
      TabIndex        =   79
      Top             =   7080
      Width           =   3195
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   8
      Left            =   5520
      TabIndex        =   78
      Top             =   6720
      Width           =   915
   End
   Begin VB.Label lblRev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   1560
      TabIndex        =   77
      Top             =   7800
      Width           =   615
   End
   Begin VB.Label lblItm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   76
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label LblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   2280
      TabIndex        =   75
      Top             =   7440
      Width           =   3195
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   2280
      TabIndex        =   74
      Top             =   7800
      Width           =   3195
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   9
      Left            =   5520
      TabIndex        =   73
      Top             =   7440
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Allow Over Allocations"
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   55
      ToolTipText     =   "Override Quantity Test And Allow Allocations To Be Greater Than The SO Item Quantity"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected     "
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
      Left            =   7800
      TabIndex        =   54
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Qty      "
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
      Left            =   7800
      TabIndex        =   53
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblAllo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7800
      TabIndex        =   52
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblRqty 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7800
      TabIndex        =   51
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblRev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   50
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lblItm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   49
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label lblRev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   48
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lblItm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   47
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblRev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   46
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblItm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   45
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblRev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   44
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblItm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   43
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblRev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   42
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblItm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   41
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   5
      Left            =   4680
      TabIndex        =   40
      Top             =   0
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Number"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   39
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblRun 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5400
      TabIndex        =   38
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lblMon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   37
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label LblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   2280
      TabIndex        =   36
      Top             =   3840
      Width           =   3195
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   2280
      TabIndex        =   35
      Top             =   4200
      Width           =   3195
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   5520
      TabIndex        =   34
      Top             =   3840
      Width           =   915
   End
   Begin VB.Label LblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   2280
      TabIndex        =   33
      Top             =   3120
      Width           =   3195
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   2280
      TabIndex        =   32
      Top             =   3480
      Width           =   3195
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   5520
      TabIndex        =   31
      Top             =   3120
      Width           =   915
   End
   Begin VB.Label LblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   2280
      TabIndex        =   30
      Top             =   2400
      Width           =   3195
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   2280
      TabIndex        =   29
      Top             =   2760
      Width           =   3195
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   5520
      TabIndex        =   28
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label LblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   2280
      TabIndex        =   27
      Top             =   1680
      Width           =   3195
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   2280
      TabIndex        =   26
      Top             =   2040
      Width           =   3195
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   5520
      TabIndex        =   25
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity                "
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
      Left            =   5520
      TabIndex        =   24
      Top             =   720
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number/Description                                     "
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
      Left            =   2280
      TabIndex        =   23
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item         "
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
      Index           =   1
      Left            =   1560
      TabIndex        =   22
      Top             =   720
      Width           =   735
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   0
      Left            =   5520
      TabIndex        =   20
      Top             =   960
      Width           =   915
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   19
      Top             =   1320
      Width           =   3195
   End
   Begin VB.Label LblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   18
      Top             =   960
      Width           =   3195
   End
End
Attribute VB_Name = "ShopSHe02d"
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
Dim AdoQry1 As ADODB.Command
Dim AdoParameter11 As ADODB.Parameter

Dim AdoQry2 As ADODB.Command
Dim AdoParameter21 As ADODB.Parameter
Dim AdoParameter22 As ADODB.Parameter
Dim AdoParameter23 As ADODB.Parameter

Dim bDataChanged As Byte
Dim bGoodItem As Byte
Dim bGoodList As Byte
Dim bOnLoad As Byte
Dim iIndex As Integer

Dim sOldSo As String

Private Sub cmbItm_Change(Index As Integer)
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub cmbItm_Click(Index As Integer)
   Dim sSoItem As String
   sSoItem = Trim(cmbItm(Index))
   lblItm(Index) = Val(sSoItem)
   If Len(sSoItem) > 0 Then
      If Asc(Right(sSoItem, 1)) > 64 Then
         lblRev(Index) = Right(sSoItem, 1)
      Else
         lblRev(Index) = ""
      End If
   End If
   bGoodItem = GetThisItem()
   
End Sub

Private Sub cmbItm_GotFocus(Index As Integer)
   SelectFormat Me
   txtQty(Index).Enabled = False
   iIndex = Index
   
End Sub


Private Sub cmbItm_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub cmbItm_LostFocus(Index As Integer)
   Dim sSoItem As String
   sSoItem = Trim(cmbItm(Index))
   lblItm(Index) = Val(sSoItem)
   If Len(sSoItem) > 0 Then
      If Asc(Right(sSoItem, 1)) > 64 Then
         lblRev(Index) = Right(sSoItem, 1)
      Else
         lblRev(Index) = ""
      End If
   End If
   bGoodItem = GetThisItem()
   
'   Dim strSOAQty As String
'   Dim lSonumber As Long
'   Dim iItem As Integer
'
'   lSonumber = Val(Right(cmbSon(iIndex), SO_NUM_SIZE))
'   GetCurSOAllocQty CStr(lSonumber), CInt(sSoItem), strSOAQty
'   lblAQty(iIndex) = Format(strSOAQty, "#####0")
   
   On Error Resume Next
   If bGoodItem Then txtQty(Index).SetFocus
   
End Sub


Private Sub cmbSon_Change(Index As Integer)
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub cmbSon_Click(Index As Integer)
   iIndex = Index
   bGoodList = GetSoItems()
   
   '   Dim strSOAQty As String
'   Dim lSonumber As Long
'   Dim iItem As Integer
'
'   lSonumber = Val(Right(cmbSon(iIndex), SO_NUM_SIZE))
'   iItem = Val(cmbItm(iIndex))
'   GetCurSOAllocQty CStr(lSonumber), iItem, strSOAQty
'   lblAQty(iIndex) = Format(strSOAQty, "#####0")
'

End Sub

Private Sub cmbSon_GotFocus(Index As Integer)
   SelectFormat Me
   iIndex = Index
   If Len(Trim(cmbSon(Index))) = 0 Then txtQty(Index).Enabled = False
   sOldSo = cmbSon(Index)
   
End Sub

Private Sub cmbSon_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub cmbSon_LostFocus(Index As Integer)
   cmbSon(Index) = CheckLen(cmbSon(Index), 7)
   If Len(cmbSon(Index)) < 7 Then
      If Val(cmbSon(Index)) > 0 Then cmbSon(Index) = Format(Val(cmbSon(Index)), SO_NUM_FORMAT)
   End If
   If sOldSo <> cmbSon(Index) Then bGoodList = GetSoItems()
   If Len(Trim(cmbSon(Index))) = 0 Then
      cmbItm(Index) = ""
      lblQty(Index) = 0
      txtQty(Index) = 0
      GetSelected
   End If
   sOldSo = cmbSon(Index)
   
End Sub

Private Sub cmdAll_Click(Index As Integer)
   Dim b As Byte
   If optOvr.Value = vbUnchecked Then
      If Val(lblAllo) > Val(lblRqty) Then
         Beep
         MsgBox "The Quantity Allocated Is Greater Than " _
            & vbCr & "The MO Quantity.  You Must Adjust.", vbExclamation, Caption
         b = 0
      Else
         b = 1
      End If
   Else
      b = 1
   End If
   If b = 1 Then AllocateItems
   
End Sub

Private Sub cmdCan_Click()
   Form_Deactivate
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4170
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      optOvr = GetSetting("Esi2000", "EsiProd", "ShopSHe02d", optOvr)
      MouseCursor 13
      GetCurrentAllocations
      FillSalesOrders
      bOnLoad = 0
      bDataChanged = False
   End If
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Load()
   Dim iList As Integer
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   Move 300, 800
   For iList = 0 To 8
      lblItm(iList).Visible = False
      lblRev(iList).Visible = False
      txtQty(iList) = Format(0)
      txtQty(iList).Enabled = False
   Next
   lblItm(iList).Visible = False
   lblRev(iList).Visible = False
   txtQty(iList) = Format(0)
   txtQty(iList).Enabled = False
'   lblMon = ShopSHe02a.cmbPrt
'   lblRun = ShopSHe02a.cmbRun
'   lblRqty = ShopSHe02a.txtQty
   sSql = "SELECT DISTINCT SONUMBER,SOTYPE,SOTEXT,ITSO,ITNUMBER,ITREV,ITACTUAL FROM " _
          & "SohdTable,SoitTable WHERE SONUMBER=ITSO AND (SONUMBER= ?" _
          & " AND ITACTUAL IS NULL)"
   
   Set AdoQry1 = New ADODB.Command
   AdoQry1.CommandText = sSql
   
   Set AdoParameter11 = New ADODB.Parameter
   AdoParameter11.Type = adInteger
      
   AdoQry1.Parameters.Append AdoParameter11
   
   sSql = "SELECT ITSO,ITNUMBER,ITREV,ITPART,ITQTY,PARTREF," _
          & "PARTNUM,PADESC FROM SoitTable,PartTable WHERE " _
          & "ITPART=PARTREF AND (ITSO= ? AND ITNUMBER=  ? AND ITREV= ? ) " _
          & "AND ITCANCELED=0"
   Set AdoQry2 = New ADODB.Command
   AdoQry2.CommandText = sSql
   
   Set AdoParameter21 = New ADODB.Parameter
   AdoParameter21.Type = adInteger
   Set AdoParameter22 = New ADODB.Parameter
   AdoParameter22.Type = adInteger
   Set AdoParameter23 = New ADODB.Parameter
   AdoParameter23.Type = adChar
   AdoParameter23.SIZE = 2
   AdoQry2.Parameters.Append AdoParameter21
   AdoQry2.Parameters.Append AdoParameter22
   AdoQry2.Parameters.Append AdoParameter23
          
          
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveSetting "Esi2000", "EsiProd", "ShopSHe02d", optOvr
   ShopSHe02a.optAll.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set AdoParameter11 = Nothing
   Set AdoParameter21 = Nothing
   Set AdoParameter22 = Nothing
   Set AdoParameter23 = Nothing
   Set AdoQry1 = Nothing
   Set AdoQry2 = Nothing
   Set ShopSHe02d = Nothing
   
End Sub



Private Sub LblPrt_Change(Index As Integer)
   If Len(Trim(cmbSon(Index))) > 0 Then
      If Left(LblPrt(Index), 6) = "*** No" Then
         LblPrt(Index).ForeColor = ES_RED
      Else
         LblPrt(Index).ForeColor = Es_TextForeColor
      End If
   End If
   
End Sub

Private Sub optOvr_Click()
   'Allows for over allocation AWI 04/19/01
   
End Sub

Private Sub optOvr_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtQty_Change(Index As Integer)
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtQty_GotFocus(Index As Integer)
   SelectFormat Me
   iIndex = Index
   
End Sub



Private Sub FillSalesOrders()
   Dim RdoSon As ADODB.Recordset
   Dim iList As Integer
   Dim iRow As Integer
   Dim lLastSo As Long
   Dim sPartNumber As String
   
   sPartNumber = Compress(lblMon)

   On Error GoTo DiaErr1
   iList = 10
   sSql = "SELECT DISTINCT SONUMBER,SOTYPE,ITSO,ITNUMBER,ITACTUAL FROM " _
          & "SohdTable,SoitTable WHERE SONUMBER=ITSO AND ITACTUAL IS NULL" _
          & " AND ITPART LIKE '" & sPartNumber & "%'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon)
   If bSqlRows Then
      lLastSo = -99
      With RdoSon
         Do Until .EOF
            If lLastSo <> !SONUMBER Then
               iList = iList + 5
               If iList > 95 Then iList = 95
               prg1.Value = iList
               For iRow = 0 To 9
                  AddComboStr cmbSon(iRow).hwnd, "" & Trim(!SOTYPE) & Format$(!SONUMBER, SO_NUM_FORMAT)
               Next
            End If
            lLastSo = !SONUMBER
            .MoveNext
         Loop
         ClearResultSet RdoSon
      End With
   End If
   prg1.Value = 100
   Set RdoSon = Nothing
   prg1.Visible = False
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "fillsalesorders"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetSoItems() As Byte
   Dim RdoItm As ADODB.Recordset
   Dim lSonumber As Long
   Dim sSoItem As String
   
   If iIndex > 9 Then iIndex = 9
   lSonumber = Val(Right(cmbSon(iIndex), SO_NUM_SIZE))
   On Error GoTo DiaErr1
   cmbItm(iIndex).Clear
   AdoParameter11.Value = lSonumber
   bSqlRows = clsADOCon.GetQuerySet(RdoItm, AdoQry1, ES_KEYSET)
   If bSqlRows Then
      With RdoItm
         cmbItm(iIndex) = "" & str(!ITNUMBER) & Trim(!itrev)
         Do Until .EOF
            cmbSon(iIndex) = "" & Trim(!SOTYPE) & "" & Format$(Trim(!SoText), SO_NUM_FORMAT)
            AddComboStr cmbItm(iIndex).hwnd, "" & Format$(!ITNUMBER) & Trim(!itrev)
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
      sSoItem = Trim(cmbItm(iIndex))
      lblItm(iIndex) = Val(sSoItem)
      If Len(sSoItem) > 0 Then
         If Asc(Right(sSoItem, 1)) > 64 Then
            lblRev(iIndex) = Right(sSoItem, 1)
         Else
            lblRev(iIndex) = ""
         End If
      End If
      
      Dim strSOAQty As String
      
      GetCurSOAllocQty CStr(lSonumber), Val(sSoItem), strSOAQty
      lblAQty(iIndex) = Format(strSOAQty, "#####0")
   
      txtQty(iIndex).Enabled = True
      GetSoItems = True
      bGoodItem = GetThisItem
   Else
      GetSoItems = False
   End If
   Set RdoItm = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getsoitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetThisItem()
   Dim RdoItm As ADODB.Recordset
   Dim lSonumber As Long
   
   If iIndex > 9 Then iIndex = 9
   On Error GoTo DiaErr1
   lSonumber = Val(Right(cmbSon(iIndex), SO_NUM_SIZE))
   AdoParameter21.Value = lSonumber
   AdoParameter22.Value = Val(lblItm(iIndex))
   AdoParameter23.Value = "" & Trim(lblRev(iIndex))
   bSqlRows = clsADOCon.GetQuerySet(RdoItm, AdoQry2)
   If bSqlRows Then
      With RdoItm
         LblPrt(iIndex) = "" & !PartNum
         lblDsc(iIndex) = "" & !PADESC
         lblQty(iIndex) = Format(0 + !ITQTY, "####0")
         ClearResultSet RdoItm
      End With
      
      Dim strSOAQty As String
      
      GetCurSOAllocQty CStr(lSonumber), Val(lblItm(iIndex)), strSOAQty
      lblAQty(iIndex) = Format(strSOAQty, ES_QuantityDataFormat) ' MM changed format
      
      txtQty(iIndex).Enabled = True
      GetThisItem = True
   Else
      LblPrt(iIndex) = "*** No Valid Item Selected ***"
      lblDsc(iIndex) = ""
      lblQty(iIndex) = ""
      txtQty(iIndex).Enabled = False
      GetThisItem = False
   End If
   Set RdoItm = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getthisitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtQty_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtQty_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtQty_LostFocus(Index As Integer)
   txtQty(Index) = CheckLen(txtQty(Index), 6)
   If Val(txtQty(Index)) > Val(lblQty(Index)) Then
      If optOvr.Value = vbUnchecked Then
         Beep
         txtQty(Index) = lblQty(Index)
      End If
   End If
   txtQty(Index) = Format(Abs(Val(txtQty(Index))), ES_QuantityDataFormat)
   GetSelected
   
End Sub



Private Sub GetSelected()
   Dim iList As Integer
   Dim l As Long
   For iList = 0 To 8
      l = l + Val(txtQty(iList))
   Next
   l = l + Val(txtQty(iList))
   lblAllo = Format(l, "#####0")
   If l > Val(lblRqty) Then
      lblAllo.ForeColor = ES_RED
   Else
      lblAllo.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub GetCurrentAllocations()
   Dim RdoAll As ADODB.Recordset
   Dim iList As Integer
   Dim n As Integer
   Dim sPartNumber As String
   Dim lSon(25, 5) As Currency ' MM change
   
   sPartNumber = Compress(lblMon)
   On Error Resume Next
   prg1.Visible = True
   prg1.Value = 5
   On Error GoTo DiaErr1
   iList = -1
   sSql = "SELECT * FROM RnalTable WHERE RAREF='" & sPartNumber & "' " _
          & "AND RARUN=" & Val(lblRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAll, ES_STATIC)
   If bSqlRows Then
      With RdoAll
         Do Until .EOF
            iList = iList + 1
            If iList > 9 Then
               iList = 9
               Exit Do
            End If
            
            lSon(iList, 0) = Format(0 + !RASO, "####0")
            lSon(iList, 1) = Format(!RAQTY, ES_QuantityDataFormat) ' MM Change
            lblItm(iList) = Format(!RASOITEM, "##0")
            lblRev(iList) = "" & Trim(!RASOREV)
            .MoveNext
         Loop
         ClearResultSet RdoAll
         prg1.Value = 10
      End With
      For n = 0 To iList
         sSql = "Qry_GetSalesOrderType " & lSon(n, 0) & " "
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoAll, ES_FORWARD)
         If bSqlRows Then
            With RdoAll
               cmbSon(n) = "" & !SOTYPE & Format(lSon(n, 0), SO_NUM_FORMAT)
               txtQty(n) = Format(0 + lSon(n, 1), ES_QuantityDataFormat) ' MM change
               cmbItm(n) = lblItm(n) & Trim(lblRev(n))
               iIndex = n
               bGoodItem = GetThisItem()
               txtQty(n).Enabled = True
               ' Get Allocated Qty for this SO
               Dim strSOAQty As String
               GetCurSOAllocQty CStr(lSon(n, 0)), lblItm(n), strSOAQty
               lblAQty(n) = Format(strSOAQty, ES_QuantityDataFormat) ' MM change
               ClearResultSet RdoAll
            End With
         End If
      Next
      GetSelected
      iIndex = 0
   End If
   On Error Resume Next
   RdoAll.Close
   Set RdoAll = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcurrentall"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetCurSOAllocQty(strSONum As String, iItem As Integer, strSOAQty As String)
   
   strSOAQty = ""
   Dim RdoSOA As ADODB.Recordset
   sSql = "SELECT SUM(RAQTY) SumQty FROM RnalTable WHERE RASO='" & strSONum & "' " _
          & "AND RASOITEM=" & Val(iItem) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSOA, ES_STATIC)
   If bSqlRows Then
      With RdoSOA
         strSOAQty = Format(!SumQty, ES_QuantityDataFormat) ' MM Change Alloc
         ClearResultSet RdoSOA
         prg1.Value = 10
      End With
   End If
   
  Set RdoSOA = Nothing
  
End Sub

Private Sub AllocateItems()
   Dim bByte As Byte
   Dim iRow As Integer
   Dim bResponse As Integer
   Dim sMsg As String
   Dim sPartNumber As String
   Dim vItems(150, 7) As Variant
   
   sPartNumber = Compress(lblMon)
   For iRow = 0 To 9
      If Trim(cmbSon(iRow)) <> "" And Val(lblItm(iRow)) > 0 _
              And Val(txtQty(iRow)) > 0 Then
         bByte = True
         Exit For
      End If
   Next
   If Not bByte Then
      sMsg = "There Are No Items Selected To Allocate." & vbCr _
             & "Do You Wish To Remove Any That Existed?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         sSql = "DELETE FROM RnalTable WHERE RAREF='" & sPartNumber & "' AND " _
                & "RARUN=" & Val(lblRun) & " "
         clsADOCon.ExecuteSQL sSql
         SysMsg "Allocations Updated", True
         Unload Me
      Else
         CancelTrans
      End If
      Exit Sub
   End If
   If Not bDataChanged Then
      MsgBox "The Data Has Not Changed.", vbInformation, Caption
      Exit Sub
   End If
   sMsg = "Do You Want to Replace Any Current " & vbCr _
          & "Allocations With These New Allocations?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      On Error GoTo RvsoAl1
      sSql = "DELETE FROM RnalTable WHERE RAREF='" & sPartNumber & "' AND " _
             & "RARUN=" & Val(lblRun) & " "
      clsADOCon.ExecuteSQL sSql
      
      For iRow = 0 To 9
         If Len(Trim(cmbSon(iRow))) > 0 And _
                Val(lblItm(iRow)) > 0 And Val(txtQty(iRow)) > 0 Then
            vItems(iRow, 0) = Right(cmbSon(iRow), SO_NUM_SIZE)
            vItems(iRow, 1) = lblItm(iRow)
            vItems(iRow, 2) = "" & Trim(lblRev(iRow))
            vItems(iRow, 3) = txtQty(iRow)
         Else
            vItems(iRow, 3) = 0
         End If
      Next
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      For iRow = 0 To 9
         If Val(vItems(iRow, 3)) > 0 Then
            sSql = "INSERT INTO RnalTable (RAREF,RARUN,RASO," _
                   & "RASOITEM,RASOREV,RAQTY) VALUES('" _
                   & sPartNumber & "'," & Val(lblRun) & "," _
                   & Val(vItems(iRow, 0)) & "," _
                   & Val(vItems(iRow, 1)) & ",'" _
                   & vItems(iRow, 2) & "'," _
                   & Val(vItems(iRow, 3)) & ")"
            clsADOCon.ExecuteSQL sSql
         End If
      Next
      MouseCursor 0
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         MsgBox "Allocations Completed.", vbInformation, Caption
         Unload Me
      Else
         clsADOCon.RollbackTrans
         MsgBox "Could Not Complete Allocations. Make Certain  " & vbCr _
            & "Each Item Is Only Allocated Once To This MO.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
RvsoAl1:
   sProcName = "allocateit"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume RvsoAl2
RvsoAl2:
   MouseCursor 0
   MsgBox CurrError.Description & vbCr _
      & "Couldn't Complete The Allocations.", vbExclamation, Caption
   DoModuleErrors Me
   
End Sub
