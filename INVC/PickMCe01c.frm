VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form PickMCe01c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pick List Items"
   ClientHeight    =   5745
   ClientLeft      =   1845
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAqt 
      Height          =   285
      Index           =   1
      Left            =   5160
      TabIndex        =   0
      ToolTipText     =   "Amount To Pick"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "PickMCe01c.frx":0000
      DownPicture     =   "PickMCe01c.frx":04F2
      Enabled         =   0   'False
      Height          =   372
      Left            =   7680
      MaskColor       =   &H00000000&
      Picture         =   "PickMCe01c.frx":09E4
      Style           =   1  'Graphical
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   4920
      Width           =   400
   End
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "PickMCe01c.frx":0ED6
      DownPicture     =   "PickMCe01c.frx":13C8
      Enabled         =   0   'False
      Height          =   372
      Left            =   7680
      MaskColor       =   &H00000000&
      Picture         =   "PickMCe01c.frx":18BA
      Style           =   1  'Graphical
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   5304
      Width           =   400
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PickMCe01c.frx":1DAC
      Style           =   1  'Graphical
      TabIndex        =   86
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   5400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5745
      FormDesignWidth =   8190
   End
   Begin VB.CommandButton cmdPck 
      Caption         =   "&Pick"
      Height          =   315
      Left            =   7200
      TabIndex        =   65
      ToolTipText     =   "Pick Selected Items"
      Top             =   450
      Width           =   915
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "type"
      Height          =   255
      Left            =   720
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAqt 
      Height          =   285
      Index           =   6
      Left            =   5160
      TabIndex        =   15
      ToolTipText     =   "Amount To Pick"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   6
      Left            =   6720
      TabIndex        =   16
      ToolTipText     =   "Work In Progress Location"
      Top             =   4560
      Width           =   735
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   6
      Left            =   7605
      TabIndex        =   17
      ToolTipText     =   "Check If Complete (Checked if More Than Rq'd)"
      Top             =   4560
      Width           =   420
   End
   Begin VB.TextBox txtAqt 
      Height          =   285
      Index           =   5
      Left            =   5160
      TabIndex        =   12
      ToolTipText     =   "Amount To Pick"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   5
      Left            =   6720
      TabIndex        =   13
      ToolTipText     =   "Work In Progress Location"
      Top             =   3840
      Width           =   735
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   5
      Left            =   7605
      TabIndex        =   14
      ToolTipText     =   "Check If Complete (Checked if More Than Rq'd)"
      Top             =   3840
      Width           =   420
   End
   Begin VB.TextBox txtAqt 
      Height          =   285
      Index           =   4
      Left            =   5160
      TabIndex        =   9
      ToolTipText     =   "Amount To Pick"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   4
      Left            =   6720
      TabIndex        =   10
      ToolTipText     =   "Work In Progress Location"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   4
      Left            =   7605
      TabIndex        =   11
      ToolTipText     =   "Check If Complete (Checked if More Than Rq'd)"
      Top             =   3120
      Width           =   420
   End
   Begin VB.TextBox txtAqt 
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   6
      ToolTipText     =   "Amount To Pick"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   3
      Left            =   6720
      TabIndex        =   7
      ToolTipText     =   "Work In Progress Location"
      Top             =   2400
      Width           =   735
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   3
      Left            =   7605
      TabIndex        =   8
      ToolTipText     =   "Check If Complete (Checked if More Than Rq'd)"
      Top             =   2400
      Width           =   420
   End
   Begin VB.TextBox txtAqt 
      Height          =   285
      Index           =   2
      Left            =   5160
      TabIndex        =   3
      ToolTipText     =   "Amount To Pick"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   2
      Left            =   6720
      TabIndex        =   4
      ToolTipText     =   "Work In Progress Location"
      Top             =   1680
      Width           =   735
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   2
      Left            =   7560
      TabIndex        =   5
      ToolTipText     =   "Check If Complete (Checked if More Than Rq'd)"
      Top             =   1680
      Width           =   420
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Index           =   1
      Left            =   6720
      TabIndex        =   1
      ToolTipText     =   "Work In Progress Location"
      Top             =   960
      Width           =   735
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   1
      Left            =   7605
      TabIndex        =   2
      ToolTipText     =   "Check If Complete (Checked if More Than Rq'd)"
      Top             =   960
      Width           =   420
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7200
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin VB.Label lblPrtCnt 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3360
      TabIndex        =   89
      Top             =   5280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   2280
      Picture         =   "PickMCe01c.frx":255A
      Top             =   5400
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   2040
      Picture         =   "PickMCe01c.frx":2A4C
      Top             =   5400
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   2520
      Picture         =   "PickMCe01c.frx":2F3E
      Top             =   5400
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   2880
      Picture         =   "PickMCe01c.frx":3430
      Top             =   5400
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Uom "
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
      Left            =   6280
      TabIndex        =   85
      Top             =   760
      Width           =   495
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   6280
      TabIndex        =   84
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   6280
      TabIndex        =   83
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   6280
      TabIndex        =   82
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   6280
      TabIndex        =   81
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   6280
      TabIndex        =   80
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   6280
      TabIndex        =   79
      Top             =   960
      Width           =   375
   End
   Begin VB.Label z3 
      BackStyle       =   0  'Transparent
      Caption         =   "Qoh"
      Height          =   255
      Index           =   6
      Left            =   3960
      TabIndex        =   78
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label z3 
      BackStyle       =   0  'Transparent
      Caption         =   "Qoh"
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   77
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label z3 
      BackStyle       =   0  'Transparent
      Caption         =   "Qoh"
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   76
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label z3 
      BackStyle       =   0  'Transparent
      Caption         =   "Qoh"
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   75
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label z3 
      BackStyle       =   0  'Transparent
      Caption         =   "Qoh"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   74
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label z3 
      BackStyle       =   0  'Transparent
      Caption         =   "Qoh"
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   73
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblQoh 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   5160
      TabIndex        =   72
      ToolTipText     =   "Quantity On Hand"
      Top             =   4880
      Width           =   1095
   End
   Begin VB.Label lblQoh 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   5160
      TabIndex        =   71
      ToolTipText     =   "Quantity On Hand"
      Top             =   4160
      Width           =   1095
   End
   Begin VB.Label lblQoh 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   5160
      TabIndex        =   70
      ToolTipText     =   "Quantity On Hand"
      Top             =   3440
      Width           =   1095
   End
   Begin VB.Label lblQoh 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   69
      ToolTipText     =   "Quantity On Hand"
      Top             =   2720
      Width           =   1095
   End
   Begin VB.Label lblQoh 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   5160
      TabIndex        =   68
      ToolTipText     =   "Quantity On Hand"
      Top             =   2000
      Width           =   1095
   End
   Begin VB.Label lblQoh 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   5160
      TabIndex        =   67
      ToolTipText     =   "Quantity On Hand"
      Top             =   1280
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "_______"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   720
      TabIndex        =   66
      Top             =   0
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   10
      Left            =   3840
      TabIndex        =   63
      Top             =   360
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   62
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblRun 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4320
      TabIndex        =   61
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblMon 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   720
      TabIndex        =   60
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      Height          =   255
      Index           =   8
      Left            =   6840
      TabIndex        =   59
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   255
      Index           =   7
      Left            =   5760
      TabIndex        =   58
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label lblLst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   57
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label lblPge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6360
      TabIndex        =   56
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label lblLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   55
      ToolTipText     =   "Inventory Location"
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   720
      TabIndex        =   54
      Top             =   4560
      Width           =   3105
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   3825
      TabIndex        =   53
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label lblPqt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   4080
      TabIndex        =   52
      ToolTipText     =   "Click To Select This Amount"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   720
      TabIndex        =   51
      Top             =   4880
      Width           =   3105
   End
   Begin VB.Label lblLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   50
      ToolTipText     =   "Inventory Location"
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   720
      TabIndex        =   49
      Top             =   3840
      Width           =   3105
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   3825
      TabIndex        =   48
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label lblPqt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   4080
      TabIndex        =   47
      ToolTipText     =   "Click To Select This Amount"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   720
      TabIndex        =   46
      Top             =   4160
      Width           =   3105
   End
   Begin VB.Label lblLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   45
      ToolTipText     =   "Inventory Location"
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   720
      TabIndex        =   44
      Top             =   3120
      Width           =   3105
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   3825
      TabIndex        =   43
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label lblPqt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   4080
      TabIndex        =   42
      ToolTipText     =   "Click To Select This Amount"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   720
      TabIndex        =   41
      Top             =   3440
      Width           =   3105
   End
   Begin VB.Label lblLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   40
      ToolTipText     =   "Inventory Location"
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   720
      TabIndex        =   39
      Top             =   2400
      Width           =   3105
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   3825
      TabIndex        =   38
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblPqt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4080
      TabIndex        =   37
      ToolTipText     =   "Click To Select This Amount"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   720
      TabIndex        =   36
      Top             =   2720
      Width           =   3105
   End
   Begin VB.Label lblLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   35
      ToolTipText     =   "Inventory Location"
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   34
      Top             =   1680
      Width           =   3105
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   3825
      TabIndex        =   33
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblPqt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   4080
      TabIndex        =   32
      ToolTipText     =   "Click To Select This Amount"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   31
      Top             =   2000
      Width           =   3105
   End
   Begin VB.Label lblLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   30
      ToolTipText     =   "Inventory Location"
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   29
      Top             =   960
      Width           =   3105
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   3825
      TabIndex        =   28
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblPqt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   27
      ToolTipText     =   "Click To Select This Amount"
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   26
      Top             =   1280
      Width           =   3105
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comp?    "
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
      Left            =   7560
      TabIndex        =   25
      Top             =   765
      Width           =   550
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Wip Loc "
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
      Left            =   6720
      TabIndex        =   24
      Top             =   760
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick Qty          "
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
      Left            =   5160
      TabIndex        =   23
      Top             =   760
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Plan Qty          "
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
      Left            =   4080
      TabIndex        =   22
      Top             =   760
      Width           =   1095
   End
   Begin VB.Label z1 
      Caption         =   "R  "
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
      Left            =   3840
      TabIndex        =   21
      Top             =   760
      Width           =   255
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
      Index           =   1
      Left            =   720
      TabIndex        =   20
      Top             =   760
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loc       "
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
      TabIndex        =   19
      Top             =   760
      Width           =   615
   End
End
Attribute VB_Name = "PickMCe01c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'More Lot Tracking 2/25/03
'5/13/04 Unit of Measure
'9/8/04 Fixed INPQTY in the pick
'9/27/04 reformated values "#####0.000"
'10/30/04 Fixed erroneous Lot request and checking
'12/7/04 Fixed Lot Picking (Parts without Lots)
'12/29/04 Added GetStdCosts and update routine to InvaTable
'3/21/07 Corrected non lottracked part LoitTable inserts
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter
Dim bType As Byte
Dim bOnLoad As Byte
Dim bUnLoad As Boolean        'set = true in load event to force unload from activate event
Dim bFIFO As Byte

Dim iCurrPage As Integer
Dim iIndex As Integer
Dim iMaxPages As Integer
Dim iTotalItems As Integer

Dim iPrtCnt As Integer
Dim iSelRunIndex As Integer

Dim sPartNumber As String
Dim sCreditAcct As String
Dim sDebitAcct As String

Dim sLots(50, 2) As String
'0 = Lot Number
'1 = Lot Quantity
'Dim sPartsGroup(250) As String   'TODO: this is redundant - should just use vitems(irow,1)

Dim vItems(250, 14) As Variant
' 0 = loc
Private Const PICK_LOCATION = 0
' 1 = part
Private Const PICK_PARTNUMBER = 1
' 2 = compressed part
Private Const PICK_PARTREF = 2
' 3 = stand cost
Private Const PICK_STANDARDCOST = 3
' 4 = desc
Private Const PICK_DESCRIPTION = 4
' 5 = rev
Private Const PICK_REVISION = 5
' 6 = planned
Private Const PICK_REQUIREDQTY = 6
' 7 = actual
Private Const PICK_QUANTITY = 7
' 8 = wip location
Private Const PICK_WIPLOCATION = 8
' 9 = complete?
Private Const PICK_ITEMCOMPLETE = 9
'10 = Qoh
Private Const PICK_QOH = 10
'11 = LotTracked Part
Private Const PICK_LOTTRACKED = 11
'12 = PKRECORD
Private Const PICK_PKRECORD = 12
'13 = Unit of Measure
Private Const PICK_UNITS = 13



'12/28/04

Private Function GetStdCosts(ListNum As Integer, Material As Currency, Labor As Currency, _
                             Expense As Currency, OverHead As Currency, hours As Currency) As Byte
   
   Dim RdoStd As ADODB.Recordset
   sSql = "SELECT PARTREF,PATOTMATL,PATOTLABOR,PATOTEXP,PATOTOH,PATOTHRS " _
          & "FROM PartTable WHERE PARTREF='" & vItems(ListNum, PICK_PARTNUMBER) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStd, ES_FORWARD)
   If bSqlRows Then
      With RdoStd
         Material = !PATOTMATL
         Labor = !PATOTLABOR
         Expense = !PATOTEXP
         OverHead = !PATOTOH
         hours = !PATOTHRS
         ClearResultSet RdoStd
      End With
   End If
   Set RdoStd = Nothing
End Function


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdDn_Click()
   iCurrPage = iCurrPage + 1
   If iCurrPage > iMaxPages Then iCurrPage = iMaxPages
   GetThisGroup
   
End Sub

Private Sub cmdDn_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5201"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdPck_Click()
   PickItems
   
End Sub

Private Sub cmdUp_Click()
   iCurrPage = iCurrPage - 1
   If iCurrPage < 1 Then iCurrPage = 1
   GetThisGroup
   
End Sub

Private Sub cmdUp_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub Form_Activate()
   If bUnLoad Then
      Unload Me
   Else
      MdiSect.lblBotPanel = Caption
      If bOnLoad Then
         Caption = Caption & " For " & PickMCe01a.cmbPrt & " Run " & PickMCe01a.cmbRun
         bFIFO = GetInventoryMethod()
         'sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yy"))
         'If Left(sJournalID, 4) = "None" Then sJournalID = ""
         bOnLoad = 0
      End If
      MouseCursor 0
   End If
End Sub

Private Sub Form_Load()
   Dim strRun As String
   FormLoad Me, True
   sCurrForm = PickMCe01a.Caption
   lblMon = Compress(PickMCe01a.cmbPrt)
   lblRun = PickMCe01a.cmbRun
   If PickMCe01a.optExp.Value Then optTyp.Value = vbChecked
   
   iPrtCnt = 0
   If (PickMCe01a.cmbRun.ListCount > 0) Then
      iPrtCnt = PickMCe01a.cmbRun.ListCount
      iSelRunIndex = PickMCe01a.cmbRun.ListIndex
      If (iSelRunIndex = -1) Then
         strRun = PickMCe01a.cmbRun
         If (strRun <> "") Then iSelRunIndex = 0 'FindTheIndex(strRun)
      End If
         
   End If
   
   If PickMCe01a.optExp.Value Then optTyp.Value = vbChecked
   
   bType = optTyp.Value
   cmdUp.Picture = Enup
   cmdUp.Enabled = True
   cmdDn.Picture = Endn
   cmdDn.Enabled = True
   sSql = "SELECT TOP 1 PARTREF,PAEXTDESC FROM PartTable WHERE " _
          & "PARTREF= ? "
  ' Set RdoQry = RdoCon.CreateQuery("", sSql)
  ' RdoQry.MaxRows = 1
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Size = 30
   AdoParameter.Type = adChar
   
   AdoQry.Parameters.Append AdoParameter

   bOnLoad = 1
   GetItems
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'PickMCe01a.FillCombo
   'PickMCe01a.optItm.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set PickMCe01c = Nothing
   
End Sub


Private Sub GetItems()
   Dim RdoPck As ADODB.Recordset
   Dim bLotsAct As Byte
   Dim iRow As Integer
   
   MouseCursor 13
   sPartNumber = Compress(lblMon)
   On Error Resume Next
   For iRow = 1 To 6
      lblLoc(iRow).Visible = False
      lblPrt(iRow).Visible = False
      lblDsc(iRow).Visible = False
      lblPqt(iRow).Visible = False
      lblRev(iRow).Visible = False
      txtAqt(iRow).Visible = False
      txtLoc(iRow).Visible = False
      lblQoh(iRow).Visible = False
      lblUom(iRow).Visible = False
      z3(iRow).Visible = False
      optPck(iRow).Visible = False
   Next
   bLotsAct = CheckLotStatus()
   iRow = 0
   Erase vItems
   'Erase sPartsGroup
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PASTDCOST,PAQOH,PALOCATION," _
          & "PALOTTRACK,PAUNITS,PKPARTREF,PKMOPART,PKMORUN,PKREV,PKTYPE,PKPQTY," _
          & "PKAQTY,PKRECORD,PKUNITS FROM PartTable,MopkTable WHERE PARTREF=PKPARTREF " _
          & "AND PKMOPART='" & sPartNumber & "' AND PKMORUN=" _
          & Trim(Val(lblRun)) & " AND PKAQTY=0 AND (PKTYPE=9 or PKTYPE=23) " _
          & "ORDER BY PALOCATION,PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPck, ES_KEYSET)
   If bSqlRows Then
      With RdoPck
         Do Until .EOF
            iRow = iRow + 1
            If iRow > 299 Then
               MouseCursor 0
               sSql = "This Pick Has More Than 300 Items." & vbCr _
                      & "You Will Have To Pick In (2) Steps."
               MsgBox sSql, vbInformation, Caption
               Exit Do
            End If
            vItems(iRow, PICK_LOCATION) = "" & Trim(!PALOCATION)
            vItems(iRow, PICK_PARTNUMBER) = "" & Trim(!PartRef)
            'spartsgroup(iRow) = "" & Trim(!PartRef)
            vItems(iRow, PICK_PARTREF) = "" & Trim(!PartNum)
            vItems(iRow, PICK_STANDARDCOST) = Format(!PASTDCOST, "#####0.000")
            vItems(iRow, PICK_DESCRIPTION) = "" & Trim(!PADESC)
            vItems(iRow, PICK_REVISION) = "" & Trim(!PKREV)
            vItems(iRow, PICK_REQUIREDQTY) = Format(!PKPQTY, "####0.000")
            If bType Then
               vItems(iRow, PICK_QUANTITY) = Format(!PKPQTY, "####0.000")
               vItems(iRow, PICK_ITEMCOMPLETE) = 1
            Else
               vItems(iRow, PICK_QUANTITY) = 0
               vItems(iRow, PICK_ITEMCOMPLETE) = 0
            End If
            vItems(iRow, PICK_WIPLOCATION) = ""
            vItems(iRow, PICK_QOH) = Format(!PAQOH, "####0.000")
            If bLotsAct = 1 Then
               vItems(iRow, PICK_LOTTRACKED) = Format(!PALOTTRACK, "0")
            Else
               vItems(iRow, PICK_LOTTRACKED) = 0
            End If
            vItems(iRow, PICK_PKRECORD) = Format(!PKRECORD)
            '5/13/04
            If Trim(!PKUNITS) = "" Then
               vItems(iRow, PICK_UNITS) = Format(!PAUNITS)
            Else
               vItems(iRow, PICK_UNITS) = Format(!PKUNITS)
            End If
            .MoveNext
         Loop
         ClearResultSet RdoPck
      End With
      iTotalItems = iRow
   End If
   If iTotalItems > 0 Then
      iCurrPage = 1
      iIndex = 0
      iMaxPages = (iTotalItems + 5) \ 6
      lblPge = "1"
   Else
      MouseCursor 0
      sSql = "UPDATE RunsTable SET RUNSTATUS='PC'" & vbCrLf _
         & "WHERE RUNREF='" & Compress(lblMon) & "' AND RUNNO=" & lblRun
      clsADOCon.ExecuteSql sSql
      MsgBox "There Are No Unpicked Items.  Status set = 'PC'", vbInformation, Caption
      PickMCe01a.bOnLoad = 1     ' force form to reload PL and PP MO's less this one
      bUnLoad = True
   End If
   lblLst = str(iMaxPages)
   If iMaxPages < 2 Then
      cmdUp.Picture = Dsup
      cmdDn.Picture = Dsdn
      cmdUp.Enabled = False
      cmdDn.Enabled = False
   End If
   cmdUp.Picture = Dsup
   cmdUp.Enabled = False
   On Error Resume Next
   For iRow = 1 To 6
      If iRow > iTotalItems Then Exit For
      lblLoc(iRow) = vItems(iRow, PICK_LOCATION)
      lblPrt(iRow) = vItems(iRow, PICK_PARTREF)
      lblDsc(iRow) = vItems(iRow, PICK_DESCRIPTION)
      lblRev(iRow) = vItems(iRow, PICK_REVISION)
      lblPqt(iRow) = Format(vItems(iRow, PICK_REQUIREDQTY), "####0.000")
      If vItems(iRow, PICK_QUANTITY) > 0 Then
         txtAqt(iRow) = Format(vItems(iRow, PICK_QUANTITY), "####0.000")
      Else
         txtAqt(iRow) = ""
      End If
      lblUom(iRow) = vItems(iRow, PICK_UNITS)
      optPck(iRow).Value = vItems(iRow, PICK_ITEMCOMPLETE)
      lblQoh(iRow) = Format(vItems(iRow, PICK_QOH), "####0.000")
      lblLoc(iRow).Visible = True
      lblPrt(iRow).Visible = True
      lblDsc(iRow).Visible = True
      lblPqt(iRow).Visible = True
      lblRev(iRow).Visible = True
      txtAqt(iRow).Visible = True
      txtLoc(iRow).Visible = True
      lblQoh(iRow).Visible = True
      lblUom(iRow).Visible = True
      z3(iRow).Visible = True
      optPck(iRow).Visible = True
      If vItems(iRow, PICK_ITEMCOMPLETE) Then
         optPck(iRow).Value = vbChecked
      Else
         optPck(iRow).Value = vbUnchecked
      End If
   Next
   On Error Resume Next
   Set RdoPck = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetThisGroup()
   Dim A As Integer
   Dim b As Integer
   Dim iRow As Integer
   
   On Error GoTo DiaErr1
   A = (iCurrPage * 6) - 5
   iIndex = A - 1
   lblPge = LTrim(str(iCurrPage))
   For iRow = 1 To 6
      lblLoc(iRow).Visible = False
      lblPrt(iRow).Visible = False
      lblDsc(iRow).Visible = False
      lblPqt(iRow).Visible = False
      lblRev(iRow).Visible = False
      txtAqt(iRow).Visible = False
      txtLoc(iRow).Visible = False
      lblQoh(iRow).Visible = False
      lblUom(iRow).Visible = False
      z3(iRow).Visible = False
      optPck(iRow).Visible = False
   Next
   If A < 0 Then A = 0
   On Error Resume Next
   For iRow = A To A + 5
      If iRow > iTotalItems Then Exit For
      b = (iRow + 1) - A
      If b < 0 Then b = 0
      If b < 7 Then
         lblLoc(b).Visible = True
         lblPrt(b).Visible = True
         lblDsc(b).Visible = True
         lblPqt(b).Visible = True
         lblRev(b).Visible = True
         txtAqt(b).Visible = True
         txtLoc(b).Visible = True
         lblQoh(b).Visible = True
         lblUom(b).Visible = True
         z3(b).Visible = True
         optPck(b).Visible = True
         lblLoc(b) = vItems(iRow, PICK_LOCATION)
         lblPrt(b) = vItems(iRow, PICK_PARTREF)
         lblRev(b) = vItems(iRow, PICK_REVISION)
         lblDsc(b) = vItems(iRow, PICK_DESCRIPTION)
         lblPqt(b) = Format(vItems(iRow, PICK_REQUIREDQTY), "####0.000")
         lblQoh(b) = Format(vItems(iRow, PICK_QOH), "####0.000")
         If Val(vItems(iRow, PICK_QUANTITY)) > 0 Then
            txtAqt(b) = Format(vItems(iRow, PICK_QUANTITY), "####0.000")
         Else
            txtAqt(b) = ""
         End If
         lblUom(b) = vItems(iRow, PICK_UNITS)
         txtLoc(b) = vItems(iRow, PICK_WIPLOCATION)
         If vItems(iRow, PICK_ITEMCOMPLETE) Then
            optPck(b).Value = vbChecked
         Else
            optPck(b).Value = vbUnchecked
         End If
      End If
   Next
   If iCurrPage = 1 Then
      cmdUp.Picture = Dsup
      cmdUp.Enabled = False
   Else
      cmdUp.Picture = Enup
      cmdUp.Enabled = True
   End If
   If iCurrPage = iMaxPages Then
      cmdDn.Picture = Dsdn
      cmdDn.Enabled = False
   Else
      cmdDn.Picture = Endn
      cmdDn.Enabled = True
   End If
   txtAqt(1).SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "getthisgr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblDsc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      If Len(Trim(lblPrt(Index))) Then GetExtendedInfo (Index)
   End If
   
End Sub


Private Sub lblLoc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      If Len(Trim(lblPrt(Index))) Then GetExtendedInfo (Index)
   End If
   
End Sub


Private Sub lblPqt_Click(Index As Integer)
   txtAqt(Index) = lblPqt(Index)
   optPck(Index).Value = vbChecked
   vItems(Index + iIndex, PICK_ITEMCOMPLETE) = optPck(Index).Value
   vItems(Index + iIndex, PICK_QUANTITY) = txtAqt(Index)
   
End Sub

Private Sub lblPrt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      If Len(Trim(lblPrt(Index))) Then GetExtendedInfo (Index)
   End If
   
End Sub


Private Sub optPck_Click(Index As Integer)
   On Error Resume Next
   If optPck(Index).Value = vbChecked Then
      If Val(txtAqt(Index)) = 0 Then
         txtAqt(Index) = lblPqt(Index)
         vItems(Index + iIndex, PICK_QUANTITY) = lblPqt(Index)
      End If
   End If
   vItems(Index + iIndex, PICK_ITEMCOMPLETE) = optPck(Index).Value
   
End Sub

Private Sub optPck_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optPck_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub optPck_LostFocus(Index As Integer)
'   If Val(vItems(Index + iIndex, PICK_REQUIREDQTY)) _
'      <= Val(vItems(Index + iIndex, PICK_QUANTITY)) Then
   If Val(vItems(Index + iIndex, PICK_REQUIREDQTY)) _
      <= Val(vItems(Index + iIndex, PICK_QUANTITY)) Then
      optPck(Index).Value = vbChecked
   End If
   If vItems(Index + iIndex, PICK_QUANTITY) = 0 Then
      optPck(Index).Value = vbUnchecked
   End If
   
End Sub

Private Sub optTyp_Click()
   'never visible. flags "all" or "none". bType true for "all"
   
End Sub

Private Sub txtAqt_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtAqt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtAqt_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtAqt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub

Private Sub txtAqt_LostFocus(Index As Integer)
   txtAqt(Index) = CheckLen(txtAqt(Index), 9)
   txtAqt(Index) = Format(Abs(Val(txtAqt(Index))), "####0.000")
   If Val(txtAqt(Index)) >= Val(lblPqt(Index)) Then
      If Val(txtAqt(Index)) <> vItems(Index + iIndex, PICK_REQUIREDQTY) Then
         optPck(Index).Value = vbChecked
      End If
   Else
      If Val(txtAqt(Index)) < vItems(Index + iIndex, PICK_QUANTITY) Then
         optPck(Index).Value = vbUnchecked
      End If
   End If
   vItems(Index + iIndex, PICK_ITEMCOMPLETE) = optPck(Index).Value
   vItems(Index + iIndex, PICK_QUANTITY) = txtAqt(Index)
   
End Sub

Private Sub txtLoc_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtLoc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtLoc_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub txtLoc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub txtLoc_LostFocus(Index As Integer)
   txtLoc(Index) = CheckLen(txtLoc(Index), 4)
   vItems(Index + iIndex, 8) = txtLoc(Index)
   
End Sub


Private Sub PickItems()
   Dim RdoPck As ADODB.Recordset
   Dim iRow As Integer
   Dim i As Integer
   Dim iLotsAvail As Integer
   Dim iPkRecord As Integer
   
   Dim bBadPick As Byte
   Dim bGoodPick As Byte
   Dim bResponse As Byte
   
   'lots
   Dim bLotsAct As Byte
   Dim bPartLot As Byte
   Dim bLotFail As Byte
   Dim bLotsFailed As Byte
   Dim bItemsPicked As Byte
   
   Dim iLots As Integer
   Dim iRef As Integer
   Dim iTrans As Integer
   
   Dim lCOUNTER As Long
   Dim lSysCount As Long
   Dim lLOTRECORD As Long
   
   Dim cCost As Currency
   Dim clineCost As Currency
   Dim lotQty As Currency
   Dim PickQty As Currency
   'Dim cItmLot As Currency
   Dim cQuantity As Currency
   Dim remainingPickQty As Currency
   
   'Costs
   Dim cMaterial As Currency
   Dim cLabor As Currency
   Dim cExpense As Currency
   Dim cOverhead As Currency
   Dim cHours As Currency
   
   Dim sLot As String
   Dim sMsg As String
   
   Dim MoPartNumber As String
   Dim moPartRef As String
   Dim moRunNo As Long
   Dim sMoRun As String * 9
   
   Dim sNewDate As String
   Dim sNewPart As String
   Dim sNewRev As String
   Dim sPickPart As String
   Dim sComment As String
   Dim vAdate As Variant
   
   MoPartNumber = Trim(lblMon)
   moPartRef = Compress(lblMon)
   moRunNo = Val(lblRun)
   'sMoPart = Compress(lblMon)
   vAdate = Format(GetServerDateTime(), "mm/dd/yy hh:mm")
   bLotsAct = CheckLotStatus()
   
   'validate information
   For iRow = 1 To iTotalItems
      If Val(vItems(iRow, PICK_QUANTITY)) > 0 Then
         bGoodPick = True
         Exit For
      End If
   Next
   If Not bGoodPick Then
      MsgBox "There Are No Selected Items to Pick.", vbInformation, Caption
      Exit Sub
   End If
   bGoodPick = 0
   bBadPick = False
   For iRow = 1 To iTotalItems
      If Val(vItems(iRow, PICK_QUANTITY)) > Val(vItems(iRow, PICK_REQUIREDQTY)) Then
         bBadPick = True
         Exit For
      End If
   Next
   If bBadPick Then
      sMsg = "One Or More Items Is Over Picked." & vbCr _
             & "Continue With The Pick."
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbNo Then Exit Sub Else bBadPick = False
   End If
   For iRow = 1 To iTotalItems
      If Val(vItems(iRow, PICK_QUANTITY)) < Val(vItems(iRow, PICK_REQUIREDQTY)) Then
         If vItems(iRow, PICK_ITEMCOMPLETE) Then
            bBadPick = True
            Exit For
         End If
      End If
   Next
   If bBadPick Then
      sMsg = "One Or More Items Is Under Picked AND " & vbCr _
             & "Is Marked Complete Continue With The Pick."
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbNo Then Exit Sub
   End If
   
   'verify that there are sufficient lot quantities for all lot-tracked items
   Dim msg As String, failures As Integer
   msg = "The following parts have insufficient lot quantities:" & vbCrLf
   failures = 0
   If bLotsAct = 1 Then
      For iRow = 1 To iTotalItems
         If Val(vItems(iRow, PICK_LOTTRACKED)) = 1 And Val(vItems(iRow, PICK_QUANTITY)) > 0 Then
            cQuantity = CCur(vItems(iRow, PICK_QUANTITY))
            'lotQty = GetRemainingLotQty(CStr(vItems(iRow, PICK_PARTNUMBER)))
            lotQty = GetLotRemainingQty(CStr(vItems(iRow, PICK_PARTNUMBER)))
            If lotQty < cQuantity Then
               failures = failures + 1
               msg = msg & "Part " & CStr(vItems(iRow, PICK_PARTNUMBER)) & " has only " & lotQty & " available.  You have requested " & cQuantity & vbCrLf
            End If
         End If
      Next
      If failures > 0 Then
         msg = msg & "Please make corrections before continuing."
         MsgBox msg
         Exit Sub
      End If
   Else
       For iRow = 1 To iTotalItems
         If Val(vItems(iRow, PICK_QUANTITY)) > 0 Then
            cQuantity = CCur(vItems(iRow, PICK_QUANTITY))
            lotQty = CCur(vItems(iRow, PICK_QOH))
            If lotQty < cQuantity Then
               failures = failures + 1
               msg = msg & "Part " & CStr(vItems(iRow, PICK_PARTNUMBER)) & " has only " & lotQty & " available.  You have requested " & cQuantity & vbCrLf
            End If
         End If
      Next
      If failures > 0 Then
         msg = msg & "Please make corrections before continuing."
         MsgBox msg
         Exit Sub
      End If
   End If
   
   'remove prior lot selections from temporary table
   sSql = "delete from TempPickLots" & vbCrLf _
      & "where ( MoPartRef = '" & moPartRef & "' and MoRunNo = " & lblRun.Caption & " )" & vbCrLf _
      & "or DateDiff( hour, WhenCreated, getdate() ) > 24"
   clsADOCon.ExecuteSql sSql
   
   'if not using lots, start the transaction here
   If bLotsAct = 0 Then
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      MouseCursor ccHourglass
   End If
   
   'Everything seems ok, so let's pick it
   'MouseCursor ccHourglass
   iRow = Len(Trim(str(lblRun)))
   iRow = 5 - iRow
   sMoRun = "RUN" & Space$(iRow) & lblRun
'   If sJournalID <> "" Then iTrans = GetNextTransaction(sJournalID)
   lCOUNTER = GetLastActivity()
   lSysCount = lCOUNTER + 1
   bItemsPicked = 0
   
   For iRow = 1 To iTotalItems
      If Val(vItems(iRow, PICK_QUANTITY)) > 0 Then
         bPartLot = Val(vItems(iRow, PICK_LOTTRACKED))
         'enough lots?
         cQuantity = Format(Val(vItems(iRow, PICK_QUANTITY)), "########0.000")
         If bPartLot = 1 And bLotsAct = 1 Then
            'lotQty = GetRemainingLotQty(CStr(vItems(iRow, PICK_PARTNUMBER)))
            lotQty = GetLotRemainingQty(CStr(vItems(iRow, PICK_PARTNUMBER)))
         Else
            lotQty = cQuantity
         End If
         
         If lotQty < cQuantity Then
            MsgBox vItems(iRow, PICK_PARTREF) & " Requires Lot Tracking And The " & vbCr _
                          & "Quantity Available Is Less Than Required.", _
                          vbInformation, Caption
            bLotFail = 1
            vItems(iRow, PICK_QUANTITY) = 0
            vItems(iRow, PICK_ITEMCOMPLETE) = 0
         Else
            
            'before beginning transaction, let user select lots
            lotQty = 0
            remainingPickQty = cQuantity
            sPickPart = vItems(iRow, PICK_PARTREF)
            Es_TotalLots = 0
            Erase lots
            If bLotsAct = 1 And bPartLot = 1 Then
               sMsg = "Part Number " & vItems(iRow, PICK_PARTREF) & " Requires lot tracking." & vbCrLf _
                      & "Do you wish to select the lots manually ( Y )," & vbCrLf _
                      & "have lots automatically assigned ( N )," & vbCrLf _
                      & "or cancel the entire pick ( Cancel )?"
               bResponse = MsgBox(sMsg, vbYesNoCancel + vbQuestion, Caption)
            Else
               bResponse = vbNo
            End If
            
            'let user pick lots
            If bResponse = vbYes Then
               'Get The lots
               LotSelect.lblPart = vItems(iRow, PICK_PARTREF)
               LotSelect.lblRequired = Abs(cQuantity)
               LotSelect.Show vbModal
               
               'if no lot selected, don't pick this item
               If Es_LotsSelected = 0 Then
                  'If sMoStatus = "PC" Then
                  '   sMoStatus = "PP"
                  'End If
                  vItems(iRow, PICK_QUANTITY) = 0
                  'RdoCon.RollbackTrans
                  bLotsFailed = bLotsFailed + 1
                  GoTo nextitem
                        
               ElseIf Es_TotalLots > 0 Then
                  For i = 1 To UBound(lots)
                     
                     'save info for lot selection
                     sSql = "INSERT INTO TempPickLots ( MoPartRef, MoRunNo, PickPartRef, LotID, LotQty,selIndex )" & vbCrLf _
                        & "Values ( '" & moPartRef & "', " & moRunNo & ", '" & vItems(iRow, PICK_PARTNUMBER) & "'," & vbCrLf _
                        & "'" & lots(i).LotSysId & "', " & lots(i).LotSelQty & ", " & iRow & " )"
                     clsADOCon.ExecuteSql sSql
                  Next
               'Else
               '   sMoStatus = "PP"
               End If
'            End If
            
            'select lots automatically
            ElseIf bResponse = vbNo Then
               iLots = GetPartLots(Compress(sPickPart))
               For i = 1 To iLots
                  lotQty = Val(sLots(i, 1))
                  If lotQty >= remainingPickQty Then
                     PickQty = remainingPickQty
                     lotQty = lotQty - remainingPickQty
                     remainingPickQty = 0
                  Else
                     PickQty = lotQty
                     remainingPickQty = remainingPickQty - lotQty
                     lotQty = 0
                  End If
                  'cItmLot = cItmLot + pickQty
                  'If cItmLot > Val(sLots(i, 1)) Then cItmLot = Val(sLots(i, 1))
                  'sLot = sLots(i, 0)
                  
                  'save info for lot selection
                  sSql = "INSERT INTO TempPickLots ( MoPartRef, MoRunNo, PickPartRef, LotID, LotQty,selIndex )" & vbCrLf _
                     & "Values ( '" & moPartRef & "', " & moRunNo & ", '" & vItems(iRow, PICK_PARTNUMBER) & "'," & vbCrLf _
                     & "'" & sLots(i, 0) & "', " & PickQty & ", " & iRow & " )"
                  clsADOCon.ExecuteSql sSql
                  If remainingPickQty <= 0 Then Exit For
               Next
               
            ElseIf bResponse = vbCancel Then
               If bLotsAct = 0 Then
                  clsADOCon.RollbackTrans
               End If
               MouseCursor 0
               bLotFail = 1
               MsgBox "Pick cancelled.", vbInformation, Caption
               Unload Me
               Exit Sub
               
               'not not tracked - just keep going
            Else
               Debug.Print "not lot tracked"
            End If
            
         End If
         
         If bLotFail <> 0 Then
            bLotsFailed = bLotsFailed + 1
         End If
      End If
nextitem:
   Next
      
   'don't quit if lot shortages have developed.  just ignore those pick items.
   'If bLotsFailed > 0 Then
   '   Exit Sub
   'End If
   
   'begin transaction if not already started
   
   If bLotsAct = 1 Then
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      MouseCursor ccHourglass
   End If
   
   'make sure lot selections are still available
   Dim rdo As ADODB.Recordset
   sSql = "select count(*) as ct" & vbCrLf _
      & "from TempPickLots tmp" & vbCrLf _
      & "join LohdTable lot on tmp.LotID = lot.LotNumber" & vbCrLf _
      & "where lot.LotRemainingQty < tmp.LotQty" & vbCrLf _
      & "and tmp.MoPartRef = '" & moPartRef & "' and tmp.MoRunNo = " & moRunNo
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      If rdo!ct > 0 Then
         If bLotsAct = 1 Then
            clsADOCon.RollbackTrans
         End If
         MsgBox "Another user has allocated quantities from the lots selected.  Please try again."
         Set rdo = Nothing
         Exit Sub
      End If
   End If
   'rdo.Close
   Set rdo = Nothing
   'pick the items
   Dim pick As New ClassPick
   Dim cUnitCost As Currency
   Dim PickRecordNo As Long
   Dim pickUnitOfMeasure As String, wipLocation As String
   Dim pickRequiredQty As Currency
   Dim pickComplete As Boolean
   
   pick.MoPartNumber = MoPartNumber
   pick.MoRunNumber = moRunNo
   
   For iRow = 1 To iTotalItems
      cQuantity = Val(vItems(iRow, PICK_QUANTITY))
      pickRequiredQty = Val(vItems(iRow, PICK_REQUIREDQTY))
      sPickPart = vItems(iRow, PICK_PARTREF)
      cUnitCost = Val(vItems(iRow, PICK_STANDARDCOST))
      PickRecordNo = CLng(vItems(iRow, PICK_PKRECORD))
      pickUnitOfMeasure = vItems(iRow, PICK_UNITS)
      wipLocation = vItems(iRow, PICK_WIPLOCATION)
      pickComplete = IIf(vItems(iRow, PICK_ITEMCOMPLETE) = 0, False, True)
      If cQuantity > 0 Then
         ' Set the Index for the Picked records
         pick.PickPartIndex = iRow
         
         If Not pick.PickPart(sPickPart, cQuantity, pickRequiredQty, pickUnitOfMeasure, cUnitCost, _
            PickRecordNo, wipLocation, pickComplete, True) Then
            
            'transaction failed
            If bLotsAct = 1 Then
               clsADOCon.RollbackTrans
            End If
            
            MouseCursor ccDefault
            MsgBox "Could Not Successfully Complete The Pick.", _
               vbExclamation, Caption
            Unload Me
            Exit Sub
         End If
            
      End If
   Next
   
'   'Open still? 5/21/03
'   sSql = "SELECT PKMOPART,PKMORUN,PKAQTY,PKTYPE FROM MopkTable WHERE (PKAQTY=0 AND PKTYPE=9 AND " _
'          & "PKMOPART='" & Compress(lblMon) & "' AND PKMORUN=" & moRunNo & ")"
'   bSqlRows = GetDataSet(RdoPck, ES_FORWARD)
'   If bSqlRows Then
'      With RdoPck
'         sMoStatus = "PP"
'         ClearResultSet RdoPck
'      End With
'   Else
'      sMoStatus = "PC"
'   End If
'   sSql = "UPDATE RunsTable SET RUNSTATUS='" & sMoStatus & "'," _
'          & "RUNCMATL=RUNCMATL+" & cCost & "," _
'          & "RUNCOST=RUNCOST+" & cCost & " " _
'          & "WHERE RUNREF='" & Compress(lblMon) & "' AND RUNNO=" & moRunNo & " "
'   RdoCon.Execute sSql, rdExecDirect

   'done
   
   MouseCursor ccDefault
   clsADOCon.CommitTrans
'   If bLotsFailed > 0 Then
'      RdoCon.RollbackTrans    'redundant?
'      UpdateWipColumns lSysCount
'      MsgBox Str(bLotsFailed) & " Lot Transaction(s) failed " & vbCrLf _
'                 & "from a quantity problem or user failure" & vbCrLf _
'                 & "to select lots and were not picked.", _
'                 vbInformation, Caption
'   ElseIf Err = 0 And bItemsPicked > 0 Then
'      RdoCon.CommitTrans      'redundant?
'      UpdateWipColumns lSysCount
'      MsgBox "Pick Completed Successfully" & vbCr _
'         & "The Manufacturing Order Status is " & sMoStatus & ".", _
'         vbInformation, Caption
'   Else
'      RdoCon.RollbackTrans    'redundant?
'      MsgBox "Could Not Successfully Complete The Pick.", _
'         vbExclamation, Caption
'   End If
   
   Set RdoPck = Nothing
   
'   'clear lot stuff
'   Erase Lots
'   Es_LotsSelected = 0
'   Es_TotalLots = 0
   If (Val(iPrtCnt) <= 1) Then
      PickMCe01a.bOnLoad = 1     ' force form to reload PL and PP MO's less this one
   Else
      ' remove the run from list
      If (iSelRunIndex <> -1) Then
         PickMCe01a.cmbRun.RemoveItem (iSelRunIndex)
         PickMCe01a.cmbRun.ListIndex = 0
      End If
      
   End If
   
   'PickMCe01a.bOnLoad = 1     ' force form to reload PL and PP MO's less this one
   Unload Me
   Exit Sub
      
DiaErr1:
   If (clsADOCon.ADOErrNum <> 0) Then
      clsADOCon.RollbackTrans
   End If
   
   sProcName = "PickItems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
   

Private Sub GetExtendedInfo(iThisindex As Integer)
   Dim RdoPrt As ADODB.Recordset
   On Error GoTo DiaErr1
   'RdoQry(0) = Compress(lblPrt(iThisindex))
   AdoQry.Parameters(0).Value = Compress(lblPrt(iThisindex))
   bSqlRows = clsADOCon.GetQuerySet(RdoPrt, AdoQry, ES_KEYSET)
   If bSqlRows Then
      With RdoPrt
         lblInfo.Top = lblPrt(iThisindex).Top + 300
         lblInfo.Visible = True
         lblInfo = "Extended Part Description: " & vbCr _
                   & Trim(!PAEXTDESC)
         lblInfo.Refresh
         Sleep 4000
         lblInfo.Visible = False
         lblInfo.Top = 0
      End With
   End If
   Set RdoPrt = Nothing
   Exit Sub
   
DiaErr1:
   On Error GoTo 0
   'Ingore Errors on this routine
   
End Sub


'3/26/02

Private Function GetPartLots(sPartWithLot As String) As Integer
   Dim RdoLots As ADODB.Recordset
   Dim iList As Integer
   
   Erase sLots
   On Error GoTo DiaErr1
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
            sLots(iList, 1) = Format$(!LOTREMAININGQTY, "#####0.000")
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

Private Function GetThisNextPickRecord() As Integer
   Dim RdoNxt As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT MAX(PKRECORD) FROM MopkTable WHERE " _
          & "PKMOPART='" & sPartNumber & "' AND PKMORUN=" & Val(lblRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNxt, ES_FORWARD)
   If bSqlRows Then
      With RdoNxt
         If Not IsNull(.Fields(0)) Then
            GetThisNextPickRecord = .Fields(0) + 1
         Else
            GetThisNextPickRecord = 1
         End If
         ClearResultSet RdoNxt
      End With
   Else
      GetThisNextPickRecord = 1
   End If
   Set RdoNxt = Nothing
   
End Function
