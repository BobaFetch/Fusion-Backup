VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form diaARf05a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Invoice Packing Slips"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   6525
   Begin VB.CheckBox chkInvPS 
      Caption         =   "___"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3720
      TabIndex        =   91
      ToolTipText     =   "Create An Advanced Payment Invoice "
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkPrintEmailedInv 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3720
      TabIndex        =   65
      Top             =   5700
      Width           =   255
   End
   Begin VB.CheckBox chkEmailInvoices 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3720
      TabIndex        =   63
      Top             =   5220
      Width           =   255
   End
   Begin VB.ComboBox cmbEmailPrefix 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2880
      TabIndex        =   87
      Tag             =   "8"
      ToolTipText     =   "Invoice Prefix(A-Z)"
      Top             =   1920
      Width           =   510
   End
   Begin VB.CheckBox chkPrintNonEmailedInv 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3720
      TabIndex        =   64
      Top             =   5460
      Width           =   255
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   6120
      TabIndex        =   67
      Top             =   5220
      Width           =   240
   End
   Begin VB.CheckBox optExt 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   6120
      TabIndex        =   68
      Top             =   5460
      Width           =   240
   End
   Begin VB.TextBox txtCopies 
      Height          =   285
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   66
      Tag             =   "1"
      Top             =   5940
      Width           =   375
   End
   Begin VB.CheckBox optIt 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   6120
      TabIndex        =   69
      Top             =   5700
      Width           =   240
   End
   Begin VB.CheckBox optLot 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   6120
      TabIndex        =   70
      Top             =   5940
      Width           =   240
   End
   Begin VB.ComboBox cmbPre 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "diaARf05a.frx":0000
      Left            =   900
      List            =   "diaARf05a.frx":0002
      TabIndex        =   71
      Tag             =   "8"
      ToolTipText     =   "Invoice Prefix(A-Z)"
      Top             =   1920
      Width           =   510
   End
   Begin ComctlLib.ProgressBar Prg1 
      Height          =   255
      Left            =   120
      TabIndex        =   76
      Top             =   6360
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ComboBox txtDte 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4320
      TabIndex        =   72
      Tag             =   "4"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Post"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4800
      TabIndex        =   73
      ToolTipText     =   "Create Invoices For These Packing Slips"
      Top             =   4920
      Width           =   915
   End
   Begin VB.CommandButton cmdDeselect 
      Caption         =   "Reselect"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5520
      TabIndex        =   74
      ToolTipText     =   "Unselect These Packing Slips"
      Top             =   1920
      Width           =   915
   End
   Begin VB.CommandButton cmdClear 
      Enabled         =   0   'False
      Height          =   330
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton cmdAll 
      Enabled         =   0   'False
      Height          =   330
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   4920
      Width           =   495
   End
   Begin VB.ComboBox txtStart 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Tag             =   "4"
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Tag             =   "4"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   25
      Left            =   6120
      TabIndex        =   31
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   24
      Left            =   5880
      TabIndex        =   30
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   23
      Left            =   5640
      TabIndex        =   29
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   22
      Left            =   5400
      TabIndex        =   28
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   21
      Left            =   5160
      TabIndex        =   27
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   20
      Left            =   4920
      TabIndex        =   26
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   19
      Left            =   4680
      TabIndex        =   25
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   18
      Left            =   4440
      TabIndex        =   24
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   17
      Left            =   4200
      TabIndex        =   23
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   16
      Left            =   3960
      TabIndex        =   22
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   15
      Left            =   3720
      TabIndex        =   21
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   14
      Left            =   3480
      TabIndex        =   20
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   13
      Left            =   3240
      TabIndex        =   19
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   12
      Left            =   3000
      TabIndex        =   18
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   11
      Left            =   2760
      TabIndex        =   17
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   10
      Left            =   2520
      TabIndex        =   16
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   9
      Left            =   2280
      TabIndex        =   15
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   8
      Left            =   2040
      TabIndex        =   14
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   7
      Left            =   1800
      TabIndex        =   13
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   6
      Left            =   1560
      TabIndex        =   12
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   5
      Left            =   1320
      TabIndex        =   11
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   4
      Left            =   1080
      TabIndex        =   10
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   3
      Left            =   840
      TabIndex        =   9
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox optLtr 
      Caption         =   "Check1"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   28
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   6255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   315
      Left            =   5520
      TabIndex        =   2
      ToolTipText     =   "Select Packing Slips With Criteria"
      Top             =   600
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2535
      Left            =   120
      TabIndex        =   59
      Top             =   2280
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      WordWrap        =   -1  'True
      Enabled         =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   78
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
      PictureUp       =   "diaARf05a.frx":0004
      PictureDn       =   "diaARf05a.frx":014A
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6060
      Top             =   6300
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6735
      FormDesignWidth =   6525
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   86
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
      PictureUp       =   "diaARf05a.frx":0290
      PictureDn       =   "diaARf05a.frx":03D6
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice same as PS"
      Height          =   285
      Index           =   20
      Left            =   2280
      TabIndex        =   92
      Top             =   4920
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Print emailed invoices"
      Height          =   285
      Index           =   13
      Left            =   120
      TabIndex        =   90
      Top             =   5760
      Width           =   3525
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Email invoices where email addresses are known"
      Height          =   285
      Index           =   12
      Left            =   120
      TabIndex        =   89
      Top             =   5280
      Width           =   3525
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Emaled Inv. Prefix"
      Height          =   315
      Index           =   6
      Left            =   1500
      TabIndex        =   88
      Top             =   1935
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Print non-emailed invoices"
      Height          =   285
      Index           =   11
      Left            =   120
      TabIndex        =   85
      Top             =   5520
      Width           =   3525
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Of Copies"
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   84
      Top             =   6000
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   9
      Left            =   4320
      TabIndex        =   83
      Top             =   5280
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   8
      Left            =   4320
      TabIndex        =   82
      Top             =   5520
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SO Item Comments"
      Height          =   285
      Index           =   7
      Left            =   4320
      TabIndex        =   81
      Top             =   5760
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Numbers"
      Height          =   285
      Index           =   5
      Left            =   4320
      TabIndex        =   80
      Top             =   6000
      Width           =   1785
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   675
      TabIndex        =   79
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inv. Prefix"
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   77
      Top             =   1935
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Date"
      Height          =   315
      Index           =   1
      Left            =   3480
      TabIndex        =   75
      Top             =   1950
      Width           =   825
   End
   Begin VB.Image imgdInc 
      Height          =   180
      Left            =   4800
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgInc 
      Height          =   180
      Left            =   5160
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include SO's With letters."
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   60
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   58
      Top             =   360
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   315
      Index           =   3
      Left            =   2280
      TabIndex        =   57
      Top             =   360
      Width           =   915
   End
   Begin VB.Label lblLtr 
      Caption         =   "Z"
      Height          =   255
      Index           =   25
      Left            =   6120
      TabIndex        =   56
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "Y"
      Height          =   255
      Index           =   24
      Left            =   5880
      TabIndex        =   55
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "X"
      Height          =   255
      Index           =   23
      Left            =   5640
      TabIndex        =   54
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "W"
      Height          =   255
      Index           =   22
      Left            =   5400
      TabIndex        =   53
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "V"
      Height          =   255
      Index           =   21
      Left            =   5160
      TabIndex        =   52
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "U"
      Height          =   255
      Index           =   20
      Left            =   4920
      TabIndex        =   51
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "T"
      Height          =   255
      Index           =   19
      Left            =   4680
      TabIndex        =   50
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "S"
      Height          =   255
      Index           =   18
      Left            =   4440
      TabIndex        =   49
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "R"
      Height          =   255
      Index           =   17
      Left            =   4200
      TabIndex        =   48
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "Q"
      Height          =   255
      Index           =   16
      Left            =   3960
      TabIndex        =   47
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "P"
      Height          =   255
      Index           =   15
      Left            =   3720
      TabIndex        =   46
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "O"
      Height          =   255
      Index           =   14
      Left            =   3480
      TabIndex        =   45
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "N"
      Height          =   255
      Index           =   13
      Left            =   3240
      TabIndex        =   44
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "M"
      Height          =   255
      Index           =   12
      Left            =   3000
      TabIndex        =   43
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "L"
      Height          =   255
      Index           =   11
      Left            =   2760
      TabIndex        =   42
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "K"
      Height          =   255
      Index           =   10
      Left            =   2520
      TabIndex        =   41
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "J"
      Height          =   255
      Index           =   9
      Left            =   2280
      TabIndex        =   40
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "I"
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   39
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "H"
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   38
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "G"
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   37
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "F"
      Height          =   255
      Index           =   5
      Left            =   1320
      TabIndex        =   36
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "E"
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   35
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "D"
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   34
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "C"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   33
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "B"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   32
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lblLtr 
      Caption         =   "A"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   180
   End
End
Attribute VB_Name = "diaARf05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'***************************************************************************************
'
' diaARf05a - Auto Invoice From Packing Slip
'
' Created: (jcw)
'
' Revisions:
'   11/25/04 (jcw) Bug Fixes.
'   02/27/04 (nth) Added LoitTable customer and invoice update.
'   03/19/04 (nth) Changed all invoice related datatypes from int to long.
'   04/21/04 (nth) Added print invoices.
'   05/12/04 (nth) Added sTemp to remember diaARp01a printer.
'   07/12/04 (nth) Fixed runtime error 380 reported by JEVINT / Cleaned up UI.
'   07/12/04 (nth) Corrected only last invouce printing.
'
'***************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bChecks() As Byte
Dim bGoodAct As Byte
Dim bGoodPs As Byte
Dim iRow As Integer
Dim lSo As Long
Dim lNewInv As Long
Dim lNextInv As Long
Dim lSalesOrder As Long
Dim iTotalItems As Integer
Dim iTotalChk As Integer
Dim sPsCust As String
Dim sPsStadr As String
Dim sPackSlip As String
Dim sAccount As String
Dim sMsg As String

Dim sTaxAccount As String
Dim sTaxState As String
Dim sTaxCode As String
Dim nTaxRate As Currency
Dim cFREIGHT As Currency

Dim cTax As Currency
Dim sType As String * 1
Dim currentCust As String

'Queries
Dim AdoQry1 As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

Dim AdoQry2 As ADODB.Command
Dim AdoParameter2 As ADODB.Parameter


' Sales journal
Dim sCOSjARAcct As String
Dim sCOSjINVAcct As String
Dim sCOSjNFRTAcct As String
Dim sCOSjTFRTAcct As String
Dim sCOSjTaxAcct As String

Dim vItems(1000, 12) As Variant
'0 = SONUMBER
'1 = SOITEM
'2 = SOREV
'3 = Part Number (compressed)
'4 = Quantity
'5 = Account
'6 = total selling
'7 = Type (level)
'8 = Product Code
'9 = Standard Cost
'10 = Number
'11 = Selling Price
'12 = Tax Exempt

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'***********************************************************************************

Private Sub cmdCan_Click()
   bCancel = True
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub CmdDeselect_Click()
   ManageBoxes False
   SetUpGrid
End Sub

Private Sub cmdGo_Click()
   
   'make sure there is a journal open for the posting date
   'CurrentJournal "SJ", txtDte.Text, sJournalID
   sJournalID = GetOpenJournal("SJ", txtDte)
   If sJournalID = "" Then
      MsgBox "No Open Sales Journal Found For " _
         & txtDte & " .", vbInformation, Caption
      txtDte.SetFocus
      Exit Sub
   End If
   
   'end date must be greater than or equal to start date
   If DateDiff("d", CDate(txtStart.Text), CDate(txtEnd.Text)) < 0 Then
      MsgBox "Ending date must be greater than or equal to start date", vbInformation, Caption
      Exit Sub
   End If
   
   'make sure the posting date is greater than or equal to the end date
   If DateDiff("d", CDate(txtEnd.Text), CDate(txtDte.Text)) < 0 Then
      MsgBox "Posting date must be greater than or equal to end date", vbInformation, Caption
      Exit Sub
   End If
   
   
   If iTotalChk > 0 Then
      CheckLoop
   Else
      MsgBox "You Must Select At Least One Packing Slip.", vbInformation, Caption
   End If
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      'find a journal for the actual posting date after it's entered
      'CurrentJournal "SJ", ES_SYSDATE, sJournalID
      CheckInvAsPS
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sCurrForm = Caption
   '    sSql = "SELECT DISTINCT PSNUMBER,PSCUST,PSTERMS,PSSTNAME,PSSTADR," _
   '        & "PSFREIGHT,CUREF,CUNICKNAME,CUNAME,PIPACKSLIP,PSPRIMARYSO FROM PshdTable," _
   '        & "CustTable,PsitTable WHERE CUREF=PSCUST AND (PSTYPE=1 AND " _
   '        & "PSSHIPPRINT=1 AND PSINVOICE=0 ) AND PSNUMBER=PIPACKSLIP AND PSNUMBER= ? "
   
   '    sSql = "SELECT DISTINCT PSNUMBER,PSCUST,PSTERMS,PSSTNAME,PSSTADR," _
   '        & "PSFREIGHT,CUREF,CUNICKNAME,CUNAME,PIPACKSLIP,PISONUMBER" & vbCrLf _
   '        & "FROM PshdTable" & vbCrLf _
   '        & "JOIN CustTable ON PshdTable.PSCUST = CustTable.CUREF" & vbCrLf _
   '        & "JOIN PsitTable ON PsitTable.PIPACKSLIP = PshdTable.PSNUMBER" & vbCrLf _
   '        & "AND (PSTYPE=1 AND PSSHIPPRINT=1 AND PSINVOICE=0 )" & vbCrLf _
   '        & "AND PSNUMBER=PIPACKSLIP AND PSNUMBER = ? "
   
   sSql = "SELECT DISTINCT PSNUMBER,PSCUST,PSTERMS,PSSTNAME,PSSTADR," _
          & "PSFREIGHT,CUREF,CUNICKNAME,CUNAME,PIPACKSLIP,PISONUMBER" & vbCrLf _
          & "FROM PshdTable" & vbCrLf _
          & "JOIN CustTable ON PshdTable.PSCUST = CustTable.CUREF" & vbCrLf _
          & "JOIN PsitTable ON PsitTable.PIPACKSLIP = PshdTable.PSNUMBER" & vbCrLf _
          & "AND (PSTYPE=1 AND PSSHIPPRINT=1 AND PSINVOICE=0 )" & vbCrLf _
          & "WHERE PSNUMBER = ? "
   
   Set AdoQry1 = New ADODB.Command
   AdoQry1.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 8
   AdoQry1.parameters.Append AdoParameter1
   
   
   
   sSql = "SELECT PIPACKSLIP,PIQTY,PIPART,PISONUMBER,PISOITEM," _
          & "PISOREV,PISELLPRICE,PARTREF,PARTNUM,PALEVEL,PAPRODCODE,PATAXEXEMPT," _
          & "PASTDCOST FROM PsitTable,PartTable WHERE PIPART=" _
          & "PARTREF AND PIPACKSLIP= ? ORDER BY PISONUMBER,PISOITEM"
   Set AdoQry2 = New ADODB.Command
   AdoQry2.CommandText = sSql
   Set AdoParameter2 = New ADODB.Parameter
   AdoParameter2.Type = adChar
   AdoParameter2.SIZE = 8
   AdoQry2.parameters.Append AdoParameter2
   
   bOnLoad = True
   txtStart = Format(Now, "mm/01/yy")
   txtEnd = Format(Now, "mm/dd/yy")
   txtDte = Format(Now, "mm/dd/yy")
   
   imgInc.Picture = Resources.imgInc.Picture
   imgdInc.Picture = Resources.imgdInc.Picture
   
   cmdAll.Picture = imgInc
   cmdClear.Picture = imgdInc
   
   cmdAll.ToolTipText = TTMARK
   cmdClear.ToolTipText = TTUNMARK
   'ShowPrinters.ToolTipText = TTPRINTER
   GetOptions
   SetUpGrid
End Sub

Private Sub SetUpGrid()
   With Grid1
      .Cols = 5
      .Rows = 1
      .ColWidth(0) = 500
      .ColWidth(1) = 1400
      .ColWidth(2) = 1400
      .ColWidth(3) = 1200
      .ColWidth(4) = 1675
      .Row = 0
      .Col = 0
      .Text = "Inc."
      .Col = 1
      .Text = "PS #"
      .Col = 2
      .Text = "SO #"
      .Col = 3
      .Text = "Cust."
      .Col = 4
      .Text = "Amount"
   End With
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   Set AdoParameter1 = Nothing
   Set AdoParameter2 = Nothing
   Set AdoQry1 = Nothing
   Set AdoQry2 = Nothing
   
   Set diaARf05a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim i As Integer
   ' 1 - 26
   For i = 0 To 25
      sOptions = sOptions & RTrim(optLtr(i).Value)
   Next
   
   ' 27
   sOptions = sOptions & Trim(cmbPre)
   
   ' 28 - 35
   sOptions = sOptions & Trim(str(chkPrintNonEmailedInv)) _
              & Trim(str(txtCopies)) _
              & Trim(str(optDsc.Value)) _
              & Trim(str(optExt.Value)) _
              & Trim(str(optIt.Value)) _
              & Trim(optLot.Value) _
              & Trim(str(chkPrintEmailedInv)) _
              & Trim(str(chkEmailInvoices))
   
   ' 36
   sOptions = sOptions & Trim(cmbEmailPrefix)
              
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim i As Integer
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      ' 1 - 26
      For i = 0 To 25
         optLtr(i).Value = Val(Mid(sOptions, i + 1, 1))
      Next
      
      '27
      If Len(Trim(sOptions)) > 26 Then
         cmbPre = Trim(Mid(sOptions, 27, 1))
      Else
         cmbPre = "I"
      End If
      
      '28 - 35
      chkPrintNonEmailedInv = Val(Mid(sOptions, 28, 1))
'      optDsc.enabled = chkPrintNonEmailedInv
'      optExt.enabled = chkPrintNonEmailedInv
'      optIt.enabled = chkPrintNonEmailedInv
'      optLot.enabled = chkPrintNonEmailedInv
'      txtCopies.enabled = chkPrintNonEmailedInv
      txtCopies = Val(Mid(sOptions, 29, 1))
      optDsc.Value = Val(Mid(sOptions, 30, 1))
      optExt.Value = Val(Mid(sOptions, 31, 1))
      optIt.Value = Val(Mid(sOptions, 32, 1))
      optLot.Value = Val(Mid(sOptions, 33, 1))
      Me.chkPrintEmailedInv.Value = Val(Mid(sOptions, 34, 1))
      Me.chkEmailInvoices.Value = Val(Mid(sOptions, 35, 1))
      
      '36
      If Len(Trim(sOptions)) >= 36 Then
         cmbEmailPrefix = Trim(Mid(sOptions, 36, 1))
      Else
         cmbEmailPrefix = "E"
      End If
   Else
      For i = 0 To 25
         optLtr(i).Value = vbChecked
      Next
      cmbPre = "I"
      cmbEmailPrefix = "E"
      txtCopies = 1
      chkPrintNonEmailedInv.Value = vbUnchecked
      chkPrintEmailedInv.Value = vbUnchecked
      chkEmailInvoices.Value = vbUnchecked
      optDsc.Value = vbUnchecked
      optExt.Value = vbUnchecked
      optIt.Value = vbUnchecked
      optLot.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub optLtr_GotFocus(Index As Integer)
   lblLtr(Index).BorderStyle = 1
End Sub

Private Sub optLtr_LostFocus(Index As Integer)
   lblLtr(Index).BorderStyle = 0
End Sub

Private Sub chkPrintNonEmailedInv_Click()
'   optDsc.enabled = chkPrintNonEmailedInv
'   optExt.enabled = chkPrintNonEmailedInv
'   optIt.enabled = chkPrintNonEmailedInv
'   optLot.enabled = chkPrintNonEmailedInv
'   txtCopies.enabled = chkPrintNonEmailedInv
End Sub



Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub txtCopies_LostFocus()
   If Val(txtCopies) > 9 Then
      txtCopies = "9"
   End If
   If Val(txtCopies) < 1 Then
      txtCopies = "1"
   End If
End Sub

Private Sub txtstart_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtstart_LostFocus()
   txtStart = CheckDate(txtStart)
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtDte = CheckDate(txtDte)
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   
   'find an open journal for the desired posting date
   CurrentJournal "SJ", txtDte.Text, sJournalID
   
End Sub

Private Sub ManageBoxes(bBool As Byte)
   Dim i As Integer
   If bBool = True Then
      cmdDeselect.enabled = True
      cmdAll.enabled = True
      cmdClear.enabled = True
      cmdGo.enabled = True
      Grid1.enabled = True
      txtDte.enabled = True
      cmbPre.enabled = True
      Me.cmbEmailPrefix.enabled = True
      txtDte.Text = Format(Now, "mm/dd/yy")
      txtStart.enabled = False
      txtEnd.enabled = False
      cmdSelect.enabled = False
      For i = 0 To 25
         optLtr(i).enabled = False
      Next
   Else
      cmdDeselect.enabled = False
      cmdAll.enabled = False
      cmdClear.enabled = False
      cmdGo.enabled = False
      Grid1.enabled = False
      txtDte.Text = ""
      txtDte.enabled = False
      cmbPre.enabled = False
      cmbEmailPrefix.enabled = False
      txtStart.enabled = True
      txtEnd.enabled = True
      cmdSelect.enabled = True
      For i = 0 To 25
         optLtr(i).enabled = True
      Next
   End If
End Sub

Private Sub cmdSelect_Click()
   GetTotalCheck
   FillGrid
   ManageBoxes True
End Sub

Private Function GetTotalCheck() As Integer
   Dim i As Integer
   Dim iCount As Integer
   GetTotalCheck = 0
   On Error Resume Next
   For i = 0 To 25
      If optLtr(i).Value = vbChecked Then
         iCount = iCount + 1
      End If
   Next
   GetTotalCheck = iCount
   
End Function

Private Sub BuildGridSQL()
   On Error GoTo DiaErr1
   Dim i As Integer
   Dim bChkflag As Byte
   Dim sSql1 As String
   
   bChkflag = 0
   '    sSql = "SELECT PshdTable.PSNUMBER, PshdTable.PSFREIGHT, PshdTable.PSPRIMARYSO," _
   '        & " SohdTable.SOTYPE, PshdTable.PSCUST, TxcdTable.TAXRATE," _
   '        & " SUM(PsitTable.PISELLPRICE * PsitTable.PIQTY) As SUBTOTAL,SohdTable.SOTAXABLE" _
   '        & " FROM  SohdTable INNER JOIN" _
   '        & " PshdTable ON SohdTable.SONUMBER = PshdTable.PSPRIMARYSO INNER JOIN" _
   '        & " PsitTable ON PshdTable.PSNUMBER = PsitTable.PIPACKSLIP INNER JOIN" _
   '        & " CustTable ON PshdTable.PSCUST = CustTable.CUREF LEFT OUTER JOIN" _
   '        & " TxcdTable ON CustTable.CUTAXCODE = TxcdTable.TAXCODE WHERE " _
   '        & "Pshdtable.PSDATE >= '" & txtStart & "' AND PshdTable.PSDATE <='" & txtEnd & "'" _
   '        & " AND PshdTable.PSPRINTED IS NOT NULL AND PSSHIPPRINT = 1 AND PSINVOICE = 0 AND PSSHIPPED=1 "
   
   sSql = "SELECT PshdTable.PSNUMBER, PshdTable.PSFREIGHT, PsitTable.PISONUMBER," & vbCrLf _
          & "SohdTable.SOTYPE, PshdTable.PSCUST, TxcdTable.TAXRATE," & vbCrLf _
          & "SUM(PsitTable.PISELLPRICE * PsitTable.PIQTY) As SUBTOTAL,SohdTable.SOTAXABLE" & vbCrLf _
          & "FROM PsitTable" & vbCrLf _
          & "INNER JOIN SohdTable ON PsitTable.PISONUMBER = SohdTable.SONUMBER" & vbCrLf _
          & "INNER JOIN PshdTable ON PsitTable.PIPACKSLIP = PshdTable.PSNUMBER" & vbCrLf _
          & "INNER JOIN CustTable ON PshdTable.PSCUST = CustTable.CUREF" & vbCrLf _
          & "LEFT OUTER JOIN TxcdTable ON CustTable.CUTAXCODE = TxcdTable.TAXCODE" & vbCrLf _
          & "WHERE Pshdtable.PSDATE >= '" & txtStart & "' AND PshdTable.PSDATE < DateAdd(day,1,'" & txtEnd & "')" & vbCrLf _
          & "AND PshdTable.PSPRINTED IS NOT NULL AND PSSHIPPRINT = 1 AND PSINVOICE = 0 AND PSSHIPPED=1 "
      
   If GetTotalCheck <> 0 And GetTotalCheck <> 26 Then
      sSql = sSql & "AND (SohdTable.SOTYPE IN ("
      For i = 0 To 25
         If optLtr(i).Value = 1 Then
            sSql = sSql & "'" & lblLtr(i) & "',"
         End If
      Next
      sSql = Left(sSql, Len(sSql) - 1)
      sSql = sSql & "))"
   End If
   '    sSql = sSql & " GROUP BY PshdTable.PSNUMBER," _
   '    & " PshdTable.PSFREIGHT,PshdTable.PSPRIMARYSO ," _
   '    & "SohdTable.SOTYPE, PshdTable.PSCUST, TxcdTable.TAXRATE, SohdTable.SOTAXABLE"
   
   
   
   sSql1 = " GROUP BY PshdTable.PSNUMBER," & vbCrLf _
          & " PshdTable.PSFREIGHT,PsitTable.PISONUMBER," & vbCrLf _
          & " SohdTable.SOTYPE, PshdTable.PSCUST, TxcdTable.TAXRATE, SohdTable.SOTAXABLE" & vbCrLf _
          & " Order by PshdTable.PSNUMBER"
   
   Debug.Print sSql1
   
   sSql = sSql & sSql1
   Debug.Print sSql
   
   Exit Sub
DiaErr1:
   sProcName = "BuildSql"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillGrid()
   Dim sItem As String
   Dim RdoLst As ADODB.Recordset
   Dim i As Integer
   ReDim bChecks(0)
   For i = 65 To 90
      AddComboStr cmbPre.hWnd, Chr$(i)
   Next
   i = 1
   On Error GoTo DiaErr1
   Grid1.Clear
   SetUpGrid
   BuildGridSQL
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      With RdoLst
         While Not .EOF
            '                sItem = "" & Chr(9) & !PSNUMBER & Chr(9) _
            '                    & !SOTYPE & Format(!PSPRIMARYSO, "00000") & Chr(9) _
            '                    & " " & Trim(!PSCUST) & Chr(9)
            sItem = "" & Chr(9) & !PsNumber & Chr(9) _
                    & !SOTYPE & Format(!PISONUMBER, SO_NUM_FORMAT) & Chr(9) _
                    & " " & Trim(!PSCUST) & Chr(9)
            If IsNull(!TAXRATE) Then
               sItem = sItem & Format((!Subtotal + !PSFREIGHT), "0.00")
            Else
               sItem = sItem & Format((!Subtotal + (!Subtotal * (!TAXRATE / 100)) + !PSFREIGHT), "0.00")
            End If
            Grid1.AddItem sItem
            If i > 0 Then
               Grid1.Row = i
               Grid1.Col = 0
               Set Grid1.CellPicture = imgdInc
               Grid1.CellPictureAlignment = flexAlignCenterCenter
            End If
            ReDim Preserve bChecks(i)
            bChecks(i) = 0
            i = i + 1
            .MoveNext
         Wend
         .Cancel
      End With
   Else
      cmdGo.enabled = False
   End If
   Set RdoLst = Nothing
   Exit Sub
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Grid1_Click()
   If iRow > 0 Then
      Grid1.Row = iRow
      If Grid1.Col = 0 And Grid1.Rows > 0 Then
         If bChecks(iRow) = 0 Then
            bChecks(iRow) = 1
            iTotalChk = iTotalChk + 1
            Set Grid1.CellPicture = imgInc
         Else
            bChecks(iRow) = 0
            iTotalChk = iTotalChk - 1
            Set Grid1.CellPicture = imgdInc
         End If
      End If
   End If
End Sub

Private Sub CmdAll_Click()
   CheckAll True
End Sub

Private Sub cmdClear_Click()
   CheckAll False
End Sub

Private Sub CheckAll(bSelect As Byte)
   Dim i As Integer
   Grid1.Col = 0
   iTotalChk = 0
   On Error Resume Next
   For i = 1 To Grid1.Rows - 1
      Grid1.Row = i
      bChecks(i) = bSelect
      If bSelect = 0 Then
         Set Grid1.CellPicture = imgdInc
      Else
         Set Grid1.CellPicture = imgInc
         iTotalChk = iTotalChk + 1
      End If
   Next
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, _
                            X As Single, Y As Single)
   iRow = Grid1.Row
End Sub

Private Sub CheckLoop()
   Dim i As Integer
   Dim A As Integer
   Dim sMsg As String
   Dim bResponse As Byte
   Dim bErrOccur As Byte
   Dim lMinInv As Long
   Dim lMaxInv As Long
   Dim sTemp As String
   
   On Error GoTo DiaErr1
   
   sMsg = "Post The Selected Invoices With Packing Slips?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      MsgBox "Transaction Cancelled", vbExclamation, Caption
      Exit Sub
   End If
   GetSJAccounts
   
   Dim solist As String
   
   A = Prg1.max \ iTotalChk
   For i = 1 To Grid1.Rows - 1
      If bChecks(i) Then
         Prg1.Visible = True
         Grid1.Row = i
         Grid1.Col = 1
         sPackSlip = Grid1.Text
         bGoodPs = GetPackslip(sPackSlip)
         Grid1.Col = 2
         lSalesOrder = CLng(Mid(Grid1.Text, 2, 6))
         Grid1.Col = 3
         sPsCust = "" & Trim(Grid1.Text)
         
         If Len(solist) > 0 Then solist = solist & ","
         solist = solist & CStr(lSalesOrder)
         
         ' Now the functions
         If bGoodPs Then
            If lMinInv = 0 Then
               lMinInv = AddInvoice(sPackSlip)
            Else
               lMaxInv = AddInvoice(sPackSlip)
            End If
            
            If (lNewInv <> 0) Then
               MasterFunction
            End If
            
            Prg1.Value = Prg1.Value + A
            If Err <> 0 Then
               bErrOccur = True
            End If
         Else
            bErrOccur = True
         End If
      End If
   Next
   If lMaxInv = 0 Then lMaxInv = lMinInv
   Prg1.Visible = False
   Prg1.Value = 0
   
   If bErrOccur = True Then
      sMsg = "Errors Occured Invoicing One or More Packing Slips"
      MsgBox sMsg, vbExclamation, Caption
   Else
     
      If lMinInv <> lMaxInv Then
         MsgBox "Invoices " & Format(lMinInv, "000000") & " Through " _
            & Format(lMaxInv, "000000") _
            & " Created From Packing Slips.", _
            vbInformation, Caption
      Else
         MsgBox "Invoice " & Format(lMinInv, "000000") & " Created.", _
            vbInformation, Caption
      End If
      iTotalChk = 0
      
      ' indicate if ITAR/EAR so's involved
      Dim rs As ADODB.Recordset
      Dim sos As String
      Dim msg As String
      sSql = "select distinct SONUMBER from SohdTable" & vbCrLf _
         & "where SOITAREAR = 1 and SONUMBER in (" & solist & ")"
         
      If clsADOCon.GetDataSet(sSql, rs, ES_FORWARD) <> 0 Then
         With rs
            Do Until rs.EOF
               If Len(sos) > 0 Then
                  sos = sos & ","
               End If
               sos = sos & !SONUMBER
               .MoveNext
            Loop
         End With
      End If
      rs.Close
      Set rs = Nothing
      If Len(sos) > 0 Then
         MsgBox "SOs with ITAR/EAR status: " & sos
      End If
      
   End If
   FillGrid
   Exit Sub
   
DiaErr1:
   sProcName = "CheckLoop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub MasterFunction()
   Dim rdoTSo As ADODB.Recordset
   Dim sSO As String
   Dim sMsg As String
   
   On Error Resume Next
   Err.Clear
   If bGoodAct = False Then
      MsgBox "One Or More Journal Accounts Are Not Registered." & vbCr _
         & "Please Install All Accounts In the Company Setup.", _
         vbInformation, Caption
   Else
      sSql = "SELECT CUTAXCODE,SOTYPE,CUNICKNAME,SOTAXABLE " _
             & "FROM CustTable INNER JOIN " _
             & "SohdTable ON CustTable.CUREF = SohdTable.SOCUST " _
             & "WHERE SONUMBER = " & lSo & " "
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoTSo)
      If Err > 0 Then
         MsgBox "This Packing Slip Cannot Be Invoiced.", _
            vbInformation, Caption
         Exit Sub
      End If
      currentCust = Compress(rdoTSo!CUNICKNAME)
      If rdoTSo!SOTAXABLE = 1 And ("" & Trim(rdoTSo!CUTAXCODE)) = "" Then
         sSO = Trim(rdoTSo!SOTYPE) & Format(lSo, "000000")
         sMsg = "Packslip " & sPackSlip & " Primary Sales Order " _
                & sSO & " Is Taxable." & vbCrLf & rdoTSo!CUNICKNAME & " Has No Tax Code Assigned. " _
                & "Packslip Cannot Be Invoiced."
         MsgBox sMsg, vbInformation, Caption
      Else
         GetPsItems
         If iTotalItems = 0 Then
            sMsg = "No Items On This Packing Slip " & sPackSlip & " To Invoice."
            MsgBox sMsg, vbInformation, Caption
         Else
            UpdateInvoice
         End If
      End If
      Set rdoTSo = Nothing
   End If
End Sub

Private Function AddInvoice(strPSNum As String) As Long
   AddInvoice = GetNextInvoice(strPSNum)
   On Error Resume Next
   Err.Clear
   lNewInv = AddInvoice
   'Use TM so that it won't show and can be safely deleted.
   sSql = "INSERT INTO CihdTable (INVNO,INVTYPE,INVSO,INVCANCELED) " _
          & "VALUES(" & lNewInv & ",'TM'," _
          & Val(sPackSlip) & ",0)"
   clsADOCon.ExecuteSql sSql
   Exit Function
   
DiaErr1:
   sProcName = "addinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Function GetPackslip(sPack As String) As Boolean
   Dim RdoPsl As ADODB.Recordset
   nTaxRate = 0
   sTaxCode = ""
   sTaxState = ""
   sTaxAccount = ""
   
   On Error GoTo DiaErr1
   
   Erase vItems
   AdoQry1.parameters(0).Value = Trim(sPack)
   bSqlRows = clsADOCon.GetQuerySet(RdoPsl, AdoQry1)
   If bSqlRows Then
      With RdoPsl
         
         sPackSlip = "" & Trim(!PsNumber)
         sPsCust = "" & Trim(!CUNICKNAME)
         sPsCust = "" & Trim(!CUREF)
         sPsStadr = "" & Trim(!PSSTNAME) & vbCrLf _
                    & Trim(!PSSTADR)
         cFREIGHT = Format(!PSFREIGHT, "#####0.00")
         'lSo = !PSPRIMARYSO
         lSo = !PISONUMBER
         .Cancel
      End With
      GetSalesTaxInfo Compress(sPsCust), nTaxRate, sTaxCode, sTaxState, sTaxAccount
      sPsStadr = CheckComments(sPsStadr)
      GetPackslip = True
   Else
      cFREIGHT = 0
      GetPackslip = False
   End If
   Set RdoPsl = Nothing
   Exit Function
DiaErr1:
   sProcName = "getpackslip"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Function GetNextInvoice(strFullPSNum As String) As Long
   
   Dim bDup As Boolean
   
   Dim inv As New ClassARInvoice
   lNextInv = inv.GetNextInvoiceNumber

   If (chkInvPS.Value <> 1) Then
      lNextInv = Format(lNextInv, "000000")
   Else
      
      Dim strPSNum As String
      strPSNum = Mid$(CStr(strFullPSNum), 3, Len(strFullPSNum))
      If (strPSNum <> "") Then
         lNextInv = Val(strPSNum)
      End If
   
   End If

   ' Validate the Invoice number
   If (Trim(lNextInv) <> "") Then
      Dim iCanceled As Integer
      
      iCanceled = 0
      bDup = inv.DuplicateInvNumber(CLng(lNextInv), iCanceled)
      
      If ((bDup = True) And (iCanceled = 0)) Then
         ' if the Inv PS is same then...get the next invoice from the invoice pool.
         If (chkInvPS.Value = 1) Then
            lNextInv = inv.GetNextInvoiceNumber
            lNextInv = Format(lNextInv, "000000")
            MsgBox "Invoice number exists for PS Number " & strFullPSNum & ".Using the New Invoice number is " & lNextInv & ".", vbInformation, Caption
         Else
            MsgBox "Invoice number exists.", vbInformation, Caption
            lNextInv = 0
         End If
      End If
   End If
   
   GetNextInvoice = lNextInv

End Function

Private Sub GetSJAccounts()
   Dim rdoJrn As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   If sJournalID = "" Then
      bGoodAct = True
      Exit Sub
   End If
   
   On Error GoTo DiaErr1
   sSql = "SELECT COREF,COSJARACCT,COSJNFRTACCT," _
          & "COSJTFRTACCT,COSJTAXACCT FROM ComnTable WHERE " _
          & "COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      With rdoJrn
         ' A/R
         sCOSjARAcct = "" & Trim(.Fields(1))
         If sCOSjARAcct = "" Then b = 1
         ' NonTaxable freight
         sCOSjNFRTAcct = "" & Trim(.Fields(2))
         If sCOSjNFRTAcct = "" Then b = 1
         ' Taxable freight
         sCOSjTFRTAcct = "" & Trim(.Fields(3))
         If sCOSjTFRTAcct = "" Then b = 1
         ' Sales tax
         sCOSjTaxAcct = "" & Trim(.Fields(4))
         If sCOSjTaxAcct = "" Then b = 1
         .Cancel
      End With
   End If
   If b = 1 Then
      bGoodAct = False
      '        lblJrn.Visible = True
   Else
      bGoodAct = True
      '        lblJrn.Visible = False
   End If
   Set rdoJrn = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getsjacco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetPsItems()
   Dim i As Integer
   Dim RdoPsi As ADODB.Recordset
   MouseCursor 13
   
   cTax = 0
   
   On Error GoTo DiaErr1
   
   ' Update the packslip dollars from sales order
   sSql = "UPDATE PsitTable SET PISELLPRICE=" _
          & "ITDOLLARS FROM PsitTable,SoitTable WHERE " _
          & "(PISONUMBER=ITSO AND PISOITEM=ITNUMBER " _
          & "AND PISOREV=ITREV) AND PIPACKSLIP='" & Trim(sPackSlip) & "'"
   clsADOCon.ExecuteSql sSql
   
   ' Return packslip items
   AdoQry2.parameters(0).Value = Trim(sPackSlip)
   bSqlRows = clsADOCon.GetQuerySet(RdoPsi, AdoQry2)
   If bSqlRows Then
      With RdoPsi
         ' Load invoice items into vItems array for use in
         ' IpdateInvoice subroutine
         Do Until .EOF
            i = i + 1
            vItems(i, 0) = Format(!PISONUMBER, SO_NUM_FORMAT)
            vItems(i, 1) = Format(!PISOITEM, "##0")
            vItems(i, 2) = Trim(!PISOREV)
            vItems(i, 3) = "" & Trim(!PartRef)
            'vItems(i, 4) = Format(!PIQTY, "#####0.000")
            vItems(i, 4) = !PIQTY
            vItems(i, 5) = "" & sAccount
            vItems(i, 6) = "" & !PIPART
            vItems(i, 7) = "" & !PALEVEL
            vItems(i, 8) = "" & Trim(!PAPRODCODE)
            'vItems(i, 9) = Format(!PASTDCOST, "#####0.000")
            vItems(i, 9) = !PASTDCOST
            vItems(i, 10) = 1
            'vItems(i, 11) = Format(!PISELLPRICE, "#####0.000")
            vItems(i, 11) = !PISELLPRICE
            vItems(i, 12) = !PATAXEXEMPT
            If vItems(i, 12) = 0 Then
               cTax = cTax + ((vItems(i, 11) * vItems(i, 4)) * (nTaxRate / 100))
            End If
            .MoveNext
         Loop
      End With
   End If
   iTotalItems = i
   Set RdoPsi = Nothing
   MouseCursor 0
   Exit Sub
   ' Error handeling
DiaErr1:
   sProcName = "getpsitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub UpdateInvoice()
   Dim bByte As Byte
   Dim bResponse As Byte
   Dim i As Integer
   Dim iTrans As Integer
   Dim iRef As Integer
   Dim iLevel As Integer
   Dim cCost As Currency
   Dim nLDollars As Single
   Dim nTdollars As Single
   Dim sPart As String
   Dim sProd As String
   Dim sMsg As String
   ' Accounts
   Dim sPartRevAcct As String
   Dim sPartCgsAcct As String
   Dim sRevAccount As String
   Dim sDisAccount As String
   Dim sCGSMaterialAccount As String
   Dim sCGSLaborAccount As String
   Dim sCGSExpAccount As String
   Dim sCGSOhAccount As String
   Dim sInvMaterialAccount As String
   Dim sInvLaborAccount As String
   Dim sInvExpAccount As String
   Dim sInvOhAccount As String
   Dim RdoTax As ADODB.Recordset
   
   ' BnO Taxes
   Dim nRate As Single
   Dim sType As String
   Dim sState As String
   Dim sCode As String
   Dim sPost As String
   Dim sTemp As String
   
   On Error GoTo DiaErr1
   
   ' Look For Accounts ?
   If sJournalID <> "" Then
      iTrans = GetNextTransaction(sJournalID)
   End If
   If iTrans > 0 Then
      bByte = True
      For i = 1 To iTotalItems
         If Val(vItems(i, 8)) > 0 Then
            sPart = vItems(i, 3)
            iLevel = Val(vItems(i, 7))
            sProd = vItems(i, 8)
            bByte = GetPartInvoiceAccounts(sPart, iLevel, sProd, _
                    sPartRevAcct, , sPartCgsAcct)
         End If
         If bByte = False Then Exit For
      Next
   End If
   
   lCurrInvoice = lNextInv
   sPost = Format(txtDte, "mm/dd/yyyy")
   
   'determine whether an email invoice is to be generated
   Dim emailInvoice As Boolean
   Dim invoicePrefix As String
   emailInvoice = False
   invoicePrefix = Me.cmbPre
   If Me.chkEmailInvoices.Value = vbChecked Then
      ' see if this report is available
      Dim emailer As KeyMailer
      Set emailer = New KeyMailer
      emailer.ReportName = "CustInvc"
      If Not emailer.GetReportInfo(True) Then
         Set emailer = Nothing
         Exit Sub
      End If
      emailer.ReportName = "CustInvc"
      emailer.DistributionListKey = currentCust
      emailer.AddLongParameter "INVNO", lCurrInvoice
      emailer.AddStringParameter "INVCUST", currentCust
      emailer.AddBooleanParameter "ShowDescription", True
      If emailer.IsRequestValid = VAL_ValidWithDistList Then
         emailInvoice = True
         invoicePrefix = Me.cmbEmailPrefix
      Else
         Set emailer = Nothing
      End If
   End If
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "UPDATE PshdTable SET PSINVOICE=" & lCurrInvoice & " " _
          & "WHERE PSNUMBER='" & sPackSlip & "' "
   clsADOCon.ExecuteSql sSql
   
   ' Update lot record
   sSql = "UPDATE LoitTable SET LOICUSTINVNO=" & lCurrInvoice _
          & ",LOICUST='" & sPsCust & "' WHERE LOIPSNUMBER = '" & sPackSlip & "'"
   clsADOCon.ExecuteSql sSql
   
   For i = 1 To iTotalItems
      ' Running invoice total
      nTdollars = nTdollars + (Val(vItems(i, 4)) * Val(vItems(i, 11)))
      ' Part accounts
      sPart = vItems(i, 3)
      sProd = vItems(i, 8)
      iLevel = Val(vItems(i, 7))
      bByte = GetPartInvoiceAccounts(sPart, iLevel, sProd, _
              sRevAccount, sDisAccount, sCGSMaterialAccount, sCGSLaborAccount, _
              sCGSExpAccount, sCGSOhAccount, sInvMaterialAccount, sInvLaborAccount, _
              sInvExpAccount, sInvOhAccount)
      
      ' BnO tax
      sCode = ""
      nRate = 0
      sState = ""
      sType = ""
      
      'new functions (12/16/10)
      GetPartBnO vItems(i, 3), nRate, sCode, sState, sType
      If sCode = "" Then
         GetCustBnO sPsCust, nRate, sCode, sState, sType
      End If
      
      
      ' Update the sales order items (revised 12/16/03)
      sSql = "UPDATE SoitTable SET " _
             & "ITINVOICE=" & lCurrInvoice & "," _
             & "ITREVACCT='" & sRevAccount & "'," _
             & "ITCGSACCT='" & sCOSjARAcct & "'," _
             & "ITBOSTATE='" & sState & "'," _
             & "ITBOCODE='" & sCode & "'," _
             & "ITSLSTXACCT='" & sTaxAccount & "'," _
             & "ITTAXCODE='" & sTaxCode & "'," _
             & "ITSTATE='" & sTaxState & "'," _
             & "ITTAXRATE=" & nTaxRate & "," _
             & "ITTAXAMT=" & CCur((nTaxRate / 100) * (Val(vItems(i, 4)) * Val(vItems(i, 11)))) & " " _
             & "WHERE ITSO=" & Val(vItems(i, 0)) & " AND " _
             & "ITNUMBER=" & Val(vItems(i, 1)) & " AND " _
             & "ITREV='" & vItems(i, 2) & "' "
      clsADOCon.ExecuteSql sSql
      
      'Journal entries
      nLDollars = (Val(vItems(i, 4)) * Val(vItems(i, 11)))
      cCost = (Val(vItems(i, 4)) * Val(vItems(i, 9)))
      
      ' Debit A/R (+)
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT," _
             & "DCACCTNO,DCPARTNO,DCSONUMBER,DCSOITNUMBER,DCSOITREV," _
             & "DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & CCur(nLDollars) & ",'" _
             & sCOSjARAcct & "','" _
             & sPart & "'," _
             & vItems(i, 0) & "," _
             & vItems(i, 1) & ",'" _
             & vItems(i, 2) & "','" _
             & Trim(sPsCust) & "','" _
             & Format(txtDte, "mm/dd/yy") & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
      
      ' Credit Revenue (-)
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCCREDIT," _
             & "DCACCTNO,DCPARTNO,DCSONUMBER,DCSOITNUMBER,DCSOITREV," _
             & "DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & CCur(nLDollars) & ",'" _
             & sRevAccount & "','" _
             & sPart & "'," _
             & vItems(i, 0) & "," _
             & vItems(i, 1) & ",'" _
             & vItems(i, 2) & "','" _
             & Trim(sPsCust) & "','" _
             & Format(txtDte, "mm/dd/yy") & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
   Next
   
   ' Tax and freight
   
   If cFREIGHT > 0 Then
      
      ' Debit A/R Freight
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT," _
             & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & cFREIGHT & ",'" _
             & sCOSjARAcct & "','" _
             & Trim(sPsCust) & "','" _
             & Format(txtDte, "mm/dd/yy") & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
      
      ' Credit Freight
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCCREDIT," _
             & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & cFREIGHT & ",'" _
             & sCOSjNFRTAcct & "','" _
             & Trim(sPsCust) & "','" _
             & Format(txtDte, "mm/dd/yy") & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
   End If
   
   'cTax = CCur(txtTax)
   If cTax > 0 Then
      ' Debit A/R Taxes
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT," _
             & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & cTax & ",'" _
             & sCOSjARAcct & "','" _
             & Trim(sPsCust) & "','" _
             & Format(txtDte, "mm/dd/yy") & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
      
      ' Credit Taxes
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCCREDIT," _
             & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & cTax & ",'" _
             & sCOSjTaxAcct & "','" _
             & Trim(sPsCust) & "','" _
             & Format(txtDte, "mm/dd/yy") & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
   End If
   
   ' Change TM invoice to PS
   sSql = "UPDATE CihdTable SET INVNO=" & lCurrInvoice & "," _
          & "INVPRE='" & invoicePrefix & "',INVSTADR='" & sPsStadr & "'," _
          & "INVTYPE='PS',INVSO=0," _
          & "INVCUST='" & Trim(sPsCust) & "' WHERE " _
          & "INVNO=" & lCurrInvoice & " AND INVTYPE='TM'"
          '& "INVNO=" & lNewInv & " AND INVTYPE='TM'"
   clsADOCon.ExecuteSql sSql
   
   If (chkInvPS.Value <> 1) Then
      Dim inv As New ClassARInvoice
      inv.SaveLastInvoiceNumber lCurrInvoice
   End If
   
   ' Add freight and tax to invoice total
   nTdollars = nTdollars + (cTax + cFREIGHT)
   
   'MM added CANCELED flag to 0
   ' Then post the total to the invoice
   sSql = "UPDATE CihdTable SET INVTOTAL=" & nTdollars & "," _
          & "INVFREIGHT=" & cFREIGHT & "," _
          & "INVTAX=" & cTax & "," _
          & "INVSHIPDATE='" & sPost & "'," _
          & "INVDATE='" & sPost & "'," _
          & "INVCANCELED=0," _
          & "INVCOMMENTS=''," _
          & "INVPACKSLIP='" & sPackSlip & "' " _
          & "WHERE INVNO=" & lCurrInvoice & " "
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE PshdTable SET PSFREIGHT=" & cFREIGHT _
          & " WHERE PSNUMBER='" & sPackSlip & "' "
   clsADOCon.ExecuteSql sSql
   
   MouseCursor 0
   
   If (clsADOCon.ADOErrNum = 0) Then
      clsADOCon.CommitTrans
         
      If emailInvoice Then
         If Not emailer.Generate Then
            MsgBox "Unable to queue invoice"
            'emailInvoice = False
         End If
      End If
      Set emailer = Nothing
   
      'print invoice if required
      If (emailInvoice And chkPrintEmailedInv.Value = vbChecked) _
         Or ((Not emailInvoice) And chkPrintNonEmailedInv.Value = vbChecked) Then
         PrintInvoices lCurrInvoice, lCurrInvoice
      End If
      
   End If
   
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   clsADOCon.ADOErrNum = 0
   sProcName = "updateinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintInvoices(MinInvoiceNumber As Long, MaxInvoiceNumber As Long)
   Dim sTemp As String
   diaARp01a.WindowState = vbMinimized
   diaARp01a.bRemote = True
   diaARp01a.cmbInv = MinInvoiceNumber
   diaARp01a.cmbInv2 = MaxInvoiceNumber
   diaARp01a.optDsc = optDsc
   diaARp01a.optExt = optExt
   diaARp01a.optIt = optIt
   diaARp01a.optLot = optLot
   diaARp01a.txtCopies = txtCopies
   sTemp = diaARp01a.lblPrinter
   diaARp01a.lblPrinter = lblPrinter
   diaARp01a.optPrn = chkPrintNonEmailedInv
   DoEvents
   diaARp01a.lblPrinter = sTemp
   Unload diaARp01a
End Sub

Private Function CheckInvAsPS()
   Dim RdoInv As ADODB.Recordset
      
   sSql = "SELECT * FROM ComnTable WHERE COALLOWINVNUMPS = 1"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_KEYSET)
   If bSqlRows Then
      'chkInvPS.enabled = True
      chkInvPS.Value = 1
      ClearResultSet RdoInv
   Else
      'chkInvPS.enabled = False
      chkInvPS.Value = 0
   End If
   Set RdoInv = Nothing
   
End Function


