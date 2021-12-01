VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form MrplMRp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MRP Requirements By Part(s)"
   ClientHeight    =   5880
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5880
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "MrplMRp01a.frx":0000
      Height          =   315
      Left            =   5160
      Picture         =   "MrplMRp01a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   960
      Width           =   350
   End
   Begin VB.TextBox txtPrt 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "3"
      Top             =   960
      Width           =   3495
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   25
      Left            =   4560
      TabIndex        =   74
      Top             =   4320
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   24
      Left            =   4320
      TabIndex        =   73
      Top             =   4320
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   23
      Left            =   4080
      TabIndex        =   72
      Top             =   4320
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   22
      Left            =   3840
      TabIndex        =   71
      Top             =   4320
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   21
      Left            =   3600
      TabIndex        =   70
      Top             =   4320
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   20
      Left            =   3360
      TabIndex        =   69
      Top             =   4320
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   19
      Left            =   3120
      TabIndex        =   68
      Top             =   4320
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   18
      Left            =   2880
      TabIndex        =   67
      Top             =   4320
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   17
      Left            =   2640
      TabIndex        =   66
      Top             =   4320
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   16
      Left            =   2400
      TabIndex        =   65
      Top             =   4320
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   15
      Left            =   2160
      TabIndex        =   64
      Top             =   4320
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   14
      Left            =   1920
      TabIndex        =   63
      Top             =   4320
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   13
      Left            =   1680
      TabIndex        =   62
      Top             =   4320
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   12
      Left            =   4560
      TabIndex        =   61
      Top             =   3840
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   11
      Left            =   4320
      TabIndex        =   60
      Top             =   3840
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   10
      Left            =   4080
      TabIndex        =   59
      Top             =   3840
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   9
      Left            =   3840
      TabIndex        =   58
      Top             =   3840
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   8
      Left            =   3600
      TabIndex        =   57
      Top             =   3840
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   7
      Left            =   3360
      TabIndex        =   56
      Top             =   3840
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   6
      Left            =   3120
      TabIndex        =   55
      Top             =   3840
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   2880
      TabIndex        =   54
      Top             =   3840
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   2640
      TabIndex        =   53
      Top             =   3840
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   2400
      TabIndex        =   52
      Top             =   3840
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   51
      Top             =   3840
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   50
      Top             =   3840
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox optTyp 
      BackColor       =   &H00000000&
      Caption         =   "__"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   0
      Left            =   1680
      TabIndex        =   49
      Top             =   3840
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.Frame z2 
      Height          =   495
      Left            =   1680
      TabIndex        =   42
      Top             =   2400
      Width           =   2775
      Begin VB.OptionButton optMbe 
         Caption         =   "M"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "B"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   45
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "E"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   44
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "ALL"
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   43
         Top             =   200
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CheckBox chkType 
      Caption         =   "8"
      Height          =   255
      Index           =   8
      Left            =   5880
      TabIndex        =   41
      Top             =   3120
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "7"
      Height          =   255
      Index           =   7
      Left            =   5280
      TabIndex        =   40
      Top             =   3120
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   39
      Top             =   3120
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   38
      Top             =   3120
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   37
      Top             =   3120
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   36
      Top             =   3120
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   35
      Top             =   3120
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   34
      Top             =   3120
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.ComboBox cmbPart 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.CheckBox ChkTimeErr 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "MrplMRp01a.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Show Printers"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "MrplMRp01a.frx":080E
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4080
      TabIndex        =   6
      Tag             =   "4"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Tag             =   "4"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   4080
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   1440
      Width           =   855
   End
   Begin VB.CheckBox chkExtDesc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   4680
      Width           =   735
   End
   Begin VB.CheckBox chkExceptions 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   5040
      Width           =   735
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "MrplMRp01a.frx":0FBC
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "MrplMRp01a.frx":113A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5880
      FormDesignWidth =   7455
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      Height          =   255
      Index           =   25
      Left            =   4560
      TabIndex        =   101
      Top             =   4080
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      Height          =   255
      Index           =   24
      Left            =   4320
      TabIndex        =   100
      Top             =   4080
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Index           =   23
      Left            =   4080
      TabIndex        =   99
      Top             =   4080
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      Height          =   255
      Index           =   22
      Left            =   3840
      TabIndex        =   98
      Top             =   4080
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      Height          =   255
      Index           =   21
      Left            =   3600
      TabIndex        =   97
      Top             =   4080
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      Height          =   255
      Index           =   20
      Left            =   3360
      TabIndex        =   96
      Top             =   4080
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      Height          =   255
      Index           =   19
      Left            =   3120
      TabIndex        =   95
      Top             =   4080
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      Height          =   255
      Index           =   18
      Left            =   2880
      TabIndex        =   94
      Top             =   4080
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      Height          =   255
      Index           =   17
      Left            =   2640
      TabIndex        =   93
      Top             =   4080
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      Height          =   255
      Index           =   16
      Left            =   2400
      TabIndex        =   92
      Top             =   4080
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      Height          =   255
      Index           =   15
      Left            =   2160
      TabIndex        =   91
      Top             =   4080
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      Height          =   255
      Index           =   14
      Left            =   1920
      TabIndex        =   90
      Top             =   4080
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      Height          =   255
      Index           =   13
      Left            =   1680
      TabIndex        =   89
      Top             =   4080
      Width           =   165
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Types"
      Height          =   255
      Index           =   15
      Left            =   240
      TabIndex        =   88
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      Height          =   255
      Index           =   12
      Left            =   4560
      TabIndex        =   87
      Top             =   3600
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      Height          =   255
      Index           =   11
      Left            =   4320
      TabIndex        =   86
      Top             =   3600
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      Height          =   255
      Index           =   10
      Left            =   4080
      TabIndex        =   85
      Top             =   3600
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      Height          =   255
      Index           =   9
      Left            =   3840
      TabIndex        =   84
      Top             =   3600
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   83
      Top             =   3600
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   82
      Top             =   3600
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   255
      Index           =   6
      Left            =   3120
      TabIndex        =   81
      Top             =   3600
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   80
      Top             =   3600
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   79
      Top             =   3600
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   78
      Top             =   3600
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   77
      Top             =   3600
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   76
      Top             =   3600
      Width           =   165
   End
   Begin VB.Label lblAlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   75
      Top             =   3600
      Width           =   165
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make, Buy, Either"
      Height          =   285
      Index           =   14
      Left            =   240
      TabIndex        =   48
      Top             =   2640
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Types"
      Height          =   285
      Index           =   13
      Left            =   240
      TabIndex        =   47
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Only Parts With Time Error"
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   33
      Top             =   5400
      Width           =   2625
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   10
      Left            =   5640
      TabIndex        =   30
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   9
      Left            =   5640
      TabIndex        =   29
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   11
      Left            =   5640
      TabIndex        =   28
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   27
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   26
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Classes"
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   25
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Codes"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   23
      Top             =   2280
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Only Parts With Exceptions"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   22
      Top             =   5040
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   21
      Top             =   4680
      Width           =   2385
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last MRP"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   20
      Top             =   360
      Width           =   975
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   19
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblMrp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblUsr 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   360
      Width           =   615
   End
   Begin VB.Label p 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   1005
      Width           =   1425
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   14
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "MrplMRp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/19/06 Revised report and selections. Removed extra report.
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd




Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If cmbCde = "" Then cmbCde = "ALL"
   
End Sub


Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 6)
   If cmbCls = "" Then cmbCls = "ALL"
   
End Sub



Private Sub cmbPart_LostFocus()
    cmbPart = CheckLen(cmbPart, 30)
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub

Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPart = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPart
   ViewParts.Show
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"
   cmbPart = txtPrt
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

'Private Sub cmdFnd_Click()
'   ViewParts.lblControl = "TXTPRT"
'   ViewParts.txtPrt = txtPrt
'   optVew.Value = vbChecked
'   ViewParts.Show
'
'End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub FillCombos()
    On Error Resume Next
    sSql = "SELECT DISTINCT PARTREF,PARTNUM " _
        & "FROM PartTable  " _
        & "INNER JOIN MrplTable ON MrplTable.MRP_PARTREF=PartTable.PARTREF " _
        & " WHERE PAINACTIVE = 0 AND PAOBSOLETE = 0 " _
        & "ORDER BY PARTREF"
    LoadComboBox cmbPart, 0
    cmbPart = "ALL"
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetMRPDates
      GetLastMrp
      cmbCde.AddItem "ALL"
      FillProductCodes
      If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
      cmbCls.AddItem "ALL"
      FillProductClasses
      If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
      
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillCombos
      
      bOnLoad = 0
   
   End If
   If optVew.Value = vbChecked Then
      optVew.Value = vbUnchecked
      Unload ViewParts
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set MrplMRp01a = Nothing
   
End Sub




Private Sub PrintReport()
    Dim A As Byte
    Dim b As Byte
    Dim C As Byte
    Dim sParts As String
    Dim sCode As String
    Dim sClass As String
    Dim sBDate As String
    Dim sEDate As String
    Dim sBegDate As String
    Dim sEndDate As String
    Dim sMbe As String
    
    Dim sCustomReport As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    Dim strIncludes As String
    Dim strDateDev As String
    Dim sSubSql As String
    

    MouseCursor 13
    On Error GoTo DiaErr1
    GetMRPCreateDates sBegDate, sEndDate

    For b = 0 To 25
       If optTyp(b).Value = vbChecked Then C = 1
    Next
    If C = 0 Then
      MsgBox "No SO Types Selected.", vbInformation, Caption
      Exit Sub
    End If
    

    If Trim(txtBeg) = "" Then txtBeg = "ALL"
    If Trim(txtEnd) = "" Then txtEnd = "ALL"
    If Not IsDate(txtBeg) Then
       sBDate = "1995,01,01"
    Else
       sBDate = Format(txtBeg, "yyyy,mm,dd")
    End If
    If Not IsDate(txtEnd) Then
       sEDate = "2024,12,31"
    Else
       sEDate = Format(txtEnd, "yyyy,mm,dd")
    End If

'    If Trim(txtPrt) = "" Then txtPrt = "ALL"
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
    
    If Trim(cmbCde) = "" Then cmbCde = "ALL"
    If Trim(cmbCls) = "" Then cmbCls = "ALL"
    If Trim(cmbPart) = "ALL" Then sParts = "" Else sParts = Compress(cmbPart)
    If Trim(cmbCde) = "ALL" Then sCode = "" Else sCode = Compress(cmbCde)
    If Trim(cmbCls) = "ALL" Then sClass = "" Else sClass = Compress(cmbCls)



    'get custom report name if one has been defined
    sCustomReport = GetCustomReport("prdmr01")

    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.ShowGroupTree False

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ExceptionsOnly"
    aFormulaName.Add "ShowOnlyTimeErr"
    aFormulaName.Add "DateDeveloped"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")

    strIncludes = Trim(cmbPart) & ", Prod Code(s) " & cmbCde & ", Class(es) " _
                            & cmbCls
    aFormulaValue.Add CStr("'" & CStr(strIncludes) & "...'")
    aFormulaValue.Add CStr("'" & CStr(sInitials) & "'")

    aFormulaValue.Add CStr("'" & CStr(chkExceptions.Value) & "'")
    'aFormulaValue.Add CStr("'1'")
    aFormulaValue.Add ChkTimeErr.Value
    strDateDev = "'MRP Created  " & sBegDate & " For Requirements Through " & sEndDate & "'"
    aFormulaValue.Add CStr(strDateDev)

    sSql = "{MrplTable.MRP_PARTREF} LIKE '" & sParts & "*' " _
          & "AND {MrplTable.MRP_PARTPRODCODE} LIKE '" & sCode _
          & "*' AND {MrplTable.MRP_PARTCLASS} LIKE '" & sClass & "*' " _
          & "AND {MrplTable.MRP_PARTDATERQD} In Date(" & sBDate & ") " _
          & " To Date(" & sEDate & ") AND " _
        & " ({MrplTable.MRP_TYPE} <> 6 AND " _
        & " {MrplTable.MRP_TYPE} <> 5)"

'          & " ({MrplTable.MRP_PARTLEVEL} <> 6 AND " _
'        & " {MrplTable.MRP_PARTLEVEL} <> 5) AND " _

    sSubSql = "{MrplTable.MRP_PARTREF} = {?Pm-PartTable.PARTREF} and " _
                & "{MrplTable.MRP_PARTDATERQD} in Date('" & sBDate & "') to Date('" & sEDate & "')  AND " _
        & " ({MrplTable.MRP_TYPE} <> 6 AND " _
        & " {MrplTable.MRP_TYPE} <> 5)"

'          & " ({MrplTable.MRP_PARTLEVEL} <> 6 AND " _
'        & " {MrplTable.MRP_PARTLEVEL} <> 5) AND " _

   If optMbe(0).Value = True Then
      sMbe = "Make"
      sSql = sSql & "AND {PartTable.PAMAKEBUY}='M'"
      sSubSql = sSubSql & "AND {PartTable.PAMAKEBUY}='M'"
   ElseIf optMbe(1).Value = True Then
      sMbe = "Buy"
      sSql = sSql & "AND {PartTable.PAMAKEBUY}='B'"
      sSubSql = sSubSql & "AND {PartTable.PAMAKEBUY}='B'"
   ElseIf optMbe(2).Value = True Then
      sMbe = "Either"
      sSql = sSql & "AND {PartTable.PAMAKEBUY}='E'"
      sSubSql = sSubSql & "AND {PartTable.PAMAKEBUY}='E'"
   Else
      sMbe = "Make, Buy And Either"
   End If
   
   aFormulaName.Add "Mbe"
   aFormulaValue.Add CStr("'" & sMbe & "'")
    
    
   'select part types
   Dim types As String
   Dim includes As String
   Dim i As Integer
   For i = 1 To 8
     If Me.chkType(i).Value = vbChecked Then
        If types = "" Then
           types = " AND ("
        Else
           types = types & " OR "
        End If
        
        types = types & "{PartTable.PALEVEL} = " & i
        includes = includes & " " & i
     End If
   Next
   
   If types = "" Then
     MsgBox "No part types selected"
     Exit Sub
   Else
     sSql = sSql & types & ")"
     sSubSql = sSubSql & types & ")"
   End If

   A = 65
   C = 0
   For b = 0 To 25
      If optTyp(b).Value = vbChecked Then
         If C = 1 Then
            sSql = sSql & "OR {MrplTable.MRP_SOTEXT} LIKE '" & Chr$(A) & "*' "
         Else
            sSql = sSql & " AND ({MrplTable.MRP_SOTEXT} = '' OR {MrplTable.MRP_SOTEXT} LIKE '" & Chr$(A) & "*' "
         End If
         C = 1
      End If
      A = A + 1
   Next
   sSql = sSql & ")"

   aFormulaName.Add "PartInc"
   aFormulaValue.Add CStr("'" & includes & "'")
    
   aFormulaName.Add "ShowExDesc"
   aFormulaValue.Add chkExtDesc.Value
        
    ' Set Formula values
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
    ' set the report Selection
    cCRViewer.SetReportSelectionFormula sSql
    ' set the sub sql variable pass the sub report name
    cCRViewer.SetSubRptSelFormula "prdmr01_SR", sSubSql

    ' report parameter
    'If chkExtDesc.Value = vbUnchecked Then
    '    cCRViewer.SetReportSection "GroupHeaderSection3", True
    'Else
    '    cCRViewer.SetReportSection "GroupHeaderSection3", False
    'End If

    cCRViewer.CRViewerSize Me
    
    ' Set report parameter
    cCRViewer.SetDbTableConnection


    cCRViewer.OpenCrystalReportObject Me, aFormulaName

    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aRptParaType
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue


   MouseCursor 0
   Exit Sub

DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
'   txtPrt = "ALL"
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sCode As String * 6
   Dim sClass As String * 4
   sCode = cmbCde
   sClass = cmbCls
   sOptions = sCode & sClass & Trim(str(Val(chkExtDesc.Value))) _
              & Trim(str(Val(chkExceptions.Value)))
   SaveSetting "Esi2000", "EsiProd", "Prdmr01", sOptions
   SaveSetting "Esi2000", "EsiProd", "Pmr01", lblPrinter
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "Prdmr01", sOptions)
   If Len(Trim(sOptions)) > 0 Then
      cmbCde = Mid$(sOptions, 1, 6)
      cmbCls = Mid$(sOptions, 7, 4)
      chkExtDesc.Value = Val(Mid$(sOptions, 11, 1))
      chkExceptions.Value = Val(Mid$(sOptions, 12, 1))
   End If
   lblPrinter = GetSetting("Esi2000", "EsiProd", "Pmr01", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub chkExtDesc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub chkExceptions_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
End Sub

Private Sub txtend_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
End Sub

'Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF4 Then
'      ViewParts.lblControl = "TXTPRT"
'      ViewParts.txtPrt = txtPrt
'      optVew.Value = vbChecked
'      ViewParts.Show
'   End If
'
'End Sub

'Private Sub txtPrt_LostFocus()
'   txtPrt = CheckLen(txtPrt, 30)
'   If Trim(txtPrt) = "" Then txtPrt = "ALL"
'
'End Sub



'Least to greatest dates 10/12/01

Private Sub GetMRPDates()
   Dim RdoDte As ADODB.Recordset
    sSql = "SELECT MIN(MRP_PARTDATERQD) FROM MrplTable WHERE " _
           & "MRP_TYPE>" & MRPTYPE_BeginningBalance
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtBeg = Format(.Fields(0), "mm/dd/yyyy")
         Else
            txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtBeg.ToolTipText = "Earliest Date By Default"
   
   sSql = "SELECT MAX(MRP_PARTDATERQD) FROM MrplTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtEnd = Format(.Fields(0), "mm/dd/yyyy")
         Else
            txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtEnd.ToolTipText = "Latest Date By Default"
   Set RdoDte = Nothing
End Sub

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPart.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPart.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function


