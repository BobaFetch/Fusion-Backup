VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form LotSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Available Lots"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   932
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "LotSelect.frx":0000
      DownPicture     =   "LotSelect.frx":04F2
      Height          =   372
      Left            =   8000
      MaskColor       =   &H00000000&
      Picture         =   "LotSelect.frx":09E4
      Style           =   1  'Graphical
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   4800
      Width           =   400
   End
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "LotSelect.frx":0ED6
      DownPicture     =   "LotSelect.frx":13C8
      Height          =   372
      Left            =   8000
      MaskColor       =   &H00000000&
      Picture         =   "LotSelect.frx":18BA
      Style           =   1  'Graphical
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   5184
      Width           =   400
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotSelect.frx":1DAC
      Style           =   1  'Graphical
      TabIndex        =   78
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Update"
      Height          =   315
      Left            =   9480
      TabIndex        =   60
      ToolTipText     =   "Update To The Selected Items And Quit"
      Top             =   540
      Width           =   875
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   10
      Left            =   7620
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   9
      Left            =   7620
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   8
      Left            =   7620
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   7
      Left            =   7620
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   6
      Left            =   7620
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   5
      Left            =   7620
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   4
      Left            =   7620
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   3
      Left            =   7620
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   2
      Left            =   7620
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   1
      Left            =   7620
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   9480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   5520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5640
      FormDesignWidth =   10515
   End
   Begin VB.Label lblLotAdate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4080
      TabIndex        =   92
      Top             =   4800
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "   Expires   "
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
      Left            =   9540
      TabIndex        =   91
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblExpires 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   9480
      TabIndex        =   90
      ToolTipText     =   "Creation Date"
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label lblExpires 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   9480
      TabIndex        =   89
      ToolTipText     =   "Creation Date"
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label lblExpires 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   9480
      TabIndex        =   88
      ToolTipText     =   "Creation Date"
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label lblExpires 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   9480
      TabIndex        =   87
      ToolTipText     =   "Creation Date"
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label lblExpires 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   9480
      TabIndex        =   86
      ToolTipText     =   "Creation Date"
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label lblExpires 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   9480
      TabIndex        =   85
      ToolTipText     =   "Creation Date"
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label lblExpires 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   9480
      TabIndex        =   84
      ToolTipText     =   "Creation Date"
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label lblExpires 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   9480
      TabIndex        =   83
      ToolTipText     =   "Creation Date"
      Top             =   3720
      Width           =   900
   End
   Begin VB.Label lblExpires 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   9480
      TabIndex        =   82
      ToolTipText     =   "Creation Date"
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label lblExpires 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   9480
      TabIndex        =   81
      ToolTipText     =   "Creation Date"
      Top             =   4440
      Width           =   900
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location "
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
      Index           =   10
      Left            =   8820
      TabIndex        =   77
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   8820
      TabIndex        =   76
      ToolTipText     =   "Location"
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   8820
      TabIndex        =   75
      ToolTipText     =   "Location"
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   8820
      TabIndex        =   74
      ToolTipText     =   "Location"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   8820
      TabIndex        =   73
      ToolTipText     =   "Location"
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   8820
      TabIndex        =   72
      ToolTipText     =   "Location"
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   8820
      TabIndex        =   71
      ToolTipText     =   "Location"
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   8820
      TabIndex        =   70
      ToolTipText     =   "Location"
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   8820
      TabIndex        =   69
      ToolTipText     =   "Location"
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   8820
      TabIndex        =   68
      ToolTipText     =   "Location"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   8820
      TabIndex        =   67
      ToolTipText     =   "Location"
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblRequired 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "39.000"
      Height          =   285
      Left            =   3600
      TabIndex        =   66
      ToolTipText     =   "Required To Be Selected"
      Top             =   620
      Width           =   960
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rq'd Qty"
      Height          =   255
      Index           =   9
      Left            =   2760
      TabIndex        =   65
      Top             =   620
      Width           =   1215
   End
   Begin VB.Label lblTotQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   6240
      TabIndex        =   64
      ToolTipText     =   "Quantity Selected To This Point"
      Top             =   620
      Width           =   945
   End
   Begin VB.Label lblTotLots 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   1800
      TabIndex        =   63
      Top             =   620
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Selected"
      Height          =   255
      Index           =   8
      Left            =   4800
      TabIndex        =   62
      ToolTipText     =   "Quantity Selected To This Point"
      Top             =   615
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lots Available"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   61
      Top             =   620
      Width           =   1815
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   960
      Picture         =   "LotSelect.frx":255A
      Top             =   4920
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   240
      Picture         =   "LotSelect.frx":2A4C
      Top             =   4920
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   480
      Picture         =   "LotSelect.frx":2F3E
      Top             =   4920
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   720
      Picture         =   "LotSelect.frx":3430
      Top             =   4920
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblLifo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   59
      Top             =   280
      Width           =   945
   End
   Begin VB.Label lblPart 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   58
      Top             =   280
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sorted"
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   57
      Top             =   285
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lots For P/N              "
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   56
      Top             =   280
      Width           =   1935
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   5460
      TabIndex        =   55
      ToolTipText     =   "Creation Date"
      Top             =   4440
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   1800
      TabIndex        =   54
      ToolTipText     =   "User Lot ID"
      Top             =   4440
      Width           =   3600
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System Lot ID"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   53
      Top             =   4440
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   6420
      TabIndex        =   52
      ToolTipText     =   "Click To Make Selected = Remaining Quantity"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   6420
      TabIndex        =   51
      ToolTipText     =   "Click To Make Selected = Remaining Quantity"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System Lot ID"
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   50
      Top             =   4080
      Width           =   1500
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   1800
      TabIndex        =   49
      ToolTipText     =   "User Lot ID"
      Top             =   4080
      Width           =   3600
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   5460
      TabIndex        =   48
      ToolTipText     =   "Creation Date"
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   6420
      TabIndex        =   47
      ToolTipText     =   "Click To Make Selected = Remaining Quantity"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System Lot ID"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   46
      Top             =   3720
      Width           =   1500
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   1800
      TabIndex        =   45
      ToolTipText     =   "User Lot ID"
      Top             =   3720
      Width           =   3600
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   5460
      TabIndex        =   44
      ToolTipText     =   "Creation Date"
      Top             =   3720
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   6420
      TabIndex        =   43
      ToolTipText     =   "Click To Make Selected = Remaining Quantity"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System Lot ID"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   42
      Top             =   3360
      Width           =   1500
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   1800
      TabIndex        =   41
      ToolTipText     =   "User Lot ID"
      Top             =   3360
      Width           =   3600
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   5460
      TabIndex        =   40
      ToolTipText     =   "Creation Date"
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   6420
      TabIndex        =   39
      ToolTipText     =   "Click To Make Selected = Remaining Quantity"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System Lot ID"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   38
      Top             =   3000
      Width           =   1500
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   1800
      TabIndex        =   37
      ToolTipText     =   "User Lot ID"
      Top             =   3000
      Width           =   3600
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   5460
      TabIndex        =   36
      ToolTipText     =   "Creation Date"
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   6420
      TabIndex        =   35
      ToolTipText     =   "Click To Make Selected = Remaining Quantity"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System Lot ID"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   34
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   33
      ToolTipText     =   "User Lot ID"
      Top             =   2640
      Width           =   3600
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   5460
      TabIndex        =   32
      ToolTipText     =   "Creation Date"
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   6420
      TabIndex        =   31
      ToolTipText     =   "Click To Make Selected = Remaining Quantity"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System Lot ID"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   30
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   29
      ToolTipText     =   "User Lot ID"
      Top             =   2280
      Width           =   3600
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   5460
      TabIndex        =   28
      ToolTipText     =   "Creation Date"
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   6420
      TabIndex        =   27
      ToolTipText     =   "Click To Make Selected = Remaining Quantity"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System Lot ID"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   26
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   25
      ToolTipText     =   "User Lot ID"
      Top             =   1920
      Width           =   3600
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   5460
      TabIndex        =   24
      ToolTipText     =   "Creation Date"
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   6420
      TabIndex        =   23
      ToolTipText     =   "Click To Make Selected = Remaining Quantity"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System Lot ID"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   22
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   21
      ToolTipText     =   "User Lot ID"
      Top             =   1560
      Width           =   3600
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   5460
      TabIndex        =   20
      ToolTipText     =   "Creation Date"
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   6420
      TabIndex        =   19
      ToolTipText     =   "Click To Make Selected = Remaining Quantity"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System Lot ID"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   18
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   17
      ToolTipText     =   "User Lot ID"
      Top             =   1200
      Width           =   3600
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   5460
      TabIndex        =   16
      ToolTipText     =   "Creation Date"
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Qty        "
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
      Left            =   7620
      TabIndex        =   15
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining Qty  "
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
      Left            =   6420
      TabIndex        =   14
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date            "
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
      Left            =   5460
      TabIndex        =   13
      Top             =   960
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID                                                                   "
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
      Left            =   1800
      TabIndex        =   12
      Top             =   960
      Width           =   3600
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "System ID                 "
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
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "LotSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'All new 2/19/03
'10/5/04 See UpdateQty (limited the precise amount)
'4/7/05 Added Lot Location
'4/12/05 Resolved blank Boxes (GetTheNextGroup)
'6/1/05 Lot Locations (INTCOA) GetSettings
'6/15/05 Added .MoreResults and removed MdiSect disabling
Option Explicit
Dim bOnLoad As Byte
Dim bGoodQty As Byte
Dim bLotLocOn As Byte
'Set strLotAdate = ""

Dim iCurrPage As Integer
Dim iIndex As Integer
Dim iLastPage As Integer
Dim iTotalItems As Integer

Dim sLoc(5) As String

Dim vLots(450, 10) As Variant
'0 = SysId
'1 = UserId
'2 = PartRef
'3 = ADate
'4 = Remaining
'5 = Cost
'6 = Qty Selected
'7 = Location
'8 = CanShipFromThisLoc
'9 = Lot Expiration Date
Public strLotAdate As String
Private Const LOT_Date = 3
Private Const LOT_Location = 7
Private Const LOT_CanShipFromThisLoc = 8
Private Const LOT_LotExpirationDate = 9


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetSettings()
   bLotLocOn = GetSetting("Esi2000", "WSLotLocs", "On", bLotLocOn)
   If bLotLocOn = 1 Then
      sLoc(1) = GetSetting("Esi2000", "WSLotLocs", "Loc1", sLoc(1))
      sLoc(2) = GetSetting("Esi2000", "WSLotLocs", "Loc2", sLoc(2))
      sLoc(3) = GetSetting("Esi2000", "WSLotLocs", "Loc3", sLoc(3))
      sLoc(4) = GetSetting("Esi2000", "WSLotLocs", "Loc4", sLoc(4))
   End If
End Sub

Private Sub cmdCan_Click()
   AlwaysOnTop hWnd, False
   CheckQuantity
   
End Sub



Private Sub cmdDn_Click()
   'Next
   iCurrPage = iCurrPage + 1
   If iCurrPage > iLastPage Then iCurrPage = iLastPage
   GetTheNextGroup
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 932
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdNew_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If bGoodQty = 0 Then
      sMsg = "The Required Quantity Does Not Equal The " & vbCrLf _
             & "Selected Quantity.  You Must Add Lots To Continue."
      MsgBox sMsg, vbExclamation, Caption
      Exit Sub
   End If
   If Val(lblTotQty) = 0 Then
      sMsg = "You Have Not Selected Quantities From" & vbCrLf _
             & "Lots.  You Must Add Lots To Continue."
      MsgBox sMsg, vbExclamation, Caption
      Exit Sub
   End If
   If lblTotQty.ForeColor = ES_RED Then
      sMsg = "You Have Selected A Greater Quantity Than Required" & vbCrLf _
             & "You Must Adjust The Quantity To The Amount Needed."
      MsgBox sMsg, vbExclamation, Caption
      Exit Sub
   End If
   'All is ducky
   UpdateLotArray
   
End Sub

Private Sub cmdUp_Click()
   'last
   iCurrPage = iCurrPage - 1
   If iCurrPage < 1 Then iCurrPage = 1
   GetTheNextGroup
   
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   If bOnLoad Then
      GetSettings
      CloseBoxes
      b = GetInventoryMethod()
      If b = 1 Then lblLifo = "FIFO" Else lblLifo = "LIFO"
      GetLots
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   If iBarOnTop Then
      Move MdiSect.Left + 500, MdiSect.Top + 1400
   Else
      Move MdiSect.Left + 2100, MdiSect.Top + 1000
   End If
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'MdiSect.Enabled = True
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   ' AlwaysOnTop hWnd, False
   Set LotSelect = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub



Private Sub GetLots()
   Dim RdoLot As ADODB.Recordset
   Dim bRow As Byte
   Dim bCanShip As Byte
   Dim iRow As Integer
   Dim cQty As Currency
   Dim sLocation As String
   On Error GoTo DiaErr1
   Erase vLots
   Es_LotsSelected = 0
   Es_TotalLots = 0
   iTotalItems = 0
   
   ' 4/28/2010 This flag is set from PackPSp01a for Packslip printing back dated.
   If strLotAdate <> "" Then
      sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTPARTREF,LOTADATE," _
            & "LOTREMAININGQTY,LOTUNITCOST,LOTLOCATION, LOTEXPIRESON " _
            & "FROM LohdTable WHERE (LOTPARTREF='" _
            & Compress(lblPart) & "' AND LOTREMAININGQTY > 0 AND LOTADATE < DATEADD(dd, 1 ,'" & Trim(strLotAdate) & "')) "
   Else
      sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTPARTREF,LOTADATE," _
            & "LOTREMAININGQTY,LOTUNITCOST,LOTLOCATION, LOTEXPIRESON " _
            & "FROM LohdTable WHERE (LOTPARTREF='" _
            & Compress(lblPart) & "' AND LOTREMAININGQTY > 0) "
   End If
   
   If lblLifo = "LIFO" Then sSql = sSql & " ORDER BY LOTNUMBER DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
   If bSqlRows Then
      With RdoLot
         Do Until .EOF
         
            ' Check for Array size
            If (iRow >= 450) Then
               Exit Do
            End If
            
            'Location approval
            sLocation = "" & Trim(!LOTLOCATION)
            bCanShip = 0
            If bLotLocOn = 1 Then
               
               'always allow shipping from blank lot location
               If sLocation = "" Then
                  bCanShip = 1
               Else
                  For bRow = 1 To 4
                     If sLoc(bRow) = sLocation Then bCanShip = 1
                  Next
               End If
            Else
               bCanShip = 1
            End If
            
            iTotalItems = iTotalItems + 1
            iRow = iRow + 1
            vLots(iRow, 0) = Format$(iRow, SO_NUM_FORMAT)
            vLots(iRow, 4) = Format(cQty, ES_QuantityDataFormat)
            vLots(iRow, 6) = "0.000"
            
            vLots(iRow, 0) = "" & Trim(!lotNumber)
            vLots(iRow, 1) = "" & Trim(!LOTUSERLOTID)
            vLots(iRow, 2) = "" & Trim(!LotPartRef)
            vLots(iRow, LOT_Date) = Format(!LotADate, "mm/dd/yy")
            vLots(iRow, 4) = Format(!LOTREMAININGQTY, ES_QuantityDataFormat)
            vLots(iRow, 5) = Format(!LotUnitCost, ES_QuantityDataFormat)
            vLots(iRow, 6) = "0.000"
            vLots(iRow, LOT_Location) = sLocation
            vLots(iRow, LOT_CanShipFromThisLoc) = bCanShip
            vLots(iRow, LOT_LotExpirationDate) = IIf(IsDate(!LOTEXPIRESON), _
               Format(!LOTEXPIRESON, "mm/dd/yy"), "")
            
            .MoveNext
         Loop
         ClearResultSet RdoLot
      End With
   End If
   lblTotLots = iTotalItems
   CloseBoxes
   '    For iRow = 1 To iTotalItems
   '        If iRow > 10 Then Exit For
   '        lblLots(iRow) = vLots(iRow, 0)
   '        lblLotu(iRow) = vLots(iRow, 1)
   '        lblDate(iRow) = vLots(iRow, LOT_Date)
   '        lblQty(iRow) = vLots(iRow, 4)
   '        lblLotLoc(iRow) = vLots(iRow, LOT_Location)
   '        lblLots(iRow).Visible = True
   '        lblLotu(iRow).Visible = True
   '        lblDate(iRow).Visible = True
   '        lblQty(iRow).Visible = True
   '        txtQty(iRow).Visible = True
   '        lblLotLoc(iRow).Visible = True
   '    Next
   '    If iTotalItems > 10 Then
   '        cmdUp.Enabled = True
   '        cmdUp.Picture = Enup.Picture
   '        cmdDn.Enabled = True
   '        cmdDn.Picture = Endn.Picture
   '    End If
   
   'If txtQty(1).Visible Then
   iLastPage = 0.4 + (iTotalItems / 10)
   '        iIndex = 0
   '        txtQty(1).SetFocus
   iCurrPage = 1
   'End If
   
   GetTheNextGroup
   Set RdoLot = Nothing
   Exit Sub
   
DiaErr1:
   Set RdoLot = Nothing
   sProcName = "getlots"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub CloseBoxes()
   Dim iList As Integer
   For iList = 1 To 10
      lblLots(iList).Visible = False
      lblLotu(iList).Visible = False
      lblDate(iList).Visible = False
      lblQty(iList).Visible = False
      txtQty(iList).Visible = False
      lblLotLoc(iList).Visible = False
      lblExpires(iList).Visible = False
      lblLots(iList) = ""
      lblLotu(iList) = ""
      lblDate(iList) = ""
      lblQty(iList) = ""
      txtQty(iList) = "0.000"
      lblLotLoc(iList) = ""
      lblExpires(iList) = ""
      lblExpires(iList).ForeColor = Es_TextForeColor
   Next
   cmdUp.Enabled = False
   cmdDn.Enabled = False
   
End Sub

Private Sub GetTheNextGroup()
   Dim iRow As Integer
   Dim iList As Integer
   Dim focusSet As Boolean
   CloseBoxes
   iIndex = (iCurrPage - 1) * 10
   
   '0 = SysId
   '1 = UserId
   '2 = PartRef
   '3 = ADate
   '4 = Remaining
   '5 = Cost
   '6 = Qty Selected
   '7 = Location
   'On Error Resume Next
   For iList = iIndex + 1 To iTotalItems
      iRow = iRow + 1
      If iRow > 10 Then Exit For
      lblLots(iRow) = vLots(iList, 0)
      lblLotu(iRow) = vLots(iList, 1)
      lblDate(iRow) = vLots(iList, LOT_Date)
      lblQty(iRow) = vLots(iList, 4)
      txtQty(iRow) = Format(vLots(iList, 6), ES_QuantityDataFormat)
      lblLotLoc(iRow) = vLots(iList, LOT_Location)
      lblExpires(iRow) = vLots(iList, LOT_LotExpirationDate)
      If IsDate(lblExpires(iRow)) Then
         If DateDiff("d", CDate(lblExpires(iRow)), Now) > 0 Then
            lblExpires(iRow).ForeColor = ES_RED
         Else
            lblExpires(iRow).ForeColor = Es_TextForeColor
         End If
      End If

      lblLots(iRow).Visible = True
      lblLotu(iRow).Visible = True
      lblDate(iRow).Visible = True
      lblQty(iRow).Visible = True
      txtQty(iRow).Visible = True
      lblLotLoc(iRow).Visible = True
      lblExpires(iRow).Visible = True
      
      If Val(vLots(iList, LOT_CanShipFromThisLoc)) = 1 Then
         txtQty(iRow).Enabled = True
         txtQty(iRow).ToolTipText = ""
         lblLotLoc(iRow).ToolTipText = ""
         If Not focusSet Then
            focusSet = True
            txtQty(iRow).SetFocus
         End If
      Else
         txtQty(iRow).Enabled = False
         txtQty(iRow).ToolTipText = "You can't ship from this location"
         lblLotLoc(iRow).ToolTipText = "You can't ship from this location"
      End If
   Next
   'If txtQty(1).Visible Then txtQty(1).SetFocus
   
   'determine whether to show up and down arrows for prev and next pages
   If iCurrPage > 1 Then
      cmdUp.Enabled = True
      cmdUp.Picture = Enup.Picture
   End If
   If iCurrPage < iLastPage Then
      cmdDn.Enabled = True
      cmdDn.Picture = Endn.Picture
   End If
   
End Sub

Private Sub lblQty_Click(Index As Integer)
   txtQty(Index) = lblQty(Index)
   vLots(Index + iIndex, 6) = Format(Val(txtQty(Index)), ES_QuantityDataFormat)
   UpdateQty
   
End Sub

Private Sub txtQty_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtQty_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtQty_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtQty_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub txtQty_Validate(Index As Integer, Cancel As Boolean)
   txtQty(Index) = CheckLen(txtQty(Index), 9)
   txtQty(Index) = Format(Abs(Val(txtQty(Index))), ES_QuantityDataFormat)
   If Val(txtQty(Index)) > 0 Then
      If Val(txtQty(Index)) > Val(lblQty(Index)) Then
         'Beep
         txtQty(Index) = lblQty(Index)
      End If
   End If
   vLots(Index + iIndex, 6) = Format(Val(txtQty(Index)), ES_QuantityDataFormat)
   UpdateQty
   
End Sub



'Updates the quantity selected box
'10/5/04 makes rounding less ridged

Private Sub UpdateQty()
   Dim iRow As Integer
   Dim cItems As Currency
   Dim cSelQty As Currency '10/5/04
   Dim cReqQty As Currency '10/5/04
   
   cReqQty = Format(Val(lblRequired), ES_QuantityDataFormat)
   cItems = 0
   For iRow = 1 To iTotalItems
      cSelQty = Format(Val(vLots(iRow, 6)), ES_QuantityDataFormat)
      cItems = cItems + cSelQty
   Next
   '10/5/04
   lblTotQty = Format(cItems, ES_QuantityDataFormat)
   If Format(cItems, "#######0.0") > Format(cReqQty, "#######0.0") Then
      lblTotQty.ForeColor = ES_RED
   Else
      lblTotQty.ForeColor = Es_TextForeColor
   End If
   If Format(cItems, "#######0.0") <> Format(cReqQty, "#######0.0") Then
      bGoodQty = 0
   Else
      bGoodQty = 1
   End If
   
End Sub

Private Sub CheckQuantity()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If Val(lblTotQty) > 0 Then
      sMsg = "You Have Selected Some Quantities From" & vbCrLf _
             & "Lots.  Do You Really Want To Cancel?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         Es_LotsSelected = 0
         Unload Me
      End If
   Else
      sMsg = "Do You Really Want To Cancel This Transaction?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         Es_LotsSelected = 0
         Unload Me
      End If
   End If
   
End Sub

'0 = SysId
'1 = UserId
'2 = PartRef
'3 = ADate
'4 = Remaining Qty
'5 = Cost
'6 = Qty Selected
'Pass the finished array

Private Sub UpdateLotArray()
   Dim iRow As Integer
   Dim iList As Integer
   
   On Error GoTo DiaErr1
   Erase lots
   For iRow = 1 To iTotalItems
      If Val(vLots(iRow, 6)) > 0 Then
         iList = iList + 1
         Es_TotalLots = Es_TotalLots + 1
         ReDim Preserve lots(iList)
         With lots(iList)
            .LotSysId = vLots(iRow, 0)
            .LotUserId = vLots(iRow, 1)
            .LotPartRef = vLots(iRow, 2)
            .LotADate = vLots(iRow, LOT_Date)
            .LotRemQty = Format(Val(vLots(iRow, 4)), ES_QuantityDataFormat)
            .LotCost = Format(Val(vLots(iRow, 5)), ES_QuantityDataFormat)
            .LotSelQty = Format(Val(vLots(iRow, 6)), ES_QuantityDataFormat)
            .LotExpirationDate = vLots(iRow, LOT_LotExpirationDate)
         End With
      End If
   Next
   Es_LotsSelected = 1
   Unload Me
   Exit Sub
DiaErr1:
   sProcName = "updatelotarr"
   Es_LotsSelected = 0
   Es_TotalLots = 0
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   Unload Me
   
End Sub
