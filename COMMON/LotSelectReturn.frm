VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form LotSelectReturn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Available Lots"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkReturn 
      Caption         =   "Return to Inventory"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1800
      TabIndex        =   80
      Top             =   4980
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Update"
      Height          =   315
      Left            =   7560
      TabIndex        =   60
      ToolTipText     =   "Update To The Selected Items And Quit"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   10
      Left            =   6600
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   9
      Left            =   6600
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   8
      Left            =   6600
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   7
      Left            =   6600
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   6
      Left            =   6600
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   5
      Left            =   6600
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   4
      Left            =   6600
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   3
      Left            =   6600
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   2
      Left            =   6600
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Index           =   1
      Left            =   6600
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   7560
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
      FormDesignWidth =   8550
   End
   Begin Threed.SSCommand cmdDn 
      Height          =   375
      Left            =   8040
      TabIndex        =   61
      TabStop         =   0   'False
      ToolTipText     =   "Next Page (Page Down)"
      Top             =   5160
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "LotSelectReturn.frx":0000
   End
   Begin Threed.SSCommand cmdUp 
      Height          =   375
      Left            =   8040
      TabIndex        =   62
      TabStop         =   0   'False
      ToolTipText     =   "Last Page (Page Up)"
      Top             =   4800
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "LotSelectReturn.frx":0502
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
      Left            =   7800
      TabIndex        =   79
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   7800
      TabIndex        =   78
      ToolTipText     =   "Location"
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   7800
      TabIndex        =   77
      ToolTipText     =   "Location"
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   7800
      TabIndex        =   76
      ToolTipText     =   "Location"
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   7800
      TabIndex        =   75
      ToolTipText     =   "Location"
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   7800
      TabIndex        =   74
      ToolTipText     =   "Location"
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   7800
      TabIndex        =   73
      ToolTipText     =   "Location"
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   7800
      TabIndex        =   72
      ToolTipText     =   "Location"
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   7800
      TabIndex        =   71
      ToolTipText     =   "Location"
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   7800
      TabIndex        =   70
      ToolTipText     =   "Location"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   7800
      TabIndex        =   69
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
      TabIndex        =   68
      Top             =   600
      Width           =   960
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rq'd Qty"
      Height          =   255
      Index           =   9
      Left            =   2760
      TabIndex        =   67
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblTotQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   6240
      TabIndex        =   66
      Top             =   600
      Width           =   945
   End
   Begin VB.Label lblTotLots 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   1800
      TabIndex        =   65
      Top             =   600
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   255
      Index           =   8
      Left            =   5400
      TabIndex        =   64
      Top             =   600
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lots Selected"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   63
      Top             =   600
      Width           =   1815
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   960
      Picture         =   "LotSelectReturn.frx":0A04
      Top             =   4920
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   240
      Picture         =   "LotSelectReturn.frx":0EF6
      Top             =   4920
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   480
      Picture         =   "LotSelectReturn.frx":13E8
      Top             =   4920
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   720
      Picture         =   "LotSelectReturn.frx":18DA
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
      Top             =   240
      Width           =   945
   End
   Begin VB.Label lblPart 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   58
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sorted"
      Height          =   255
      Index           =   6
      Left            =   5400
      TabIndex        =   57
      Top             =   240
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lots For P/N              "
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   56
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   4440
      TabIndex        =   55
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
      Top             =   4440
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5400
      TabIndex        =   52
      ToolTipText     =   "Click To Add"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   5400
      TabIndex        =   51
      ToolTipText     =   "Click To Add"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Top             =   4080
      Width           =   2580
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   4440
      TabIndex        =   48
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   5400
      TabIndex        =   47
      ToolTipText     =   "Click To Add"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Top             =   3720
      Width           =   2580
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   4440
      TabIndex        =   44
      Top             =   3720
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   5400
      TabIndex        =   43
      ToolTipText     =   "Click To Add"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Top             =   3360
      Width           =   2580
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   4440
      TabIndex        =   40
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   5400
      TabIndex        =   39
      ToolTipText     =   "Click To Add"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Top             =   3000
      Width           =   2580
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   4440
      TabIndex        =   36
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   5400
      TabIndex        =   35
      ToolTipText     =   "Click To Add"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Top             =   2640
      Width           =   2580
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   4440
      TabIndex        =   32
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   5400
      TabIndex        =   31
      ToolTipText     =   "Click To Add"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Top             =   2280
      Width           =   2580
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   28
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   5400
      TabIndex        =   27
      ToolTipText     =   "Click To Add"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Top             =   1920
      Width           =   2580
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4440
      TabIndex        =   24
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   5400
      TabIndex        =   23
      ToolTipText     =   "Click To Add"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Top             =   1560
      Width           =   2580
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   4440
      TabIndex        =   20
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   5400
      TabIndex        =   19
      ToolTipText     =   "Click To Add"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Top             =   1200
      Width           =   2580
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   16
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
      Left            =   6600
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
      Left            =   5400
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
      Left            =   4440
      TabIndex        =   13
      Top             =   960
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID                                              "
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
      Width           =   2535
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
Attribute VB_Name = "LotSelectReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'All new 2/19/03
'10/5/04 See UpdateQty (limited the precise amount)
'4/7/05 Added Lot Location
'4/12/05 Resolved blank Boxes (GetTheNextGroup)
'6/1/05 Lot Locations (INTCOA) GetSettings
'
Option Explicit
Dim bOnLoad As Byte
Dim bGoodQty As Byte
Dim bLotLocOn As Byte

Dim iCurrPage As Integer
Dim iIndex As Integer
Dim iLastPage As Integer
Dim iTotalItems As Integer

Dim sLoc(5) As String
Public SqlFilter As String       'to refine search if nonblank

Dim vLots(300, 8) As Variant
'0 = SysId
'1 = UserId
'2 = PartRef
'3 = ADate
'4 = Remaining
'5 = Cost
'6 = Qty Selected
'7 = Location
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
   CheckQuantity
   
End Sub



Private Sub cmdDn_Click()
   'Next
   iCurrPage = iCurrPage + 1
   If iCurrPage > iLastPage Then iCurrPage = iLastPage
   GetTheNextGroup
   
End Sub

Private Sub cmdNew_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If bGoodQty = 0 Then
      sMsg = "The Required Quantity Does Not Equal The " & vbCr _
             & "Selected Quantity.  You Must Add Lots To Continue."
      MsgBox sMsg, vbExclamation, Caption
      Exit Sub
   End If
   If Val(lblTotQty) = 0 Then
      sMsg = "You Have Not Selected Quantities From" & vbCr _
             & "Lots.  You Must Add Lots To Continue."
      MsgBox sMsg, vbExclamation, Caption
      Exit Sub
   End If
   If lblTotQty.ForeColor = ES_RED Then
      sMsg = "You Have Selected A Greater Quantity Than Required" & vbCr _
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
   MdiSect.enabled = False
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


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   MdiSect.enabled = True
   Set LotSelect = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub



Private Sub GetLots()
   Dim RdoLot As ADODB.Recordset
   Dim bRow As Byte
   Dim bIncludeLot As Byte
   Dim iRow As Integer
   Dim cQty As Currency
   Dim sLocation As String
   On Error GoTo DiaErr1
   Erase vLots
   Es_LotsSelected = 0
   Es_TotalLots = 0
   iTotalItems = 0
   '    sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTPARTREF,LOTADATE," _
   '        & "LOTREMAININGQTY,LOTUNITCOST,LOTLOCATION " _
   '        & "FROM LohdTable WHERE (LOTPARTREF='" _
   '        & Compress(lblPart) & "' AND LOTREMAININGQTY>0) "
   sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTPARTREF,LOTADATE," _
          & "LOTREMAININGQTY,LOTUNITCOST,LOTLOCATION " & vbCrLf _
          & "FROM LohdTable WHERE LOTPARTREF='" _
          & Compress(lblPart) & "'" & vbCrLf
          
   If SqlFilter <> "" Then
      sSql = sSql & "and " & SqlFilter
   End If
   
   If chkReturn.Value = 0 Then
      sSql = sSql & " AND LOTREMAININGQTY > 0"
   End If
   If lblLifo = "LIFO" Then sSql = sSql & " ORDER BY LOTNUMBER DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
   If bSqlRows Then
      With RdoLot
         Do Until .EOF
            'Location approval
            sLocation = "" & Trim(!LOTLOCATION)
            bIncludeLot = 0
            If bLotLocOn = 1 Then
               For bRow = 1 To 4
                  'Blank Locations are accepted by proxy
                  If sLoc(bRow) = sLocation Then bIncludeLot = 1
               Next
            Else
               bIncludeLot = 1
            End If
            If bIncludeLot = 1 Then
               'include this group
               iTotalItems = iTotalItems + 1
               iRow = iRow + 1
               vLots(iRow, 0) = Format$(iRow, "00000")
               vLots(iRow, 4) = Format(cQty, ES_QuantityDataFormat)
               vLots(iRow, 6) = "0.000"
               
               vLots(iRow, 0) = "" & Trim(!lotNumber)
               vLots(iRow, 1) = "" & Trim(!LOTUSERLOTID)
               vLots(iRow, 2) = "" & Trim(!LotPartRef)
               vLots(iRow, 3) = Format(!LotADate, "mm/dd/yy")
               vLots(iRow, 4) = Format(!LOTREMAININGQTY, ES_QuantityDataFormat)
               vLots(iRow, 5) = Format(!LotUnitCost, ES_QuantityDataFormat)
               vLots(iRow, 6) = "0.000"
               vLots(iRow, 7) = "" & Trim(!LOTLOCATION)
            End If
            .MoveNext
         Loop
      End With
   End If
   lblTotLots = iTotalItems
   For iRow = 1 To iTotalItems
      If iRow > 10 Then Exit For
      lblLots(iRow) = vLots(iRow, 0)
      lblLotu(iRow) = vLots(iRow, 1)
      lblDate(iRow) = vLots(iRow, 3)
      lblQty(iRow) = vLots(iRow, 4)
      lblLotLoc(iRow) = vLots(iRow, 7)
      lblLots(iRow).Visible = True
      lblLotu(iRow).Visible = True
      lblDate(iRow).Visible = True
      lblQty(iRow).Visible = True
      txtQty(iRow).Visible = True
      lblLotLoc(iRow).Visible = True
   Next
   If iTotalItems > 10 Then
      cmdUp.enabled = True
      cmdUp.Picture = Enup.Picture
      cmdDn.enabled = True
      cmdDn.Picture = Endn.Picture
   End If
   If txtQty(1).Visible Then
      iLastPage = 0.4 + (iTotalItems / 10)
      iIndex = 0
      txtQty(1).SetFocus
   End If
   Set RdoLot = Nothing
   Exit Sub
   
DiaErr1:
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
      lblLots(iList) = ""
      lblLotu(iList) = ""
      lblDate(iList) = ""
      lblQty(iList) = ""
      txtQty(iList) = "0.000"
      lblLotLoc(iList) = ""
   Next
   
End Sub

Private Sub GetTheNextGroup()
   Dim iRow As Integer
   Dim iList As Integer
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
   On Error Resume Next
   For iList = iIndex + 1 To iTotalItems
      iRow = iRow + 1
      If iRow > 10 Then Exit For
      lblLots(iRow) = vLots(iList, 0)
      lblLotu(iRow) = vLots(iList, 1)
      lblDate(iRow) = vLots(iList, 3)
      lblQty(iRow) = vLots(iList, 4)
      txtQty(iRow) = Format(vLots(iList, 6), ES_QuantityDataFormat)
      lblLotLoc(iRow) = vLots(iList, 7)
      lblLots(iRow).Visible = True
      lblLotu(iRow).Visible = True
      lblDate(iRow).Visible = True
      lblQty(iRow).Visible = True
      txtQty(iRow).Visible = True
      lblLotLoc(iRow).Visible = True
   Next
   If txtQty(1).Visible Then txtQty(1).SetFocus
   
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
   If chkReturn.Value = 0 Then
      If Val(txtQty(Index)) > 0 Then
         If Val(txtQty(Index)) > Val(lblQty(Index)) Then
            'Beep
            txtQty(Index) = lblQty(Index)
         End If
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
      sMsg = "You Have Selected Some Quantities From" & vbCr _
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
            .LotADate = vLots(iRow, 3)
            .LotRemQty = Format(Val(vLots(iRow, 4)), ES_QuantityDataFormat)
            .LotCost = Format(Val(vLots(iRow, 5)), ES_QuantityDataFormat)
            .LotSelQty = Format(Val(vLots(iRow, 6)), ES_QuantityDataFormat)
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
