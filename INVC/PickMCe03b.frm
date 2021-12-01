VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form PickMCe03b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restock Lot Tracked Items"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLst 
      Caption         =   "<< &Last    "
      Enabled         =   0   'False
      Height          =   315
      Left            =   6600
      TabIndex        =   108
      Top             =   7200
      Width           =   875
   End
   Begin VB.CommandButton cmdNxt 
      Caption         =   " &Next >>"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7500
      TabIndex        =   107
      Top             =   7200
      Width           =   875
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   15
      Left            =   6480
      TabIndex        =   101
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   14
      Left            =   6480
      TabIndex        =   95
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   13
      Left            =   6480
      TabIndex        =   89
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   6480
      TabIndex        =   83
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   6480
      TabIndex        =   77
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   6480
      TabIndex        =   7
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   6480
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   6480
      TabIndex        =   9
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   6480
      TabIndex        =   0
      ToolTipText     =   "Quantity To Restock"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   6480
      TabIndex        =   1
      ToolTipText     =   "Quantity To Restock"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   6480
      TabIndex        =   2
      ToolTipText     =   "Quantity To Restock"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   6480
      TabIndex        =   3
      ToolTipText     =   "Quantity To Restock"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   6480
      TabIndex        =   4
      ToolTipText     =   "Quantity To Restock"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   6480
      TabIndex        =   5
      ToolTipText     =   "Quantity To Restock"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   6480
      TabIndex        =   6
      ToolTipText     =   "Quantity To Restock"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   7560
      TabIndex        =   22
      ToolTipText     =   "Update To The Selected Items And Quit"
      Top             =   1080
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7560
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7800
      Top             =   7920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8070
      FormDesignWidth =   8595
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   15
      Left            =   7680
      TabIndex        =   106
      ToolTipText     =   "Location"
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   15
      Left            =   4320
      TabIndex        =   105
      ToolTipText     =   "Date Of Pick"
      Top             =   6720
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   15
      Left            =   1680
      TabIndex        =   104
      ToolTipText     =   "User Lot ID"
      Top             =   6720
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   15
      Left            =   120
      TabIndex        =   103
      ToolTipText     =   "System Lot ID"
      Top             =   6720
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   15
      Left            =   5280
      TabIndex        =   102
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   7680
      TabIndex        =   100
      ToolTipText     =   "Location"
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   4320
      TabIndex        =   99
      ToolTipText     =   "Date Of Pick"
      Top             =   6360
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   1680
      TabIndex        =   98
      ToolTipText     =   "User Lot ID"
      Top             =   6360
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   120
      TabIndex        =   97
      ToolTipText     =   "System Lot ID"
      Top             =   6360
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   5280
      TabIndex        =   96
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   13
      Left            =   7680
      TabIndex        =   94
      ToolTipText     =   "Location"
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   13
      Left            =   4320
      TabIndex        =   93
      ToolTipText     =   "Date Of Pick"
      Top             =   6000
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   13
      Left            =   1680
      TabIndex        =   92
      ToolTipText     =   "User Lot ID"
      Top             =   6000
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   13
      Left            =   120
      TabIndex        =   91
      ToolTipText     =   "System Lot ID"
      Top             =   6000
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   13
      Left            =   5280
      TabIndex        =   90
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   7680
      TabIndex        =   88
      ToolTipText     =   "Location"
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   4320
      TabIndex        =   87
      ToolTipText     =   "Date Of Pick"
      Top             =   5640
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   1680
      TabIndex        =   86
      ToolTipText     =   "User Lot ID"
      Top             =   5640
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   120
      TabIndex        =   85
      ToolTipText     =   "System Lot ID"
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   5280
      TabIndex        =   84
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   7680
      TabIndex        =   82
      ToolTipText     =   "Location"
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   4320
      TabIndex        =   81
      ToolTipText     =   "Date Of Pick"
      Top             =   5280
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   1680
      TabIndex        =   80
      ToolTipText     =   "User Lot ID"
      Top             =   5280
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   120
      TabIndex        =   79
      ToolTipText     =   "System Lot ID"
      Top             =   5280
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   5280
      TabIndex        =   78
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   4320
      TabIndex        =   76
      ToolTipText     =   "Date Of Pick"
      Top             =   4200
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   1680
      TabIndex        =   75
      ToolTipText     =   "User Lot ID"
      Top             =   4200
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   74
      ToolTipText     =   "System Lot ID"
      Top             =   4200
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   5280
      TabIndex        =   73
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   4320
      TabIndex        =   72
      ToolTipText     =   "Date Of Pick"
      Top             =   4560
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   1680
      TabIndex        =   71
      ToolTipText     =   "User Lot ID"
      Top             =   4560
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   70
      ToolTipText     =   "System Lot ID"
      Top             =   4560
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   5280
      TabIndex        =   69
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   5280
      TabIndex        =   68
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   67
      ToolTipText     =   "System Lot ID"
      Top             =   4920
      Width           =   1500
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   1680
      TabIndex        =   66
      ToolTipText     =   "User Lot ID"
      Top             =   4920
      Width           =   2580
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   4320
      TabIndex        =   65
      ToolTipText     =   "Date Of Pick"
      Top             =   4920
      Width           =   900
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   7680
      TabIndex        =   64
      ToolTipText     =   "Location"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   7680
      TabIndex        =   63
      ToolTipText     =   "Location"
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   7680
      TabIndex        =   62
      ToolTipText     =   "Location"
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lblSelected 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   5880
      TabIndex        =   61
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Quantity"
      Height          =   255
      Index           =   7
      Left            =   4320
      TabIndex        =   60
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblRestockQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   1920
      TabIndex        =   59
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4320
      TabIndex        =   58
      ToolTipText     =   "Date Of Pick"
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   57
      ToolTipText     =   "User Lot ID"
      Top             =   1680
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   56
      ToolTipText     =   "System Lot ID"
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   55
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   4320
      TabIndex        =   54
      ToolTipText     =   "Date Of Pick"
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   53
      ToolTipText     =   "User Lot ID"
      Top             =   2040
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   52
      ToolTipText     =   "System Lot ID"
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   51
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4320
      TabIndex        =   50
      ToolTipText     =   "Date Of Pick"
      Top             =   2400
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   49
      ToolTipText     =   "User Lot ID"
      Top             =   2400
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   48
      ToolTipText     =   "System Lot ID"
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   5280
      TabIndex        =   47
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   4320
      TabIndex        =   46
      ToolTipText     =   "Date Of Pick"
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   45
      ToolTipText     =   "User Lot ID"
      Top             =   2760
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   44
      ToolTipText     =   "System Lot ID"
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   5280
      TabIndex        =   43
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   4320
      TabIndex        =   42
      ToolTipText     =   "Date Of Pick"
      Top             =   3120
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   41
      ToolTipText     =   "User Lot ID"
      Top             =   3120
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   40
      ToolTipText     =   "System Lot ID"
      Top             =   3120
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   5280
      TabIndex        =   39
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   4320
      TabIndex        =   38
      ToolTipText     =   "Date Of Pick"
      Top             =   3480
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   1680
      TabIndex        =   37
      ToolTipText     =   "User Lot ID"
      Top             =   3480
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   36
      ToolTipText     =   "System Lot ID"
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   5280
      TabIndex        =   35
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   4320
      TabIndex        =   34
      ToolTipText     =   "Date Of Pick"
      Top             =   3840
      Width           =   900
   End
   Begin VB.Label lblLotu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   1680
      TabIndex        =   33
      ToolTipText     =   "User Lot ID"
      Top             =   3840
      Width           =   2580
   End
   Begin VB.Label lblLots 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   32
      ToolTipText     =   "System Lot ID"
      Top             =   3840
      Width           =   1500
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   5280
      TabIndex        =   31
      ToolTipText     =   "Click To Make Restock = Picked Quantity "
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   7680
      TabIndex        =   30
      ToolTipText     =   "Location"
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   7680
      TabIndex        =   29
      ToolTipText     =   "Location"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   7680
      TabIndex        =   28
      ToolTipText     =   "Location"
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   7680
      TabIndex        =   27
      ToolTipText     =   "Location"
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   7680
      TabIndex        =   26
      ToolTipText     =   "Location"
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   7680
      TabIndex        =   25
      ToolTipText     =   "Location"
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lblLotLoc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   7680
      TabIndex        =   24
      ToolTipText     =   "Location"
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Restock Quantity"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   23
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblRun 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4800
      TabIndex        =   21
      ToolTipText     =   "Manufacturing Order And Run"
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblMon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   20
      ToolTipText     =   "Manufacturing Order And Run"
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturing Order      "
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   240
      Width           =   1935
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
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   1455
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
      Left            =   1680
      TabIndex        =   17
      Top             =   1440
      Width           =   2535
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
      Left            =   4320
      TabIndex        =   16
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Picked Qty      "
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
      Left            =   5280
      TabIndex        =   15
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Restock Qty        "
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
      Left            =   6480
      TabIndex        =   14
      Top             =   1440
      Width           =   1335
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
      Left            =   7680
      TabIndex        =   13
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lots For P/N              "
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblPart 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      ToolTipText     =   "Picked Part Number"
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "PickMCe03b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 7/8/05
Option Explicit
Dim bOnLoad As Byte
Dim bSaved As Byte

Dim iTotalParts As Integer
Dim iCurrIdx As Integer

Dim sPartLotItems(301, 6) As String

Dim iTotalItems As Integer

Private Sub cmdCan_Click()
   If bSaved = 0 Then
      'MsgBox "Please Save The Information Before Exiting.", _
      '   vbInformation
      Select Case MsgBox("Cancel this restock transaction?", vbYesNo + vbQuestion)
      Case vbYes
         Es_LotSelectionCanceled = True
         Unload Me
      Case vbNo
      End Select
     
   Else
      Unload Me
   End If
   
End Sub



Private Sub cmdNew_Click()
   If lblSelected.ForeColor = ES_RED Then
      MsgBox "The Restock And Lot Select Quantities Do Not Match.", _
         vbInformation, Caption
   Else
      UpdateLots
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      Es_TotalLots = 0
      Es_LotSelectionCanceled = False
      GetPickedLots
      cmdLst.Enabled = True
      cmdNxt.Enabled = True
      bOnLoad = 0
      MouseCursor 0
      If iTotalItems = 0 Then
         Select Case MsgBox("No lots to select from.  Shall I create one?", vbYesNo + vbQuestion)
         Case vbNo
            Es_LotSelectionCanceled = True
         End Select
         Unload Me
      End If
   End If
   
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
   Set PickMCe03b = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   ' b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   For b = 1 To 15
      txtQty(b).BackColor = Me.BackColor
   Next
   'txtQty(b).BackColor = Me.BackColor
   lblSelected.ForeColor = ES_RED
   ManageBoxes
End Sub


Private Sub lblQty_Click(Index As Integer)
   txtQty(Index) = lblQty(Index)
   CheckSelected
   
End Sub

Private Sub txtQty_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtQty_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtQty_LostFocus(Index As Integer)
   On Error Resume Next
   txtQty(Index) = CheckLen(txtQty(Index), 9)
   txtQty(Index) = Format(Abs(Val(txtQty(Index))), ES_QuantityDataFormat)
   If Val(txtQty(Index)) > Val(lblQty(Index)) Then txtQty(Index) = lblQty(Index)
   sPartLotItems(Index + iCurrIdx, 5) = txtQty(Index)
   CheckSelected
   
   'sPartLotItems(Index + iCurrIdx, 5) = txtQty(Index)
   
End Sub

Private Sub cmdNxt_Click()
   iCurrIdx = iCurrIdx + 15
   cmdLst.Enabled = True
   If iCurrIdx > (iTotalItems - 15) Then cmdNxt.Enabled = False
   GetNextGroup
   
End Sub

Private Sub cmdLst_Click()
   iCurrIdx = iCurrIdx - 15
   If iCurrIdx < 0 Then iCurrIdx = 0
   If iCurrIdx = 0 Then cmdLst.Enabled = False
   If iCurrIdx >= 0 Then cmdNxt.Enabled = True
   GetNextGroup
   
End Sub

Private Sub ManageBoxes()
   On Error Resume Next
   Dim iRow As Integer
   For iRow = 1 To 15
      lblLots(iRow) = ""
      lblLotu(iRow) = ""
      lblLotLoc(iRow) = ""
      lblQty(iRow) = ""
      lblDate(iRow) = ""
      txtQty(iRow) = "0.000"
      txtQty(iRow).Enabled = True
      txtQty(iRow).BackColor = vbWhite
   Next
   
End Sub

Private Sub GetNextGroup()
   Dim iRow As Integer
   Dim iEnd As Integer
   ManageBoxes
   On Error Resume Next
   For iRow = 1 To 15
      If iRow + iCurrIdx > iTotalItems Then Exit For
      lblLots(iRow) = sPartLotItems(iRow + iCurrIdx, 0)
      lblLotu(iRow) = sPartLotItems(iRow + iCurrIdx, 1)
      lblLotLoc(iRow) = sPartLotItems(iRow + iCurrIdx, 2)
      lblQty(iRow) = sPartLotItems(iRow + iCurrIdx, 3)
      lblDate(iRow) = sPartLotItems(iRow + iCurrIdx, 4)
      txtQty(iRow) = sPartLotItems(iRow + iCurrIdx, 5)
      txtQty(iRow).Enabled = True
      txtQty(iRow).BackColor = vbWhite
      
   Next
   'If cmbAbc(1).Enabled Then cmbAbc(1).SetFocus

End Sub


Private Sub GetPickedLots()
   Dim RdoPick As ADODB.Recordset
   Dim bByte As Byte
   Dim cReStock As Currency
   On Error GoTo DiaErr1
   Es_LotsSelected = 0
   
'   sSql = "select LOTNUMBER,LOTUSERLOTID,LOTREMAININGQTY,LOTLOCATION,LOINUMBER," & vbCrLf _
'          & "LOITYPE,LOIPARTREF,LOIQUANTITY,LOIPDATE FROM LohdTable,LoitTable " & vbCrLf _
'          & "WHERE (LOTNUMBER=LOINUMBER AND LOIPARTREF='" & Compress(lblPart) & "' " & vbCrLf _
'          & "AND LOITYPE=10 AND LOIMOPARTREF='" & Compress(lblMon) & "' " & vbCrLf _
'          & "AND LOIMORUNNO=" & Val(lblRun) & ")"
   
'   sSql = "select LOTUSERLOTID, LOTLOCATION, LOIPDATE, x.*" & vbCrLf _
'      & "from LoitTable loit " & vbCrLf _
'      & "join LohdTable lohd on LOINUMBER = LOTNUMBER and LOITYPE = 10" & vbCrLf _
'      & "join " & vbCrLf _
'      & "(select LOINUMBER as LOTNUMBER, - sum(LOIQUANTITY) as ReturnQtyAvail" & vbCrLf _
'      & "from Loittable" & vbCrLf _
'      & "where LOIMOPARTREF='" & Compress(lblMon) & "' and LOIMORUNNO=" & lblRun & vbCrLf _
'      & "and LOIPARTREF='" & Compress(lblPart) & "'" & vbCrLf _
'      & "and LOITYPE in (10,21)" & vbCrLf _
'      & "group by LOINUMBER having sum(LOIQUANTITY) < 0 )" & vbCrLf _
'      & "as x" & vbCrLf _
'      & "on x.LOTNUMBER = lohd.LOTNUMBER" & vbCrLf
'
'
   Dim lot As New ClassLot
   Set RdoPick = lot.GetPickRestockLots(lblMon, CLng(lblRun), lblPart)
   If Not RdoPick.EOF Then
      With RdoPick
         Do Until .EOF
            bByte = bByte + 1
            ' 1/27/2008 - No need to restrict 10 items.
            'If bByte > 10 Then Exit Do
            'cReStock = GetSumRetLotQty(!LotNumber, Abs(!LOIQUANTITY))
            If bByte < 16 Then
               cReStock = !ReturnQtyAvail
               lblLots(bByte) = "" & Trim(!lotNumber)
               lblLotu(bByte) = "" & Trim(!LOTUSERLOTID)
               lblLotLoc(bByte) = "" & Trim(!LOTLOCATION)
               lblQty(bByte) = Format(cReStock, ES_QuantityDataFormat)
               lblDate(bByte) = Format(!LOIPDATE, "mm/dd/yy")
               txtQty(bByte) = "0.000"
               txtQty(bByte).Enabled = True
               txtQty(bByte).BackColor = vbWhite
            End If
            
            sPartLotItems(bByte, 0) = "" & Trim(!lotNumber)
            sPartLotItems(bByte, 1) = "" & Trim(!LOTUSERLOTID)
            sPartLotItems(bByte, 2) = "" & Trim(!LOTLOCATION)
            sPartLotItems(bByte, 3) = Format(cReStock, ES_QuantityDataFormat)
            sPartLotItems(bByte, 4) = Format(!LOIPDATE, "mm/dd/yy")
            sPartLotItems(bByte, 5) = "0.000"
            
            .MoveNext
         Loop
         ClearResultSet RdoPick
      End With
   End If
   iTotalItems = bByte
   On Error Resume Next
   Set RdoPick = Nothing
   If bByte > 0 Then txtQty(1).SetFocus
   Exit Sub
DiaErr1:
   Set RdoPick = Nothing
   sProcName = "getpickedlots"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub CheckSelected()
   Dim bByte As Byte
   Dim cSelected As Currency
   On Error Resume Next
   For bByte = 1 To iTotalItems
      cSelected = cSelected + Val(sPartLotItems(bByte, 5))   'txtQty(bByte)
   Next
   'cSelected = cSelected + Val(txtQty(bByte))
   lblSelected = Format(cSelected, ES_QuantityDataFormat)
   If lblSelected = lblRestockQty Then
      lblSelected.ForeColor = ES_BLUE
   Else
      lblSelected.ForeColor = ES_RED
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

Private Sub UpdateLots()
   Dim iRow As Integer
   Dim iList As Integer
   
   'txtQty(iRow)
   On Error GoTo DiaErr1
   Erase lots
   For iRow = 1 To iTotalItems
      If Val(sPartLotItems(iRow, 5)) > 0 Then
         iList = iList + 1
         Es_TotalLots = Es_TotalLots + 1
         ReDim Preserve lots(iList)
         With lots(iList)
            .LotSysId = sPartLotItems(iRow, 0) 'lblLots(iRow)
            .LotUserId = sPartLotItems(iRow, 1) 'lblLotu(iRow)
            .LotPartRef = Compress(lblPart)
            .LotADate = Format(Now, "mm/dd/yy")
            .LotSelQty = Format(Val(sPartLotItems(iRow, 5)), ES_QuantityDataFormat)
         End With
      End If
   Next
   bSaved = 1
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

Private Function GetSumRetLotQty(sLotNumber As String, cLotQty) As Currency
   Dim RdoQty As ADODB.Recordset
   sSql = "SELECT SUM(LOIQUANTITY) As LotQty FROM LoitTable WHERE (LOIPARTREF='" _
          & Compress(lblPart) & "' AND LOINUMBER='" & sLotNumber & "'AND LOITYPE=21 AND " _
          & "LOIMOPARTREF='" & Compress(lblMon) & "' AND LOIMORUNNO=" & Val(lblRun) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoQty, ES_FORWARD)
   If bSqlRows Then
      With RdoQty
         If Not IsNull(!lotQty) Then
            GetSumRetLotQty = !lotQty
         Else
            GetSumRetLotQty = cLotQty
         End If
      End With
   End If
   Set RdoQty = Nothing
End Function
