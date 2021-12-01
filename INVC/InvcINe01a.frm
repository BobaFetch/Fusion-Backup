VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form InvcINe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parts"
   ClientHeight    =   5865
   ClientLeft      =   1095
   ClientTop       =   540
   ClientWidth     =   7080
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   HelpContextID   =   5101
   Icon            =   "InvcINe01a.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5865
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin VB.Frame tabFrame 
      Height          =   3972
      Index           =   1
      Left            =   2280
      TabIndex        =   76
      Top             =   3360
      Visible         =   0   'False
      Width           =   6732
      Begin VB.CheckBox chkEAR 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   5400
         TabIndex        =   146
         Top             =   2985
         Width           =   615
      End
      Begin VB.CheckBox chkITAR 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   5400
         TabIndex        =   145
         Top             =   2640
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdateLotLocations 
         Caption         =   "Update Lot Locations"
         Height          =   315
         Left            =   3420
         TabIndex        =   13
         Top             =   600
         Width           =   1995
      End
      Begin VB.CheckBox chkInactive 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   5400
         TabIndex        =   22
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox chkObsolete 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   2685
         Width           =   615
      End
      Begin VB.CheckBox chkLotsExpire 
         Caption         =   "___"
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   2340
         Width           =   615
      End
      Begin VB.CheckBox optTaxEx 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   5400
         TabIndex        =   20
         Top             =   1992
         Width           =   615
      End
      Begin VB.ComboBox cmbAbc 
         ForeColor       =   &H00800000&
         Height          =   288
         ItemData        =   "InvcINe01a.frx":030A
         Left            =   2040
         List            =   "InvcINe01a.frx":031A
         TabIndex        =   14
         Tag             =   "3"
         ToolTipText     =   "Select From List Or Leave Blank"
         Top             =   912
         Width           =   615
      End
      Begin VB.TextBox txtPun 
         Height          =   285
         Left            =   5400
         TabIndex        =   16
         Tag             =   "3"
         ToolTipText     =   "Vendor Part Unit Of Measure"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox optDte 
         Caption         =   "___"
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   5760
         TabIndex        =   25
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox optLot 
         Caption         =   "___"
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   1992
         Width           =   615
      End
      Begin VB.TextBox txtPrc 
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Tag             =   "1"
         Text            =   "1.0000"
         ToolTipText     =   "PO Quantity/Conversion = Purchasing Factor"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox optCom 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   5400
         TabIndex        =   18
         Top             =   1680
         Width           =   615
      End
      Begin VB.CheckBox optIns 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   1680
         Width           =   615
      End
      Begin VB.CheckBox optBom 
         Caption         =   "___"
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   3312
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtWhs 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         TabIndex        =   11
         Tag             =   "3"
         Text            =   "NONE"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtLoc 
         Height          =   285
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   12
         Tag             =   "3"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtWht 
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Tag             =   "1"
         Text            =   "1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "EAR"
         Height          =   285
         Index           =   53
         Left            =   3480
         TabIndex        =   148
         Top             =   2985
         Width           =   1815
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "ITAR"
         Height          =   285
         Index           =   57
         Left            =   3480
         TabIndex        =   147
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inactive Part?"
         Height          =   285
         Index           =   56
         Left            =   3480
         TabIndex        =   140
         Top             =   2340
         Width           =   1815
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Obsolete Part?"
         Height          =   285
         Index           =   55
         Left            =   240
         TabIndex        =   139
         Top             =   2685
         Width           =   1695
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Do Lots Expire?"
         Height          =   285
         Index           =   54
         Left            =   240
         TabIndex        =   138
         Top             =   2340
         Width           =   1695
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Exempt Item?"
         Height          =   288
         Index           =   51
         Left            =   3480
         TabIndex        =   88
         Top             =   1992
         Width           =   1812
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchasing Unit of Meas"
         Height          =   288
         Index           =   28
         Left            =   3480
         TabIndex        =   87
         Top             =   1320
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "ABC Class"
         Height          =   288
         Index           =   22
         Left            =   240
         TabIndex        =   86
         Top             =   960
         Width           =   1512
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Track Lots By Expiration Date?"
         Height          =   288
         Index           =   20
         Left            =   3360
         TabIndex        =   85
         Top             =   3360
         Visible         =   0   'False
         Width           =   2412
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Tracking Req'd?"
         Height          =   288
         Index           =   19
         Left            =   240
         TabIndex        =   84
         Top             =   1992
         Width           =   1812
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Commission?"
         Height          =   288
         Index           =   17
         Left            =   3480
         TabIndex        =   83
         Top             =   1680
         Width           =   2472
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phantom On Bom?"
         Height          =   288
         Index           =   16
         Left            =   240
         TabIndex        =   82
         Top             =   3312
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchasing Conversion"
         Height          =   288
         Index           =   13
         Left            =   240
         TabIndex        =   81
         Top             =   1320
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Requires On Dock Insp"
         Height          =   288
         Index           =   12
         Left            =   240
         TabIndex        =   80
         Top             =   1680
         Width           =   1752
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Warehouse"
         Height          =   285
         Index           =   11
         Left            =   3480
         TabIndex        =   79
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   288
         Index           =   10
         Left            =   240
         TabIndex        =   78
         Top             =   600
         Width           =   1512
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Weight"
         Height          =   285
         Index           =   9
         Left            =   240
         TabIndex        =   77
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   3972
      Index           =   3
      Left            =   3120
      TabIndex        =   96
      Top             =   7200
      Visible         =   0   'False
      Width           =   6732
      Begin VB.TextBox txtCustFld1 
         Height          =   285
         Left            =   1920
         TabIndex        =   152
         Tag             =   "2"
         ToolTipText     =   "Process Color"
         Top             =   2880
         Width           =   2985
      End
      Begin VB.TextBox txtColor 
         Height          =   285
         Left            =   1920
         TabIndex        =   40
         Tag             =   "2"
         ToolTipText     =   "Process Color"
         Top             =   2472
         Width           =   2985
      End
      Begin VB.TextBox txtMatType 
         Height          =   285
         Left            =   1920
         TabIndex        =   38
         Tag             =   "2"
         ToolTipText     =   "Material Type"
         Top             =   1752
         Width           =   2985
      End
      Begin VB.TextBox txtMatAlloy 
         Height          =   285
         Left            =   1920
         TabIndex        =   39
         Tag             =   "3"
         ToolTipText     =   "Material Alloy"
         Top             =   2112
         Width           =   2985
      End
      Begin VB.TextBox txtSsq 
         Height          =   285
         Left            =   5280
         TabIndex        =   35
         Tag             =   "1"
         Top             =   672
         Width           =   1095
      End
      Begin VB.TextBox txtOst 
         Height          =   285
         Left            =   5280
         TabIndex        =   33
         Tag             =   "1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtPlt 
         Height          =   285
         Left            =   5280
         TabIndex        =   37
         Tag             =   "1"
         Top             =   1152
         Width           =   855
      End
      Begin VB.TextBox txtMft 
         Height          =   285
         Left            =   1920
         TabIndex        =   36
         Tag             =   "1"
         Top             =   1152
         Width           =   855
      End
      Begin VB.TextBox txtEoq 
         Height          =   285
         Left            =   1920
         TabIndex        =   34
         Tag             =   "1"
         Top             =   672
         Width           =   1095
      End
      Begin VB.TextBox txtSfs 
         Height          =   285
         Left            =   1920
         TabIndex        =   32
         Tag             =   "1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblCustFlds1 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   153
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Height          =   252
         Left            =   120
         TabIndex        =   105
         Top             =   2472
         Width           =   1272
      End
      Begin VB.Label lblMatType 
         BackStyle       =   0  'Transparent
         Height          =   252
         Left            =   120
         TabIndex        =   104
         Top             =   1752
         Width           =   1152
      End
      Begin VB.Label lblMatAlloy 
         BackStyle       =   0  'Transparent
         Height          =   252
         Left            =   120
         TabIndex        =   103
         Top             =   2112
         Width           =   1272
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship Set Quantity"
         Height          =   288
         Index           =   35
         Left            =   3360
         TabIndex        =   102
         Top             =   672
         Width           =   1572
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Overstock (Maximum)"
         Height          =   288
         Index           =   23
         Left            =   3360
         TabIndex        =   101
         Top             =   240
         Width           =   1572
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchasing Lead Time (Days)"
         Height          =   408
         Index           =   32
         Left            =   3360
         TabIndex        =   100
         Top             =   1152
         Width           =   2052
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacturing Flow Time (Days)"
         Height          =   408
         Index           =   31
         Left            =   120
         TabIndex        =   99
         Top             =   1152
         Width           =   1572
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Economic Order Qty"
         Height          =   288
         Index           =   30
         Left            =   120
         TabIndex        =   98
         Top             =   672
         Width           =   1692
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Safety Stock (Min)"
         Height          =   288
         Index           =   29
         Left            =   120
         TabIndex        =   97
         Top             =   240
         Width           =   1572
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   3972
      Index           =   5
      Left            =   360
      TabIndex        =   117
      Top             =   4320
      Visible         =   0   'False
      Width           =   6732
      Begin VB.CommandButton cmdVew 
         DownPicture     =   "InvcINe01a.frx":032C
         Height          =   350
         Index           =   1
         Left            =   6240
         Picture         =   "InvcINe01a.frx":0806
         Style           =   1  'Graphical
         TabIndex        =   118
         TabStop         =   0   'False
         ToolTipText     =   "Lookup Accounts"
         Top             =   120
         Width           =   350
      End
      Begin VB.ComboBox txtGex 
         Height          =   288
         Left            =   1920
         TabIndex        =   53
         Tag             =   "3"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.ComboBox txtGoh 
         Height          =   288
         Left            =   1920
         TabIndex        =   52
         Tag             =   "3"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.ComboBox txtGma 
         Height          =   288
         Left            =   1920
         TabIndex        =   51
         Tag             =   "3"
         Top             =   2640
         Width           =   1935
      End
      Begin VB.ComboBox txtGla 
         Height          =   288
         Left            =   1920
         TabIndex        =   50
         Tag             =   "3"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.ComboBox txtWex 
         Height          =   288
         Left            =   1920
         TabIndex        =   49
         Tag             =   "3"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox txtWoh 
         Height          =   288
         Left            =   1920
         TabIndex        =   48
         Tag             =   "3"
         Top             =   1296
         Width           =   1935
      End
      Begin VB.ComboBox txtWma 
         Height          =   288
         Left            =   1920
         TabIndex        =   47
         Tag             =   "3"
         Top             =   924
         Width           =   1935
      End
      Begin VB.ComboBox txtWla 
         Height          =   288
         Left            =   1920
         TabIndex        =   46
         Tag             =   "3"
         Top             =   528
         Width           =   1935
      End
      Begin VB.Label lblGex 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   136
         Top             =   3360
         Width           =   2400
      End
      Begin VB.Label lblGoh 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   135
         Top             =   3000
         Width           =   2400
      End
      Begin VB.Label lblGma 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   134
         Top             =   2640
         Width           =   2400
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Account"
         Height          =   252
         Index           =   50
         Left            =   240
         TabIndex        =   133
         Top             =   2280
         Width           =   1800
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material Account"
         Height          =   252
         Index           =   49
         Left            =   240
         TabIndex        =   132
         Top             =   2640
         Width           =   1800
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Overhead Account"
         Height          =   252
         Index           =   48
         Left            =   240
         TabIndex        =   131
         Top             =   3000
         Width           =   1800
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Account"
         Height          =   252
         Index           =   47
         Left            =   240
         TabIndex        =   130
         Top             =   3360
         Width           =   1800
      End
      Begin VB.Label lblGla 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   129
         Top             =   2280
         Width           =   2400
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Of Goods Sold Accounts:"
         Height          =   252
         Index           =   46
         Left            =   120
         TabIndex        =   128
         Top             =   2040
         Width           =   3072
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory/Expense Accounts:"
         Height          =   255
         Index           =   39
         Left            =   240
         TabIndex        =   127
         Top             =   240
         Width           =   3075
      End
      Begin VB.Label lblWex 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   126
         Top             =   1680
         Width           =   2400
      End
      Begin VB.Label lblWoh 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   125
         Top             =   1296
         Width           =   2400
      End
      Begin VB.Label lblWma 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   124
         Top             =   924
         Width           =   2400
      End
      Begin VB.Label lblWla 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   123
         Top             =   528
         Width           =   2400
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Account"
         Height          =   252
         Index           =   42
         Left            =   240
         TabIndex        =   122
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Overhead Account"
         Height          =   252
         Index           =   43
         Left            =   240
         TabIndex        =   121
         Top             =   1296
         Width           =   1800
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material Account"
         Height          =   252
         Index           =   44
         Left            =   240
         TabIndex        =   120
         Top             =   924
         Width           =   1800
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Account"
         Height          =   252
         Index           =   45
         Left            =   240
         TabIndex        =   119
         Top             =   600
         Width           =   1800
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   3972
      Index           =   4
      Left            =   5160
      TabIndex        =   106
      Top             =   1440
      Visible         =   0   'False
      Width           =   6732
      Begin VB.ComboBox txtRev 
         Height          =   288
         Left            =   1920
         TabIndex        =   41
         Top             =   648
         Width           =   1935
      End
      Begin VB.ComboBox txtDis 
         Height          =   288
         Left            =   1920
         TabIndex        =   42
         Top             =   1032
         Width           =   1935
      End
      Begin VB.ComboBox txtTrv 
         Height          =   288
         Left            =   1920
         TabIndex        =   43
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox txtTcg 
         Height          =   288
         Left            =   1920
         TabIndex        =   44
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmdVew 
         DownPicture     =   "InvcINe01a.frx":0CE0
         Height          =   350
         Index           =   0
         Left            =   6240
         Picture         =   "InvcINe01a.frx":11BA
         Style           =   1  'Graphical
         TabIndex        =   107
         TabStop         =   0   'False
         ToolTipText     =   "Lookup Accounts"
         Top             =   240
         Width           =   350
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "General Accounts:"
         Height          =   252
         Index           =   38
         Left            =   120
         TabIndex        =   116
         Top             =   360
         Width           =   3072
      End
      Begin VB.Label lblRev 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   115
         Top             =   648
         Width           =   2400
      End
      Begin VB.Label lblDis 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   114
         Top             =   1032
         Width           =   2400
      End
      Begin VB.Label lblTrv 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   113
         Top             =   1440
         Width           =   2400
      End
      Begin VB.Label lblTcg 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4080
         TabIndex        =   112
         Top             =   1800
         Width           =   2400
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Revenue Account"
         Height          =   252
         Index           =   41
         Left            =   120
         TabIndex        =   111
         Top             =   648
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Account"
         Height          =   252
         Index           =   40
         Left            =   120
         TabIndex        =   110
         Top             =   1032
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer Revenue Acct"
         Height          =   252
         Index           =   37
         Left            =   120
         TabIndex        =   109
         Top             =   1440
         Width           =   1992
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer CGS Acct"
         Height          =   252
         Index           =   36
         Left            =   120
         TabIndex        =   108
         Top             =   1800
         Width           =   1992
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   3972
      Index           =   2
      Left            =   4200
      TabIndex        =   89
      Top             =   7080
      Visible         =   0   'False
      Width           =   6732
      Begin VB.TextBox txtQoh 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   142
         Tag             =   "1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame Fr1 
         Height          =   495
         Left            =   1920
         TabIndex        =   90
         Top             =   1320
         Width           =   3015
         Begin VB.OptionButton cstStd 
            Caption         =   "Standard"
            Height          =   255
            Left            =   1200
            TabIndex        =   29
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton cstAve 
            Caption         =   "Average"
            Height          =   255
            Left            =   2400
            TabIndex        =   30
            Top             =   120
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.OptionButton cstAct 
            Caption         =   "Actual"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.TextBox txtSpr 
         Height          =   285
         Left            =   1920
         TabIndex        =   27
         Tag             =   "1"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtStd 
         Height          =   285
         Left            =   1920
         TabIndex        =   26
         Tag             =   "1"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label txtPrice 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   151
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label txtDate 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1920
         TabIndex        =   150
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label z1 
         Caption         =   "Last Cost"
         Height          =   255
         Index           =   34
         Left            =   3360
         TabIndex        =   149
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label z1 
         Caption         =   "Last Purchased Date"
         Height          =   255
         Index           =   33
         Left            =   120
         TabIndex        =   144
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity on hand"
         Height          =   285
         Index           =   27
         Left            =   120
         TabIndex        =   143
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Costing"
         Height          =   288
         Index           =   21
         Left            =   120
         TabIndex        =   95
         Top             =   1440
         Width           =   1212
      End
      Begin VB.Label txtAve 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4800
         TabIndex        =   94
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Price"
         Height          =   288
         Index           =   26
         Left            =   120
         TabIndex        =   93
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Average Cost"
         Height          =   288
         Index           =   25
         Left            =   3360
         TabIndex        =   92
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Cost"
         Height          =   288
         Index           =   24
         Left            =   120
         TabIndex        =   91
         Top             =   600
         Width           =   1152
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   3972
      Index           =   0
      Left            =   40
      TabIndex        =   63
      Top             =   1560
      Width           =   6816
      Begin VB.CheckBox optTol 
         Caption         =   "___"
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   3312
         Width           =   615
      End
      Begin VB.ComboBox cmbCde 
         Height          =   288
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   6
         Tag             =   "3"
         ToolTipText     =   "Select Product Code From List"
         Top             =   1032
         Width           =   1215
      End
      Begin VB.ComboBox cmbCls 
         Height          =   288
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   7
         Tag             =   "3"
         ToolTipText     =   "Select Product Class From List"
         Top             =   1392
         Width           =   855
      End
      Begin VB.ComboBox cmbTyp 
         Height          =   288
         ItemData        =   "InvcINe01a.frx":1694
         Left            =   1920
         List            =   "InvcINe01a.frx":16B0
         TabIndex        =   2
         Tag             =   "1"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cmbMbe 
         Height          =   288
         ItemData        =   "InvcINe01a.frx":16CC
         Left            =   5280
         List            =   "InvcINe01a.frx":16D9
         TabIndex        =   3
         Tag             =   "8"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtUom 
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "3"
         Top             =   672
         Width           =   375
      End
      Begin VB.TextBox txtComt 
         Height          =   1245
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Tag             =   "9"
         ToolTipText     =   "Comments (3072 Chars Max)"
         Top             =   1992
         Width           =   4200
      End
      Begin VB.TextBox txtRrq 
         Height          =   285
         Left            =   5280
         TabIndex        =   5
         Tag             =   "1"
         Text            =   "0"
         Top             =   672
         Width           =   855
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "This Part Is A Tool"
         Height          =   288
         Index           =   18
         Left            =   120
         TabIndex        =   75
         ToolTipText     =   "Tools Are Created In The Tooling Group"
         Top             =   3312
         Width           =   1812
      End
      Begin VB.Label lblPClass 
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   3360
         TabIndex        =   74
         Top             =   1392
         Width           =   2412
      End
      Begin VB.Label lblPCode 
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   3360
         TabIndex        =   73
         Top             =   1032
         Width           =   2412
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
         Height          =   288
         Index           =   14
         Left            =   120
         TabIndex        =   72
         Top             =   1080
         Width           =   1512
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Class"
         Height          =   288
         Index           =   15
         Left            =   120
         TabIndex        =   71
         Top             =   1392
         Width           =   1512
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(M,B or E)"
         Height          =   288
         Index           =   8
         Left            =   6000
         TabIndex        =   70
         Top             =   240
         Width           =   828
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Make, Buy or Either"
         Height          =   288
         Index           =   7
         Left            =   3360
         TabIndex        =   69
         Top             =   240
         Width           =   1872
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Type"
         Height          =   288
         Index           =   2
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   " (1-8)"
         Height          =   288
         Index           =   3
         Left            =   2640
         TabIndex        =   67
         Top             =   240
         Width           =   828
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit of Meas"
         Height          =   288
         Index           =   4
         Left            =   120
         TabIndex        =   66
         Top             =   672
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Extended Description:"
         Height          =   192
         Index           =   5
         Left            =   120
         TabIndex        =   65
         Top             =   2040
         Width           =   1632
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Recommended Run Qty"
         Height          =   288
         Index           =   6
         Left            =   3360
         TabIndex        =   64
         Top             =   672
         Width           =   1872
      End
   End
   Begin VB.Frame z2 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   61
      Top             =   1080
      Width           =   6972
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InvcINe01a.frx":16E6
      Style           =   1  'Graphical
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdFnd 
      Height          =   315
      Left            =   4440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "InvcINe01a.frx":1E94
      Style           =   1  'Graphical
      TabIndex        =   59
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   360
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   58
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Leading Char Search  (*  In Front Is A Legal Wild Card)"
      Top             =   360
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.TextBox txtDoc 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "0"
      ToolTipText     =   "Document List Revision. NONE Means That No List Has Been Assigned"
      Top             =   720
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Index           =   0
      Interval        =   6000
      Left            =   6180
      Top             =   6120
   End
   Begin VB.CheckBox optVewAcct 
      Caption         =   "Just for acct lookup"
      Height          =   255
      Left            =   1320
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      CausesValidation=   0   'False
      Height          =   435
      Left            =   6000
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "2"
      ToolTipText     =   "Part Description (30 chars)"
      Top             =   750
      Width           =   2985
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter a New Part or Select From List (30 chars)"
      Top             =   360
      Width           =   3255
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   6000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5865
      FormDesignWidth =   7080
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   4572
      Left            =   0
      TabIndex        =   62
      Top             =   1200
      Width           =   6972
      _ExtentX        =   12303
      _ExtentY        =   8070
      TabWidthStyle   =   2
      TabFixedWidth   =   2028
      TabFixedHeight  =   473
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Setup"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Costing"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Manufacturing"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Be sure to sendTabStrip to back when done editing."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7380
      TabIndex        =   141
      Top             =   1080
      Width           =   4755
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "More >>>"
      Height          =   252
      Left            =   6120
      TabIndex        =   137
      Top             =   480
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doc List Revision "
      Height          =   285
      Index           =   52
      Left            =   4560
      TabIndex        =   57
      ToolTipText     =   "Document List Revision. NONE Means That No List Has Been Assigned"
      Top             =   750
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   55
      Top             =   750
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add/Revise Part Number"
      Height          =   375
      Index           =   0
      Left            =   90
      TabIndex        =   54
      Top             =   300
      Width           =   1155
   End
End
Attribute VB_Name = "InvcINe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
Option Explicit

Dim AdoEdit As ADODB.Recordset

Dim iTimer As Integer

Dim bActFilled As Byte
Dim bCanceled As Byte
Dim bFirstPart As Byte
Dim bFrom As Byte
Dim bGoodPart As Byte
Dim bOnLoad As Byte
Dim bNewPart As Byte
Dim bNoAccts As Byte
Dim bShowParts As Byte
Dim bOpeningPart As Byte


Dim iPalevel As Integer
Dim iPartCount As Integer
Dim cOldCost As Currency

Dim sNewPartNo As String
Dim sDebitAcct As String
Dim sCreditAcct As String
Dim sLotNumber As String
Dim sTimeStamp1 As String
Dim sTimeStamp2 As String
'Arrays
Dim cPrices(7, 4) As Currency

'Change On Dock
Dim sODPart(300) As String
Dim bODReqd(300) As Byte
Dim bODChng(300) As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtDropDown() As New EsiKeyBd
Private txtChange() As New EsiKeyBd

Private Function GetProductClass()
   On Error Resume Next
   Dim RdoCode As ADODB.Recordset
   
   sSql = "SELECT CCREF,CCDESC FROM PclsTable WHERE " _
          & "CCREF='" & Compress(cmbCls) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCode, ES_FORWARD)
   If bSqlRows Then GetProductClass = "" & Trim(RdoCode!CCDESC) _
                                      Else GetProductClass = ""
   Set RdoCode = Nothing
   
End Function

Private Function GetProductCode() As String
   On Error Resume Next
   Dim RdoCode As ADODB.Recordset
   
   sSql = "SELECT PCREF,PCDESC FROM PcodTable WHERE " _
          & "PCREF='" & Compress(cmbCde) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCode, ES_FORWARD)
   If bSqlRows Then GetProductCode = "" & Trim(RdoCode!PCDESC) _
                                     Else GetProductCode = ""
   Set RdoCode = Nothing
   
End Function


Private Sub FormatControls()
   Dim RdoSet As ADODB.Recordset
   Dim b As Byte
   
   On Error Resume Next
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   b = AutoFormatChange(Me, txtChange())
   cmbMbe = cmbMbe.List(0)
   cmbTyp = cmbTyp.List(0)
   cmbMbe.ForeColor = ES_BLUE
   cmbTyp.ForeColor = ES_BLUE
   txtDoc.BackColor = InvcINe01a.BackColor
   If ES_PARTCOUNT > 5000 Then
     ' cmbPrt.Visible = False
'      txtPrt.Visible = True
       cmbPrt.TabIndex = 0
 '     cmdFnd.Visible = True
   End If
   sSql = "Qry_GetABCPreference"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSet, ES_FORWARD)
   If bSqlRows Then
      With RdoSet
         If Not IsNull(.Fields(0)) Then
            b = .Fields(0)
         Else
            b = 0
         End If
         ClearResultSet RdoSet
      End With
   End If
   
   If b = 0 Then
      cmbAbc.ToolTipText = "Caution ABC Class Codes Have Not Been Properly Initialized. Recommend Leaving Blank."
   Else
      cmbAbc.ToolTipText = "ABC Class Codes Have Been Initialized. Cautiously Select From The List Or Leave Blank."
   End If
   optTol.Caption = ""
   Set RdoSet = Nothing
End Sub

Private Sub cmbAbc_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub cmbAbc_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbAbc = CheckLen(cmbAbc, 2)
   If Len(cmbAbc) > 0 Then
      For iList = 0 To cmbAbc.ListCount - 1
         If cmbAbc = cmbAbc.List(iList) Then
            bByte = 1
            Exit For
         End If
      Next
   Else
      bByte = 1
   End If
   If bByte = 0 Then
      'SysSysSysBeep
      cmbAbc = ""
   End If
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAABC = "" & cmbAbc
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
   
End Sub


Private Sub cmbCde_Click()
   lblPCode = GetProductCode()
   
End Sub

Private Sub cmbCde_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub cmbCde_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbCde = CheckLen(cmbCde, 6)
   If cmbCde.ListCount > 0 Then
      For iList = 0 To cmbCde.ListCount - 1
         If cmbCde = cmbCde.List(iList) Then b = 1
      Next
      If b = 0 And cmbCde <> "" Then
         'SysSysSysBeep
         cmbCde = cmbCde.List(0)
      End If
   Else
      cmbCde = "   "
   End If
   lblPCode = GetProductCode()
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAPRODCODE = "" & Compress(cmbCde)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub cmbCls_Change()
   If Len(cmbCls) > 4 Then cmbCls = Left(cmbCls, 4)
   
End Sub

Private Sub cmbCls_Click()
   lblPClass = GetProductClass()
   
End Sub

Private Sub cmbCls_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub cmbCls_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbCls = CheckLen(cmbCls, 4)
   If cmbCls.ListCount > 0 Then
      For iList = 0 To cmbCls.ListCount - 1
         If cmbCls = cmbCls.List(iList) Then b = 1
      Next
      If b = 0 And cmbCls <> "" Then
         'SysSysSysBeep
         cmbCls = cmbCls.List(0)
      End If
   Else
      cmbCls = "   "
   End If
   lblPClass = GetProductClass()
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PACLASS = "" & Compress(cmbCls)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub cmbMbe_Change()
   If Len(cmbMbe) > 1 Then cmbMbe = Left(cmbMbe, 1)
   
End Sub

Private Sub cmbMbe_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub cmbMbe_LostFocus()
   Dim bByte As Byte
   iTimer = 0
   cmbMbe = CheckLen(cmbMbe, 1)
   Select Case cmbMbe
      Case "M", "B", "E"
         bByte = 1
      Case Else
         bByte = 0
   End Select
   If bByte = 0 Then
      MsgBox "Must Be M, B or E.", vbInformation, Caption
      cmbMbe = "M"
   End If
   
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAMAKEBUY = "" & cmbMbe
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub cmbPrt_Click()
   'bGoodPart = GetPart(0)
   bGoodPart = GetPart(1)
   
End Sub

Private Sub cmbPrt_Change()
   'bGoodPart = GetPart(0)
   'bGoodPart = GetPart(1)
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bCanceled Then Exit Sub
   If Len(cmbPrt) = 0 Then
      cmdCan.SetFocus
      Exit Sub
   End If
   bGoodPart = GetPart(1)
   If bGoodPart = 0 Then
      Addpart
      'TODO
      bGoodPart = 1
      'bGoodPart = GetPart(1)
   Else
      sPassedPart = cmbPrt
   End If
   
End Sub

Private Sub cmbTyp_Change()
   If Len(cmbTyp) > 1 Then cmbTyp = Left(cmbTyp, 1)
   
End Sub

Private Sub cmbTyp_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub cmbTyp_Validate(Cancel As Boolean)
   If Me.cmdCan.Value Then
      Exit Sub
   End If
   cmbTyp = CheckLen(cmbTyp, 1)
   If Val(cmbTyp) < 1 Or Val(cmbTyp) > 8 Then
      'SysSysSysBeep
      cmbTyp = "3"
   End If
   
   'check that part type is valid for any containing assemblies and any subassemblies
   Dim part As New ClassPart, oldType As Integer
   oldType = part.IsPartTypeChangeOK(Me.cmbPrt, Val(cmbTyp))
   If oldType = -1 Then
   Else
      Cancel = True
      cmbTyp = oldType
      'cmbTyp.SetFocus
      Exit Sub
   End If
   
   iPalevel = Val(cmbTyp)
   If bGoodPart = 1 Then
      'On Error Resume Next
      AdoEdit!PALEVEL = Val(cmbTyp)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
   CheckPartType
End Sub

Private Sub cmdCan_Click()
   Timer1(0).Enabled = False
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = 1
   
End Sub


Private Sub cmdFnd_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   ViewParts.Show
   bShowParts = 0
   
End Sub

Private Sub cmdFnd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bShowParts = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 5101
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub cmdUpdateLotLocations_Click()
   Dim part As New ClassPart
   part.UpdatePartLotLocations cmbPrt, txtLoc, True
End Sub

Private Sub cmdVew_Click(Index As Integer)
   optVewAcct.Value = vbChecked
   VewAcct.Show
   
End Sub



Private Sub cstAct_Click()
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAUSEACTUALCOST = Abs(cstAct.Value)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub

Private Sub cstAve_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' UnderCons.Show 1
   
End Sub

Private Sub cstStd_Click()
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAUSEACTUALCOST = Abs(cstAct.Value)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   Dim ErrVal As Boolean
   Dim strLabelDesc As String
   MdiSect.lblBotPanel = Caption
   MouseCursor 13
   If optVewAcct.Value = vbChecked Then Unload VewAcct
   If bOnLoad Then
      bDataHasChanged = False
      b = CheckLotStatus
      If b = 1 Then
         optLot.Enabled = True
         chkLotsExpire.Enabled = True
      End If
      b = GetTimeStamp()
      FillProductCodes
      FillProductClasses
      
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillPartCombo
      
      cmbPrt = ""
      bOnLoad = 0
      FillCombo
      bGoodPart = 0
      bDataHasChanged = False
      GetLabelName "MaterialID", strLabelDesc, ErrVal
      
      If ErrVal Then
        lblMatType = strLabelDesc
        strLabelDesc = ""
        
        GetLabelName "AlloyID", strLabelDesc, ErrVal
        lblMatAlloy = strLabelDesc
        strLabelDesc = ""
      
        GetLabelName "ColorID", strLabelDesc, ErrVal
        lblColor = strLabelDesc
        strLabelDesc = ""
        
        GetLabelName "CustFld1", strLabelDesc, ErrVal
        lblCustFlds1 = strLabelDesc
        strLabelDesc = ""
      Else
        lblMatType = "Material"
        lblMatAlloy = "Alloy/Temper"
        lblColor = "Color"
      End If
      
   End If
   
End Sub

Private Sub Form_Deactivate()
   If bFrom > 0 Then Unload Me
   
End Sub

Private Sub Form_Load()
   Dim b As Byte
   
   bDataHasChanged = False
   
   For b = 0 To Forms.count - 1
      If Forms(b).Name = "BompBMe01a" Then bFrom = b
   Next
   If bFrom = 0 Then
      FormLoad Me
      
   Else
      FormLoad Me, ES_DONTLIST
      Move 1000, 500
   End If
   For b = 0 To 5
      With tabFrame(b)
         .BorderStyle = 0
         .Left = 40
         .Visible = False
      End With
   Next
   tabFrame(0).Visible = True
   FormatControls
   cUR.CurrentPart = GetSetting("Esi2000", "Current", "Part", cUR.CurrentPart)
   bFirstPart = 1
   bOnLoad = 1
   
End Sub

Private Sub FillPartCombo()
   On Error GoTo DiaErr1
   'cmbPrt.Clear
   sSql = "SELECT DISTINCT PARTREF,PARTNUM FROM " _
          & "PartTable where paobsolete = 0  ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim iList As Integer
   On Error Resume Next
   MouseCursor 13
   If bGoodPart = 1 And bDataHasChanged Then UpdatePartSet True
   
   'update On Dock Requirements
   For iList = 1 To iPartCount
      If bODChng(iList) Then
         sSql = "UPDATE PoitTable SET PIONDOCK=" _
                & bODReqd(iList) & " WHERE (PIPART='" _
                & sODPart(iList) & "' AND PIAQTY=0 AND PITYPE=14)"
         clsADOCon.ExecuteSql sSql
      End If
   Next
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   'RdoEdit.Close
   If optVewAcct.Value = vbChecked Then Unload VewAcct
   If bFrom = 0 Then FormUnload
   If Val(cmbTyp) < 4 Then SaveCurrentSelections
   MouseCursor 0
   Set AdoEdit = Nothing
   Set InvcINe01a = Nothing
   
End Sub



Private Sub optCom_Click()
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PACOMMISSION = optCom.Value
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub

Private Sub optIns_Click()
   If optIns.Value <> bODReqd(iPartCount) Then
      bODReqd(iPartCount) = optIns.Value
      bODChng(iPartCount) = 1
      If bGoodPart = 1 Then
         On Error Resume Next
         AdoEdit!PAONDOCK = optIns.Value
         AdoEdit.Update
         If Err Then RefreshPart
      End If
   End If
   
End Sub

Private Sub optLot_Click()
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PALOTTRACK = optLot.Value
      AdoEdit.Update
      If Err Then RefreshPart
      
      chkLotsExpire.Enabled = optLot.Value
      If optLot.Value = vbUnchecked Then
         chkLotsExpire.Value = vbUnchecked
      End If
   Else
      
'      If (bOpeningPart = 0) Then
'         If (optLot.Value = vbUnchecked) Then
'            optLot.Value = vbChecked
'         End If
'      End If
   End If
   
End Sub

Private Sub chkLotsExpire_Click()
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PALOTSEXPIRE = chkLotsExpire.Value
      AdoEdit.Update
      If Err Then RefreshPart
   End If
End Sub

Private Sub chkInactive_Click()
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAINACTIVE = chkInactive.Value
      AdoEdit.Update
      If Err Then RefreshPart
   End If
End Sub

Private Sub chkObsolete_Click()
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAOBSOLETE = chkObsolete.Value
      AdoEdit.Update
      If Err Then RefreshPart
   End If
End Sub

Private Sub chkITAR_Click()
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAITARRPT = chkITAR.Value
      AdoEdit.Update
      If Err Then RefreshPart
   End If
End Sub


Private Sub chkEAR_Click()
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAEARRPT = chkEAR.Value
      AdoEdit.Update
      If Err Then RefreshPart
   End If
End Sub

Private Sub optTaxEx_Click()
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PATAXEXEMPT = optTaxEx.Value
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub

Private Sub tab1_Click()
   Dim b As Byte
   For b = 0 To 5
      tabFrame(b).Visible = False
   Next
   On Error Resume Next
   ClearCombos
   tabFrame(tab1.SelectedItem.Index - 1).Visible = True
   'tabFrame(tab1.SelectedItem.Index - 1).ZOrder = 0   'bring to front
   tabFrame(tab1.SelectedItem.Index - 1).Top = tabFrame(0).Top
   tabFrame(tab1.SelectedItem.Index - 1).Left = tabFrame(0).Left
   Select Case tab1.SelectedItem.Index
      Case 1
         cmbTyp.SetFocus
      Case 2
         txtWht.SetFocus
      Case 3
'         If txtQoh.Enabled = True Then
'            txtQoh.SetFocus
'         Else
            txtStd.SetFocus
'         End If
      Case 4
         txtSfs.SetFocus
      Case 5
         txtRev.SetFocus
      Case 6
         txtWla.SetFocus
   End Select
   
End Sub

Private Sub Tab1_GotFocus()
   On Error Resume Next
   iTimer = 0
   If bActFilled = 0 Then FillAccounts
   
End Sub

Private Sub Timer1_Timer(Index As Integer)
   Dim bReponse As Byte
   Dim sMsg As String
   iTimer = iTimer + 1
'   If iTimer > 300 Then
'      '6000/300 sec 30 min
'      On Error Resume Next
'      'RdoEdit.Close
'      sMsg = Format(Now, "mm/dd/yy hh:mm AM/PM") _
'             & " This Normal Time Out." & vbCrLf _
'             & "All Current Data Has Been Saved."
'      MsgBox sMsg, vbInformation, Caption
'      Timer1(0).Enabled = False
'      Unload Me
'   End If
   
End Sub

Private Sub txtColor_LostFocus()
   
   Dim ErrVal As Boolean
   txtColor = CheckLen(txtColor, 30)
   txtColor = StrCase(txtColor)
   If bGoodPart = 1 Then
      On Error Resume Next
      UpdateDetail "ColorID", txtColor, ErrVal
    If Not ErrVal Then
      'RdoEdit.Edit
      AdoEdit!PACOLOR = Trim(txtColor)
      AdoEdit.Update
    End If
      If Err Then RefreshPart
   End If
'
'   txtColor = CheckLen(txtColor, 30)
'   txtColor = StrCase(txtColor)
'   If bGoodPart = 1 Then
'      On Error Resume Next
'      AdoEdit!PACOLOR = Trim(txtColor)
'      AdoEdit.Update
'      If Err Then RefreshPart
'   End If
   
End Sub


Private Sub txtComt_LostFocus()
   txtComt = CheckLen(txtComt, 3072)
   txtComt = StrCase(txtComt, ES_FIRSTWORD)
   iTimer = 0
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAEXTDESC = Trim(txtComt)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtDis_Click()
   GetAccount txtDis, "txtDis"
   
End Sub


Private Sub txtDis_LostFocus()
   txtDis = CheckLen(txtDis, 12)
   GetAccount txtDis, "txtDis"
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PADISACCT = Compress(txtDis)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PADESC = "" & txtDsc
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   iPartCount = iPartCount + 1
   If iPartCount > 299 Then
      MsgBox "This Is A Normal Refresh After 300 Edits.  The Form " & vbCrLf _
         & "Will Close And You May Reopen and Continue.", _
         vbInformation, Caption
      Unload Me
   End If
   If optTol.Value = vbChecked And bGoodPart = 1 Then
      sSql = "UPDATE TohdTable SET TOOL_DESC='" & Trim(txtDsc) & "' " _
             & "WHERE TOOL_PARTREF='" & Compress(cmbPrt) & "'"
      clsADOCon.ExecuteSql sSql
   End If
   sODPart(iPartCount) = Compress(cmbPrt)
   bODReqd(iPartCount) = optIns.Value
   bODChng(iPartCount) = 0
   
End Sub

Private Sub txtEoq_LostFocus()
   txtEoq = CheckLen(txtEoq, 7)
   txtEoq = Format(Abs(Val(txtEoq)), ES_QuantityDataFormat)
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAEOQ = Format(Val(txtEoq), ES_QuantityDataFormat)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub





Private Sub txtGex_Click()
   GetAccount txtGex, "txtGex"
   
End Sub


Private Sub txtGex_LostFocus()
   txtGex = CheckLen(txtGex, 12)
   GetAccount txtGex, "txtGex"
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PACGSEXPACCT = Compress(txtGex)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtGla_Click()
   GetAccount txtGla, "txtGla"
   
End Sub


Private Sub txtGla_LostFocus()
   txtGla = CheckLen(txtGla, 12)
   GetAccount txtGla, "txtGla"
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PACGSLABACCT = Compress(txtGla)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtGma_Click()
   GetAccount txtGma, "txtGma"
   
End Sub


Private Sub txtGma_LostFocus()
   txtGma = CheckLen(txtGma, 12)
   GetAccount txtGma, "txtGma"
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PACGSMATACCT = Compress(txtGma)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtGoh_Click()
   GetAccount txtGoh, "txtGoh"
   
End Sub


Private Sub txtGoh_LostFocus()
   txtGoh = CheckLen(txtGoh, 12)
   GetAccount txtGoh, "txtGoh"
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PACGSOHDACCT = Compress(txtGoh)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtLoc_LostFocus()
   txtLoc = CheckLen(txtLoc, 4)
   iTimer = 0
   If bGoodPart = 1 Then
      'On Error Resume Next
      AdoEdit!PALOCATION = "" & txtLoc
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
   'if there are lots with a different location, ask whether to update
   Dim part As New ClassPart
   part.UpdatePartLotLocations cmbPrt, txtLoc, False
End Sub



Private Sub txtMatAlloy_LostFocus()
   
   Dim ErrVal As Boolean
   txtMatAlloy = CheckLen(txtMatAlloy, 30)
   txtMatAlloy = StrCase(txtMatAlloy)
   If bGoodPart = 1 Then
      On Error Resume Next
      UpdateDetail "ALloyID", txtMatAlloy, ErrVal
     If Not ErrVal Then
      'RdoEdit.Edit
      AdoEdit!PAMATERIALALLOY = Trim(txtMatAlloy)
      AdoEdit.Update
     End If
      If Err Then RefreshPart
   End If
   
'   txtMatAlloy = CheckLen(txtMatAlloy, 30)
'   If bGoodPart = 1 Then
'      On Error Resume Next
'      AdoEdit!PAMATERIALALLOY = Trim(txtMatAlloy)
'      AdoEdit.Update
'      If Err Then RefreshPart
'   End If
   
End Sub


Private Sub txtMatType_LostFocus()
   
   Dim ErrVal As Boolean
   txtMatType = CheckLen(txtMatType, 30)
   txtMatType = StrCase(txtMatType)
   If bGoodPart = 1 Then
      On Error Resume Next
      UpdateDetail "MaterialID", txtMatType, ErrVal
     If Not ErrVal Then
      'RdoEdit.Edit
      AdoEdit!PAMATERIALTYPE = Trim(txtMatType)
      AdoEdit.Update
     End If
      If Err Then RefreshPart
   End If
   
'   txtMatType = CheckLen(txtMatType, 30)
'   txtMatType = StrCase(txtMatType)
'   If bGoodPart = 1 Then
'      On Error Resume Next
'      AdoEdit!PAMATERIALTYPE = Trim(txtMatType)
'      AdoEdit.Update
'      If Err Then RefreshPart
'   End If
   
End Sub

Private Sub txtCustFld1_LostFocus()
   
   Dim ErrVal As Boolean
   txtCustFld1 = CheckLen(txtCustFld1, 30)
   If bGoodPart = 1 Then
      On Error Resume Next
      UpdateDetail "CustFld1ID", txtCustFld1, ErrVal
     If Not ErrVal Then
      'RdoEdit.Edit
      AdoEdit!PACUSTFLD1 = Trim(txtCustFld1)
      AdoEdit.Update
     End If
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtMft_LostFocus()
   txtMft = CheckLen(txtMft, 7)
   txtMft = Format(Abs(Val(txtMft)), "###0.00")
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAFLOWTIME = Val(txtMft)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub

Private Sub txtOst_LostFocus()
   txtOst = CheckLen(txtOst, 7)
   txtOst = Format(Abs(Val(txtOst)), ES_QuantityDataFormat)
   iTimer = 0
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAOVERSTOCK = Val(txtOst)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtPlt_LostFocus()
   txtPlt = CheckLen(txtPlt, 7)
   txtPlt = Format(Abs(Val(txtPlt)), "###0.00")
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PALEADTIME = Val(txtPlt)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub

Private Sub txtPrc_LostFocus()
   txtPrc = CheckLen(txtPrc, 10)
   If Val(txtPrc) = 0 Then txtPrc = 1
   txtPrc = Format(Abs(Val(txtPrc)), ES_QuantityDataFormat)
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAPURCONV = Val(txtPrc)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtPrt_GotFocus()
   cmbPrt = txtPrt
   'bGoodPart = GetPart(0)
   bGoodPart = GetPart(1)
   
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
   If bCanceled = 1 Or bShowParts = 1 Then Exit Sub
   If Len(Trim(txtPrt.Text)) = 0 Then
      cmdCan.SetFocus
      Exit Sub
   End If
   cmbPrt = txtPrt
   bGoodPart = GetPart(1)
   If bGoodPart = 0 Then
      ' TODO
      Addpart
      bGoodPart = 1
   Else
      sPassedPart = txtPrt
   End If
   
End Sub


Private Sub txtPun_LostFocus()
   txtPun = CheckLen(txtPun, 2)
   If Len(txtPun) = 0 Then txtPun = txtUom
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAPUNITS = "" & txtPun
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub



Private Sub txtRev_Click()
   GetAccount txtRev, "txtRev"
   
End Sub


Private Sub txtRev_LostFocus()
   txtRev = CheckLen(txtRev, 12)
   GetAccount txtRev, "txtRev"
   iTimer = 0
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAREVACCT = Compress(txtRev)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtRrq_LostFocus()
   txtRrq = CheckLen(txtRrq, 6)
   txtRrq = Format(Val(txtRrq), "#####0")
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PARRQ = Val(txtRrq)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub

Private Sub txtSfs_LostFocus()
   txtSfs = CheckLen(txtSfs, 7)
   txtSfs = Format(Abs(Val(txtSfs)), ES_QuantityDataFormat)
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PASAFETY = Val(txtSfs)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub

Private Sub txtSpr_LostFocus()
   txtSpr = CheckLen(txtSpr, 10)
   txtSpr = Format(Abs(Val(txtSpr)), ES_QuantityDataFormat)
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAPRICE = Format(Val(txtSpr), ES_QuantityDataFormat)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub

Private Sub txtSsq_LostFocus()
   txtSsq = CheckLen(txtSsq, 7)
   txtSsq = Format(Abs(Val(txtSsq)), ES_QuantityDataFormat)
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PASHIPSET = Val(txtSsq)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtStd_LostFocus()
   Dim b As Byte
   Dim sPart As String
   Dim cCost As Currency
   txtStd = CheckLen(txtStd, 10)
   txtStd = Format(Abs(Val(txtStd)), ES_QuantityDataFormat)
   iTimer = 0
   ' GetPrice
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PASTDCOST = Format(Val(txtStd), ES_QuantityDataFormat)
      AdoEdit!PAPRICE = Format(Val(txtSpr), ES_QuantityDataFormat)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
End Sub

Private Sub txtTcg_Click()
   GetAccount txtTcg, "txtTcg"
   
End Sub


Private Sub txtTcg_LostFocus()
   txtTcg = CheckLen(txtTcg, 12)
   GetAccount txtTcg, "txtTcg"
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PATFRCGSACCT = Compress(txtTcg)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtTrv_Click()
   GetAccount txtTrv, "txtTrv"
   
End Sub


Private Sub txtTrv_LostFocus()
   txtTrv = CheckLen(txtTrv, 12)
   GetAccount txtTrv, "txtTrv"
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PATFRREVACCT = Compress(txtTrv)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub



Private Sub FillCombo()
   Dim b As Byte
   Dim iList As Integer
   On Error GoTo DiaErr1
   MouseCursor 13
   sProcName = "fillcombo"
   If ES_PARTCOUNT < 5001 Then
      cmbPrt.Clear
      sSql = "Qry_FillParts"
      LoadComboBox cmbPrt
      sPassedPart = Trim(cUR.CurrentPart)
      If Trim(sPassedPart) <> "" Then
         cmbPrt = cUR.CurrentPart
      Else
         If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
      End If
   End If
   
   MouseCursor 0
   sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   
   
   If b = 0 Then
      MsgBox "There Is No Open Inventory Journal For This Period.", _
         vbExclamation, Caption
      Sleep 500
   End If
   cmbAbc.Clear
   sSql = "Qry_FillABCCombo"
   cmbAbc.AddItem "  "
   LoadComboBox cmbAbc
   If Not bSqlRows Then
      For iList = 65 To 89
         cmbAbc.AddItem Chr$(iList) & "+"
         cmbAbc.AddItem Chr$(iList)
         cmbAbc.AddItem Chr$(iList) & "-"
      Next
      cmbAbc.AddItem Chr$(iList) & "+"
      cmbAbc.AddItem Chr$(iList)
      cmbAbc.AddItem Chr$(iList) & "-"
   End If
   If b = 0 Then
      Unload Me
   Else
      bGoodPart = GetPart(1)
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Addpart()
   Dim bResponse As Byte
   'Dim iTrans As Integer
   'Dim iRef As Integer
   Dim lCOUNTER As Long
   Dim cCost As Currency
   
   Dim sADate As String
   Dim sNewPart As String
   Dim sMbe As String
   
   On Error GoTo DiaErr1
   If bShowParts = 1 Then Exit Sub
   sADate = Format$(Now, "mm/dd/yy")
   bResponse = MsgBox(cmbPrt & "  Wasn't Found. Add It?", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      bGoodPart = 0
      On Error Resume Next
      If ES_PARTCOUNT < 5001 Then
         cmbPrt = cmbPrt - cmbPrt.List(0)
         cmbPrt.SetFocus
      Else
         txtPrt.SetFocus
      End If
      CancelTrans
      Exit Sub
   End If
   sNewPart = Compress(cmbPrt)
   If sNewPart = "ALL" Then
      MsgBox "ALL Is An Illegal Part Number.", vbExclamation, Caption
      Exit Sub
   End If
   bResponse = IllegalCharacters(cmbPrt)
   If bResponse > 0 Then
      MsgBox "The Part Number Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   MouseCursor 13
   If bFirstPart And iPalevel = 0 Then
      iPalevel = cmbTyp
      sMbe = Trim(cmbMbe)
   Else
      If Val(cmbTyp) > 0 Then iPalevel = Val(cmbTyp) Else iPalevel = 3
      If Len(cmbMbe) > 0 Then sMbe = cmbMbe
   End If
   bFirstPart = 0
   
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   
   Dim part As New ClassPart
   Dim level As ePartType
   level = iPalevel
   part.CreateNewPart cmbPrt, level, "", sMbe
   clsADOCon.CommitTrans
   
   sSql = "UPDATE PartTable SET PALOTTRACK = 1 WHERE PARTREF='" & sNewPart & "'"
   clsADOCon.ExecuteSql sSql
   
   sSql = "Select * FROM PartTable WHERE PARTREF='" & sNewPart & "'"
   Set AdoEdit = clsADOCon.GetRecordSet(sSql)
'   Set RdoEdit = RdoCon.OpenResultset(sSql, rdOpenKeyset, rdConcurRowVer)
   bGoodPart = 1
   AddComboStr cmbPrt.hWnd, cmbPrt
   sNewPartNo = cmbPrt
   bNewPart = 1
   txtComt = " "
   cmbTyp = Trim(str(iPalevel))
   txtUom = "EA"
   txtDsc = " "
   txtLoc = " "
   txtPun = " "
'   txtQoh = "0.000"
   txtAve = "0.000"
   txtStd = "0.000"
   txtSpr = "0.000"
   txtWht = "0.000"
   txtPrc = "1.000"
   txtSfs = "0.000"
   txtOst = "0.000"
   txtRrq = "0"
   txtEoq = "0.000"
   txtSsq = "0.000"
   txtMft = "0.00"
   txtPlt = "0.00"
   
   optCom = vbUnchecked
   optTol = vbUnchecked
   optBom = vbUnchecked
   optLot = getLotTrack
   chkLotsExpire = vbUnchecked
   optDte = vbUnchecked
   cmbTyp.Enabled = True
   txtUom.Enabled = True
   txtComt.Enabled = True
   'txtQoh.Enabled = True
   
   If AdoEdit!PAUSEACTUALCOST = 0 Then
     cstStd.Value = True
   Else
     cstAct.Value = True
   End If
   
   bDataHasChanged = False
   txtDsc.SetFocus
   MouseCursor 0
   cOldCost = Abs(Val(txtStd))
   SysMsg cmbPrt & " Added.", True
   Exit Sub
   
DiaErr1:
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   On Error Resume Next
   cOldCost = 0
   clsADOCon.RollbackTrans
   bGoodPart = 0
   MouseCursor 0
   MsgBox CurrError.Description & vbCr & "Couldn't Add Part.", vbExclamation, Caption
   
End Sub

Private Function getLotTrack() As Byte
   On Error Resume Next
   Dim RdoLotTrack As ADODB.Recordset
   sSql = "Select COLOTSACTIVE from ComnTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLotTrack, ES_FORWARD)
   If bSqlRows Then
    getLotTrack = Trim(RdoLotTrack.Fields(0))
   End If
   Set RdoLotTrack = Nothing

End Function

Private Function getPrice(ByRef lPrice As Double, ByRef lDte As String) As Byte
   On Error Resume Next
   Dim RdoPrc As ADODB.Recordset
   sSql = "select MAX(PIAMT), PIADATE from poitTable where PIPART = '" & Compress(cmbPrt) & "' and " & _
            " PIADATE = (select max(PIADATE) from poitTable where PIPART = '" & Compress(cmbPrt) & "') group by PIADATE"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrc, ES_FORWARD)
   If bSqlRows Then
      With RdoPrc
         lPrice = Trim(.Fields(0))
         lDte = Trim(.Fields(1))
      End With
   End If
   Set RdoPrc = Nothing

End Function


Public Function GetPart(bOpen As Byte) As Byte
   Dim b As Byte
   Dim lPrice As Double
   Dim lDate As String
   
   If bGoodPart = 1 And bDataHasChanged Then UpdatePartSet
   On Error GoTo DiaErr1
   GetPart = 0
   bGoodPart = 0
   sLotNumber = ""
   bOpeningPart = 0
   
   getPrice lPrice, lDate
   
   sSql = "SELECT TOP 1 * FROM PartTable WHERE PARTREF= '" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoEdit, ES_KEYSET)
   If bSqlRows Then
      'Speed it up by showing just a few
      With AdoEdit
         MouseCursor 13
         cmbPrt = "" & Trim(!PartNum)
         txtDsc = "" & Trim(!PADESC)
         cmbTyp = Format(!PALEVEL, "0")
         txtUom = "" & Trim(!PAUNITS)
         cmbMbe = "" & Trim(!PAMAKEBUY)
         txtRrq = !PARRQ
         txtComt = "" & Trim(!PAEXTDESC)
         txtDoc = "" & Trim(!PADOCLISTREV)
         If Trim(!PADOCLISTREF) = "" And Trim(!PADOCLISTREV) = "" Then _
                 txtDoc = "NONE"
         cmbCde = "" & Trim(!PAPRODCODE)
         cmbCls = "" & Trim(!PACLASS)
         optTol.Value = !PATOOL
         If bOpen = 1 Then
            CheckPartType
            bOpeningPart = 1
            ''Tab1.Enabled = False
            txtWht = Format(!PAWEIGHT, ES_QuantityDataFormat)
            If !PAPURCONV <> 0 Then
               txtPrc = Format(!PAPURCONV, ES_QuantityDataFormat)
            Else
               txtPrc = Format(1, ES_QuantityDataFormat)
            End If
            txtPun = "" & Trim(!PAPUNITS)
            txtLoc = "" & Trim(!PALOCATION)
            cmbAbc = "" & Trim(!PAABC)
            txtQoh = Format(!PAQOH, ES_QuantityDataFormat)
            txtAve = Format(!PAAVGCOST, ES_QuantityDataFormat)
            txtStd = Format(!PASTDCOST, ES_QuantityDataFormat)
            txtSpr = Format(!PAPRICE, ES_QuantityDataFormat)
            
            txtPrice = lPrice
            txtDate = lDate
            
            optIns.Value = !PAONDOCK
                  optLot.Value = !PALOTTRACK
            chkLotsExpire.Value = !PALOTSEXPIRE
            chkLotsExpire.Enabled = optLot.Value
            If optLot.Value = vbUnchecked Then
               chkLotsExpire.Value = vbUnchecked
            End If
            chkInactive.Value = !PAINACTIVE
            chkObsolete.Value = !PAOBSOLETE
            
            chkITAR.Value = IIf(IsNull(!PAITARRPT), False, !PAITARRPT)
            chkEAR.Value = IIf(IsNull(!PAEARRPT), False, !PAEARRPT)
            
            optTaxEx.Value = !PATAXEXEMPT
            If Not IsNull(!PAPRICE) Then
               cOldCost = !PAPRICE
            Else
               cOldCost = 0
            End If
            txtSfs = Format(!PASAFETY, ES_QuantityDataFormat)
            txtOst = Format(!PAOVERSTOCK, ES_QuantityDataFormat)
            txtEoq = Format(!PAEOQ, ES_QuantityDataFormat)
            txtSsq = Format(!PASHIPSET, ES_QuantityDataFormat)
            txtMft = Format(!PAFLOWTIME, "###0.00")
            txtPlt = Format(!PALEADTIME, "###0.00")
            
            'ACCOUNTS 3/25/99
            txtRev = "" & Trim(!PAREVACCT)
            GetAccount txtRev, "txtRev"
            
            txtDis = "" & Trim(!PADISACCT)
            GetAccount txtDis, "txtDis"
            
            txtTrv = "" & Trim(!PATFRREVACCT)
            GetAccount txtTrv, "txtTrv"
            
            txtTcg = "" & Trim(!PATFRCGSACCT)
            GetAccount txtTcg, "txtTcg"
            
            '9/15/99
            txtWla = "" & Trim(!PAINVLABACCT)
            GetAccount txtWla, "txtWla"
            
            txtWex = "" & Trim(!PAINVEXPACCT)
            GetAccount txtWex, "txtWex"
            
            txtWma = "" & Trim(!PAINVMATACCT)
            GetAccount txtWma, "txtWma"
            
            txtWoh = "" & Trim(!PAINVOHDACCT)
            GetAccount txtWoh, "txtWoh"
            
            txtGla = "" & Trim(!PACGSLABACCT)
            GetAccount txtGla, "txtGla"
            
            txtGex = "" & Trim(!PACGSEXPACCT)
            GetAccount txtGex, "txtGex"
            
            txtGma = "" & Trim(!PACGSMATACCT)
            GetAccount txtGma, "txtGma"
            
            txtGoh = "" & Trim(!PACGSOHDACCT)
            GetAccount txtGoh, "txtGoh"
            
            '8/28/03
            optCom.Value = !PACOMMISSION
            
            If !PALEVEL < 5 Then cUR.CurrentPart = cmbPrt
            'b = GetNewPartAccounts(sDebitAcct, sCreditAcct)
            Dim part As New ClassPart
            part.GetPartAccounts cmbPrt, True, sDebitAcct, sCreditAcct
            If sNewPartNo <> cmbPrt Then
               bNewPart = 0
               sNewPartNo = ""
            End If
            'Tab1.Enabled = True
            If !PAUSEACTUALCOST = 0 Then
               cstStd.Value = True
            Else
               cstAct.Value = True
            End If
            If optTol.Value = vbChecked Then
               cmbTyp.Enabled = False
               optLot.Enabled = False
               chkLotsExpire.Enabled = False
            Else
               cmbTyp.Enabled = True
               optLot.Enabled = True
               chkLotsExpire.Enabled = True
            End If
            '1/6/06
            'txtMatType = "" & Trim(!PAMATERIALTYPE)
            'txtMatAlloy = "" & Trim(!PAMATERIALALLOY)
            'txtColor = "" & Trim(!PACOLOR)
            Dim strLabelDesc As String
            Dim strTxtVal As String
            Dim ErrVal As Boolean
            GetColDetail "MaterialID", strLabelDesc, strTxtVal, ErrVal
            
            If ErrVal Then
              txtMatType = "" & Trim(strTxtVal)
              lblMatType = strLabelDesc
              strLabelDesc = ""
              strTxtVal = ""
            
              GetColDetail "AlloyID", strLabelDesc, strTxtVal, ErrVal
              txtMatAlloy = "" & Trim(strTxtVal)
              lblMatAlloy = strLabelDesc
              strLabelDesc = ""
              strTxtVal = ""
            
              GetColDetail "ColorID", strLabelDesc, strTxtVal, ErrVal
              txtColor = "" & Trim(strTxtVal)
              lblColor = strLabelDesc
              strLabelDesc = ""
              strTxtVal = ""
              
              GetColDetail "CustFld1ID", strLabelDesc, strTxtVal, ErrVal
              txtCustFld1 = "" & Trim(strTxtVal)
              lblCustFlds1 = strLabelDesc
              strLabelDesc = ""
              strTxtVal = ""
              
            Else
              lblMatType = "Material"
              txtMatType = "" & Trim(!PAMATERIALTYPE)
              lblMatAlloy = "Alloy/Temper"
              txtMatAlloy = "" & Trim(!PAMATERIALALLOY)
              lblColor = "Color"
              txtColor = "" & Trim(!PACOLOR)
            End If
            
            bOpeningPart = 0
   
         End If
         lblPCode = GetProductCode()
         lblPClass = GetProductClass()
         MouseCursor 0
      End With
      Err.Clear
      GetPart = 1
   Else
      'Tab1.Enabled = True
      bNewPart = 0
      sNewPartNo = ""
      GetPart = 0
      txtDsc = ""
      If iPalevel = 0 Then
         cmbTyp = "3"
      Else
         cmbTyp = Trim(str(iPalevel))
      End If
      txtUom = "EA"
      cmbMbe = "M"
      txtComt = " "
      txtAve = "0.000"
      txtQoh = "0.000"
      txtStd = "0.000"
      txtSpr = "0.000"
      txtPun = "EA"
      txtPrc = "1.0000"
      txtSfs = "0.000"
      txtEoq = "0.000"
      txtMft = "0.00"
      txtPlt = "0.00"
      'RdoEdit.Close
   End If
   bDataHasChanged = False
   iTimer = 0
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   GetPart = 0
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Public Sub txtUom_LostFocus()
   txtUom = CheckLen(txtUom, 2)
   If Len(txtUom) < 2 Then
      MsgBox "2 Characters Please", vbInformation, Caption
      txtUom = "EA"
   End If
   On Error Resume Next
   If bGoodPart = 1 Then
      AdoEdit!PAUNITS = "" & txtUom
      If bNewPart Then AdoEdit!PAPUNITS = "" & txtUom
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub



Private Sub txtWex_Click()
   GetAccount txtWex, "txtWex"
   
End Sub


Private Sub txtWex_LostFocus()
   txtWex = CheckLen(txtWex, 12)
   GetAccount txtWex, "txtWex"
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAINVEXPACCT = Compress(txtWex)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtWhs_LostFocus()
   txtWhs = CheckLen(txtWhs, 4)
   
End Sub

Private Sub txtWht_LostFocus()
   txtWht = CheckLen(txtWht, 10)
   txtWht = Format(Abs(Val(txtWht)), ES_QuantityDataFormat)
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAWEIGHT = Format(Val(txtWht), ES_QuantityDataFormat)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub




Private Sub FillAccounts()
   Dim iList As Integer
   Dim RdoGlm As ADODB.Recordset
   On Error GoTo DiaErr1
   MouseCursor 13
   If bActFilled = 1 Then Exit Sub
   
   sSql = "Qry_FillLowAccounts"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm)
   If bSqlRows Then
      With RdoGlm
         Do Until .EOF
            AddComboStr txtRev.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtDis.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtTrv.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtTcg.hWnd, "" & Trim(!GLACCTNO)
            
            
            'Inv/Exp
            AddComboStr txtWla.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWma.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWoh.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtWex.hWnd, "" & Trim(!GLACCTNO)
            
            'CGS
            AddComboStr txtGla.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGma.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGoh.hWnd, "" & Trim(!GLACCTNO)
            AddComboStr txtGex.hWnd, "" & Trim(!GLACCTNO)
            
            .MoveNext
         Loop
         ClearResultSet RdoGlm
      End With
   Else
      bNoAccts = 1
      CloseBoxes
   End If
   bActFilled = 1
   Set RdoGlm = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   bNoAccts = 1
   CloseBoxes
   
End Sub

Private Sub CloseBoxes()
   txtRev.Enabled = False
   txtDis.Enabled = False
   txtTrv.Enabled = False
   txtTcg.Enabled = False
   
End Sub

Private Sub GetAccount(sAccount As String, sBox As String)
   Dim RdoGlm As ADODB.Recordset
   On Error GoTo DiaErr1
   If bNoAccts = 0 Then
      sSql = "Qry_GetAccount '" & Compress(sAccount) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm)
      If bSqlRows Then
         With RdoGlm
            Select Case sBox
               Case "txtRev"
                  txtRev = "" & Trim(!GLACCTNO)
                  lblRev = "" & Trim(!GLDESCR)
               Case "txtDis"
                  txtDis = "" & Trim(!GLACCTNO)
                  lblDis = "" & Trim(!GLDESCR)
               Case "txtTrv"
                  txtTrv = "" & Trim(!GLACCTNO)
                  lblTrv = "" & Trim(!GLDESCR)
               Case "txtTcg"
                  txtTcg = "" & Trim(!GLACCTNO)
                  lblTcg = "" & Trim(!GLDESCR)
                  'Inv/Exp
               Case "txtWla"
                  txtWla = "" & Trim(!GLACCTNO)
                  lblWla = "" & Trim(!GLDESCR)
               Case "txtWex"
                  txtWex = "" & Trim(!GLACCTNO)
                  lblWex = "" & Trim(!GLDESCR)
               Case "txtWma"
                  txtWma = "" & Trim(!GLACCTNO)
                  lblWma = "" & Trim(!GLDESCR)
               Case "txtWoh"
                  txtWoh = "" & Trim(!GLACCTNO)
                  lblWoh = "" & Trim(!GLDESCR)
                  'CGS
               Case "txtGla"
                  txtGla = "" & Trim(!GLACCTNO)
                  lblGla = "" & Trim(!GLDESCR)
               Case "txtGex"
                  txtGex = "" & Trim(!GLACCTNO)
                  lblGex = "" & Trim(!GLDESCR)
               Case "txtGma"
                  txtGma = "" & Trim(!GLACCTNO)
                  lblGma = "" & Trim(!GLDESCR)
               Case "txtGoh"
                  txtGoh = "" & Trim(!GLACCTNO)
                  lblGoh = "" & Trim(!GLDESCR)
            End Select
            ClearResultSet RdoGlm
         End With
      Else
         Select Case sBox
            Case "txtRev"
               txtRev = ""
               lblRev = ""
            Case "txtDis"
               txtDis = ""
               lblDis = ""
            Case "txtTrv"
               txtTrv = ""
               lblTrv = ""
            Case "txtTcg"
               txtTcg = ""
               lblTcg = ""
            Case "txtWla"
               txtWla = ""
               lblWla = ""
            Case "txtWex"
               txtWex = ""
               lblWex = ""
            Case "txtWma"
               txtWma = ""
               lblWma = ""
            Case "txtWoh"
               txtWoh = ""
               lblWoh = ""
         End Select
      End If
   End If
   Set RdoGlm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getaccount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub txtWla_Click()
   GetAccount txtWla, "txtWla"
   
End Sub


Private Sub txtWla_LostFocus()
   txtWla = CheckLen(txtWla, 12)
   GetAccount txtWla, "txtWla"
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAINVLABACCT = Compress(txtWla)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtWma_Click()
   GetAccount txtWma, "txtWma"
   
End Sub


Private Sub txtWma_LostFocus()
   txtWma = CheckLen(txtWma, 12)
   GetAccount txtWma, "txtWma"
   iTimer = 0
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAINVMATACCT = Compress(txtWma)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub


Private Sub txtWoh_Click()
   GetAccount txtWoh, "txtWoh"
   
End Sub


Private Sub txtWoh_LostFocus()
   txtWoh = CheckLen(txtWoh, 12)
   GetAccount txtWoh, "txtWoh"
   If bGoodPart = 1 Then
      On Error Resume Next
      AdoEdit!PAINVOHDACCT = Compress(txtWoh)
      AdoEdit.Update
      If Err Then RefreshPart
   End If
   
End Sub



Private Sub CheckPartType()
   Dim b As Byte
   On Error Resume Next
   Select Case Val(cmbTyp)
      Case 4
         txtWma.Enabled = True
         txtGma.Enabled = True
         
         txtWla.Enabled = False
         txtWex.Enabled = False
         txtWoh.Enabled = False
         
         txtGla.Enabled = False
         txtGex.Enabled = False
         txtGoh.Enabled = False
      Case 5, 6, 7
         txtWex.Enabled = True
         txtGex.Enabled = True
         
         txtWla.Enabled = False
         txtWma.Enabled = False
         txtWoh.Enabled = False
         
         txtGla.Enabled = False
         txtGma.Enabled = False
         txtGoh.Enabled = False
      Case Else
         txtWla.Enabled = True
         txtWma.Enabled = True
         txtWoh.Enabled = True
         txtWex.Enabled = True
         
         txtGla.Enabled = True
         txtGma.Enabled = True
         txtGoh.Enabled = True
         txtGex.Enabled = True
   End Select
'   b = GetNewPartAccounts(sDebitAcct, sCreditAcct)
   Dim part As New ClassPart
   part.GetPartAccounts cmbPrt, True, sDebitAcct, sCreditAcct
   
   
End Sub

Private Sub RefreshPart()
MsgBox Err.Description
Exit Sub
   MsgBox "The Current Part Was Updated Outside Of This" & vbCrLf _
      & "Function. Press OK To Refresh The Part Number.", _
      vbInformation, Caption
   bGoodPart = GetPart(1)
   
End Sub

'Timestamps to ensure bogus Journal entries get dumped

Private Function GetTimeStamp() As Byte
   Dim l As Long
   Dim r As Single
   Dim S As String
   Dim P As String
   
   r = TimeValue(Now)
   S = str(r)
   S = Mid(S, 3, 5)
   
   r = DateValue(Now)
   P = Trim(str(r))
   P = Right(P, 1)
   
   sTimeStamp1 = P & S
   
   r = TimeValue(Now) + 0.00001
   S = str(r)
   S = Mid(S, 3, 5)
   
   r = DateValue(Now)
   P = Trim(str(r))
   P = Right(P, 1)
   
   sTimeStamp2 = P & S
   
   
   
End Function


Private Sub ClearCombos()
   Dim iControl As Integer
   For iControl = 0 To Controls.count - 1
      If TypeOf Controls(iControl) Is ComboBox Then _
                         Controls(iControl).SelLength = 0
   Next
   
End Sub

'10/2/04 Fixed jumps

Private Sub UpdatePartSet(Optional ShowMouse As Boolean)
   On Error GoTo DiaErr1
   If ShowMouse Then MouseCursor 13
   If bGoodPart = 1 Then
      If Val(cmbTyp) = 0 Then cmbTyp = "1"
      With AdoEdit
         If bDataHasChanged = 1 Then !PAREVDATE = Format(ES_SYSDATE, "mm/dd/yy hh:mm")
         !PAPRODCODE = Compress(cmbCde)
         !PACLASS = Compress(cmbCls)
         !PAMAKEBUY = cmbMbe
         !PALEVEL = Val(cmbTyp)
         !PAUSEACTUALCOST = Abs(cstAct.Value)
         !PACOMMISSION = optCom.Value
         !PAONDOCK = optIns.Value
         !PATAXEXEMPT = optTaxEx.Value
         !PAEXTDESC = Trim(txtComt)
         !PADISACCT = Compress(txtDis)
         !PADESC = Trim(txtDsc)
         !PAEOQ = Format(Val(txtEoq), ES_QuantityDataFormat)
         !PACGSEXPACCT = Compress(txtGex)
         !PACGSLABACCT = Compress(txtGla)
         !PACGSOHDACCT = Compress(txtGoh)
         !PACGSMATACCT = Compress(txtGma)
         !PALOCATION = Trim(txtLoc)
         !PAFLOWTIME = Format(Val(txtMft), ES_QuantityDataFormat)
         !PALEADTIME = Format(Val(txtPlt), ES_QuantityDataFormat)
         !PAPURCONV = Format(Val(txtPrc), ES_QuantityDataFormat)
         !PAPUNITS = Trim(txtPun)
         !PAREVACCT = Compress(txtRev)
         !PAPRICE = Format(Val(txtSpr), ES_QuantityDataFormat)
         !PASAFETY = Format(Val(txtSfs), ES_QuantityDataFormat)
         !PASTDCOST = Format(Val(txtStd), ES_QuantityDataFormat)
         !PATFRCGSACCT = Compress(txtTcg)
         !PATFRREVACCT = Compress(txtTrv)
         !PAUNITS = "" & txtUom
         !PAABC = Trim(cmbAbc)
         .Update
      End With
      bDataHasChanged = False
   End If
   Exit Sub
   
DiaErr1:
   On Error GoTo 0
   
End Sub


Private Function GetColDetail(strCustID As String, ByRef strLabelDesc As String, ByRef strTxtVal As String, ByRef ErrVal As Boolean)
    
    Dim RdoLbl As ADODB.Recordset
    Dim bSqlRows As Boolean
    Dim strColName, strTblName, strColRef As String
    On Error GoTo DiaErr
    sSql = "select * from CustomFields where LABEL_ID = '" & strCustID & "'"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoLbl, ES_FORWARD)

    If bSqlRows Then
        strLabelDesc = "" & Trim(RdoLbl!LABEL_NAME)
        strColName = "" & Trim(RdoLbl!TBL_COLNAME)
        strTblName = "" & Trim(RdoLbl!TBL_NAME)
        strColRef = "" & Trim(RdoLbl!COLREF_KEY)
        
        Set RdoLbl = Nothing
        sSql = "select " & strColName & " from " & strTblName & " where " & strColRef & " = " & "'" & Compress(cmbPrt) & "'"
        bSqlRows = clsADOCon.GetDataSet(sSql, RdoLbl, ES_FORWARD)
    
        If bSqlRows Then
            strTxtVal = "" & Trim(RdoLbl.Fields(0))
        End If
        Set RdoLbl = Nothing
        ErrVal = True
    End If
    Exit Function
DiaErr:
      ErrVal = False
      Exit Function

End Function

Private Function UpdateDetail(strCustID As String, strTxtVal As String, ByRef ErrVal As Boolean)
    
    Dim RdoLbl As ADODB.Recordset
    Dim bSqlRows As Boolean
    Dim strColName, strTblName, strColRef As String
    On Error GoTo DiaErr
    sSql = "select * from CustomFields where LABEL_ID = '" & strCustID & "'"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoLbl, ES_FORWARD)
    If bSqlRows Then
'        strLabelDesc = "" & Trim(RdoLbl!LABEL_NAME)
        strColName = "" & Trim(RdoLbl!TBL_COLNAME)
        strTblName = "" & Trim(RdoLbl!TBL_NAME)
        strColRef = "" & Trim(RdoLbl!COLREF_KEY)
        
        Set RdoLbl = Nothing
        sSql = "update " & strTblName & " Set " & strColName & " = '" & strTxtVal & _
                "' where " & strColRef & "= '" & Compress(cmbPrt) & "'"
        clsADOCon.ExecuteSql sSql
        ErrVal = True
    Else
        ErrVal = False
    End If
    Exit Function
DiaErr:
      ErrVal = False
      Exit Function
        
End Function



Private Function GetLabelName(strCustID As String, ByRef strLabelDesc As String, ByRef ErrVal As Boolean)
    
    Dim RdoLbl As ADODB.Recordset
    Dim bSqlRows As Boolean
    On Error GoTo DiaErr
    sSql = "select * from CustomFields where LABEL_ID = '" & strCustID & "'"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoLbl, ES_FORWARD)
    If bSqlRows Then
        strLabelDesc = "" & Trim(RdoLbl!LABEL_NAME)
    End If
        Set RdoLbl = Nothing
        ErrVal = True
    Exit Function
DiaErr:
      ErrVal = False
      Exit Function

End Function

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPrt.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPrt.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function

