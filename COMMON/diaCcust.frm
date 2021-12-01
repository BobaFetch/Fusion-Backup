VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form diaCcust 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customers"
   ClientHeight    =   6525
   ClientLeft      =   1950
   ClientTop       =   750
   ClientWidth     =   7620
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "diaCcust.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6525
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Index           =   0
      Interval        =   6000
      Left            =   4680
      Top             =   120
   End
   Begin VB.CheckBox optVewAcct 
      Caption         =   "Just for acct lookup"
      Height          =   255
      Left            =   0
      TabIndex        =   146
      Top             =   6660
      Visible         =   0   'False
      Width           =   2175
   End
   Begin Threed.SSFrame fra1 
      Height          =   30
      Left            =   105
      TabIndex        =   141
      Top             =   960
      Width           =   7335
      _Version        =   65536
      _ExtentX        =   12938
      _ExtentY        =   53
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtNme 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Tag             =   "2"
      Top             =   600
      Width           =   3475
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Customer (10 Char Max)"
      Top             =   240
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6600
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin TabDlg.SSTab tab1 
      Height          =   5475
      Left            =   0
      TabIndex        =   74
      Top             =   1005
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   9657
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&General       "
      TabPicture(0)   =   "diaCcust.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "z1(2)"
      Tab(0).Control(1)=   "z1(6)"
      Tab(0).Control(2)=   "z1(7)"
      Tab(0).Control(3)=   "z1(3)"
      Tab(0).Control(4)=   "z1(4)"
      Tab(0).Control(5)=   "z1(5)"
      Tab(0).Control(6)=   "z1(12)"
      Tab(0).Control(7)=   "z1(13)"
      Tab(0).Control(8)=   "z1(14)"
      Tab(0).Control(9)=   "z1(15)"
      Tab(0).Control(10)=   "z1(16)"
      Tab(0).Control(11)=   "z1(17)"
      Tab(0).Control(12)=   "z1(22)"
      Tab(0).Control(13)=   "z1(25)"
      Tab(0).Control(14)=   "z1(26)"
      Tab(0).Control(15)=   "lblSlp"
      Tab(0).Control(16)=   "lblDiv"
      Tab(0).Control(17)=   "lblReg"
      Tab(0).Control(18)=   "z1(64)"
      Tab(0).Control(19)=   "z1(51)"
      Tab(0).Control(20)=   "z1(68)"
      Tab(0).Control(21)=   "txtFax"
      Tab(0).Control(22)=   "txtPhn"
      Tab(0).Control(23)=   "txtAdr"
      Tab(0).Control(24)=   "txtCty"
      Tab(0).Control(25)=   "txtEml"
      Tab(0).Control(26)=   "txtCnt"
      Tab(0).Control(27)=   "txtWeb"
      Tab(0).Control(28)=   "txtAcd"
      Tab(0).Control(29)=   "txtExt"
      Tab(0).Control(30)=   "txtSln"
      Tab(0).Control(31)=   "txtZip"
      Tab(0).Control(32)=   "cmbDiv"
      Tab(0).Control(33)=   "cmbReg"
      Tab(0).Control(34)=   "cmbSlp"
      Tab(0).Control(35)=   "txtCph"
      Tab(0).Control(36)=   "cmbSte"
      Tab(0).Control(37)=   "txtCntr"
      Tab(0).Control(38)=   "txtCint"
      Tab(0).Control(39)=   "txtCintf"
      Tab(0).Control(40)=   "txtCcInt"
      Tab(0).Control(41)=   "txtBeg"
      Tab(0).Control(42)=   "txtNumber"
      Tab(0).ControlCount=   43
      TabCaption(1)   =   "&Shipping/Selling"
      TabPicture(1)   =   "diaCcust.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optEom"
      Tab(1).Control(1)=   "txtStcntr"
      Tab(1).Control(2)=   "cmbStSte"
      Tab(1).Control(3)=   "cmbTrm"
      Tab(1).Control(4)=   "txtStDis"
      Tab(1).Control(5)=   "txtStFra"
      Tab(1).Control(6)=   "txtStFrd"
      Tab(1).Control(7)=   "txtStVia"
      Tab(1).Control(8)=   "txtStFob"
      Tab(1).Control(9)=   "txtStSel"
      Tab(1).Control(10)=   "txtStZip"
      Tab(1).Control(11)=   "txtStCty"
      Tab(1).Control(12)=   "txtStAdr"
      Tab(1).Control(13)=   "txtStNme"
      Tab(1).Control(14)=   "z1(67)"
      Tab(1).Control(15)=   "z1(65)"
      Tab(1).Control(16)=   "lblTrm"
      Tab(1).Control(17)=   "z1(43)"
      Tab(1).Control(18)=   "z1(42)"
      Tab(1).Control(19)=   "z1(41)"
      Tab(1).Control(20)=   "z1(39)"
      Tab(1).Control(21)=   "z1(37)"
      Tab(1).Control(22)=   "z1(36)"
      Tab(1).Control(23)=   "z1(29)"
      Tab(1).Control(24)=   "z1(28)"
      Tab(1).Control(25)=   "z1(27)"
      Tab(1).Control(26)=   "z1(24)"
      Tab(1).Control(27)=   "z1(23)"
      Tab(1).Control(28)=   "z1(21)"
      Tab(1).Control(29)=   "z1(20)"
      Tab(1).Control(30)=   "z1(19)"
      Tab(1).Control(31)=   "z1(18)"
      Tab(1).ControlCount=   32
      TabCaption(2)   =   "&Billing           "
      TabPicture(2)   =   "diaCcust.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtBtcntr"
      Tab(2).Control(1)=   "cmbBtSte"
      Tab(2).Control(2)=   "txtType"
      Tab(2).Control(3)=   "txtBtCol"
      Tab(2).Control(4)=   "txtBtCty"
      Tab(2).Control(5)=   "txtBtZip"
      Tab(2).Control(6)=   "txtBtAdr"
      Tab(2).Control(7)=   "txtBtNme"
      Tab(2).Control(8)=   "z1(66)"
      Tab(2).Control(9)=   "z1(63)"
      Tab(2).Control(10)=   "z1(61)"
      Tab(2).Control(11)=   "z1(35)"
      Tab(2).Control(12)=   "z1(34)"
      Tab(2).Control(13)=   "z1(33)"
      Tab(2).Control(14)=   "z1(32)"
      Tab(2).Control(15)=   "z1(31)"
      Tab(2).Control(16)=   "z1(30)"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "&Quality Assurance"
      TabPicture(3)   =   "diaCcust.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "z1(8)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "z1(9)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "z1(10)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "z1(11)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "z1(62)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtQaFax"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtQaPhn"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "txtQaRep"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "txtQaEml"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "txtQaExt"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "txtQaIntp"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "txtQaIntf"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).ControlCount=   12
      TabCaption(4)   =   "&Receivables    "
      TabPicture(4)   =   "diaCcust.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "optEInv"
      Tab(4).Control(1)=   "txtApEmail"
      Tab(4).Control(2)=   "txtCreditLimit"
      Tab(4).Control(3)=   "cmbArSvc"
      Tab(4).Control(4)=   "txtArTxc"
      Tab(4).Control(5)=   "txtArIntf"
      Tab(4).Control(6)=   "txtArIntp"
      Tab(4).Control(7)=   "cmbTxw"
      Tab(4).Control(8)=   "cmbTxr"
      Tab(4).Control(9)=   "txtArExt"
      Tab(4).Control(10)=   "txtArCon"
      Tab(4).Control(11)=   "optCut"
      Tab(4).Control(12)=   "txtArDpp"
      Tab(4).Control(13)=   "txtArExm"
      Tab(4).Control(14)=   "txtArPct"
      Tab(4).Control(15)=   "txtArNet"
      Tab(4).Control(16)=   "txtArDay"
      Tab(4).Control(17)=   "txtArDis"
      Tab(4).Control(18)=   "txtArPhn"
      Tab(4).Control(19)=   "txtArFax"
      Tab(4).Control(20)=   "z1(71)"
      Tab(4).Control(21)=   "z1(72)"
      Tab(4).Control(22)=   "z1(69)"
      Tab(4).Control(23)=   "z1(60)"
      Tab(4).Control(24)=   "z1(59)"
      Tab(4).Control(25)=   "z1(58)"
      Tab(4).Control(26)=   "z1(57)"
      Tab(4).Control(27)=   "z1(56)"
      Tab(4).Control(28)=   "z1(55)"
      Tab(4).Control(29)=   "z1(54)"
      Tab(4).Control(30)=   "z1(53)"
      Tab(4).Control(31)=   "z1(52)"
      Tab(4).Control(32)=   "z1(50)"
      Tab(4).Control(33)=   "z1(49)"
      Tab(4).Control(34)=   "z1(48)"
      Tab(4).Control(35)=   "z1(47)"
      Tab(4).Control(36)=   "z1(46)"
      Tab(4).Control(37)=   "z1(45)"
      Tab(4).Control(38)=   "z1(44)"
      Tab(4).Control(39)=   "z1(40)"
      Tab(4).Control(40)=   "z1(38)"
      Tab(4).ControlCount=   41
      Begin VB.CheckBox optEInv 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   252
         Left            =   -73560
         TabIndex        =   152
         ToolTipText     =   "Electronic Invoicing"
         Top             =   4620
         Width           =   735
      End
      Begin VB.TextBox txtApEmail 
         Height          =   285
         Left            =   -73560
         TabIndex        =   151
         Tag             =   "2"
         ToolTipText     =   "Double Click Here To Send E-Mail (Requires An Entry)"
         Top             =   4980
         Width           =   5220
      End
      Begin VB.TextBox txtCreditLimit 
         Height          =   285
         Left            =   -73560
         TabIndex        =   53
         Tag             =   "1"
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox txtNumber 
         Height          =   285
         Left            =   -68400
         TabIndex        =   3
         Tag             =   "3"
         Top             =   660
         Width           =   735
      End
      Begin VB.CheckBox optEom 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   -69960
         TabIndex        =   30
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox txtBeg 
         Height          =   315
         Left            =   -73560
         TabIndex        =   2
         Tag             =   "4"
         Top             =   700
         Width           =   1095
      End
      Begin VB.ComboBox cmbArSvc 
         Height          =   315
         Left            =   -73560
         TabIndex        =   57
         Tag             =   "3"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox txtArTxc 
         Height          =   315
         Left            =   -73560
         Sorted          =   -1  'True
         TabIndex        =   62
         Tag             =   "3"
         ToolTipText     =   "Select Or Enter Customer (10 Char Max)"
         Top             =   2760
         Width           =   1680
      End
      Begin VB.TextBox txtArIntf 
         Height          =   285
         Left            =   -70080
         TabIndex        =   69
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   4200
         Width           =   475
      End
      Begin VB.TextBox txtArIntp 
         Height          =   285
         Left            =   -73560
         TabIndex        =   66
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   4200
         Width           =   475
      End
      Begin VB.TextBox txtQaIntf 
         Height          =   285
         Left            =   4920
         TabIndex        =   50
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   1440
         Width           =   475
      End
      Begin VB.TextBox txtQaIntp 
         Height          =   285
         Left            =   1440
         TabIndex        =   47
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   1425
         Width           =   475
      End
      Begin VB.TextBox txtCcInt 
         Height          =   285
         Left            =   -70560
         TabIndex        =   15
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   2745
         Width           =   475
      End
      Begin VB.TextBox txtCintf 
         Height          =   285
         Left            =   -70560
         TabIndex        =   12
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   2420
         Width           =   475
      End
      Begin VB.TextBox txtCint 
         Height          =   285
         Left            =   -73560
         TabIndex        =   9
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   2420
         Width           =   475
      End
      Begin VB.ComboBox cmbTxw 
         Height          =   315
         Left            =   -70080
         Sorted          =   -1  'True
         TabIndex        =   64
         Tag             =   "3"
         ToolTipText     =   "Select Or Enter Customer (10 Char Max)"
         Top             =   3360
         Width           =   1680
      End
      Begin VB.ComboBox cmbTxr 
         Height          =   315
         Left            =   -73560
         Sorted          =   -1  'True
         TabIndex        =   63
         Tag             =   "3"
         ToolTipText     =   "Select Or Enter Customer (10 Char Max)"
         Top             =   3360
         Width           =   1680
      End
      Begin VB.ComboBox txtBtcntr 
         Height          =   315
         Left            =   -73560
         Sorted          =   -1  'True
         TabIndex        =   43
         Tag             =   "3"
         ToolTipText     =   "Select Country From The List"
         Top             =   2625
         Width           =   1935
      End
      Begin VB.ComboBox txtStcntr 
         Height          =   315
         Left            =   -73560
         Sorted          =   -1  'True
         TabIndex        =   29
         Tag             =   "3"
         ToolTipText     =   "Select Country From The List"
         Top             =   2505
         Width           =   1935
      End
      Begin VB.ComboBox txtCntr 
         Height          =   315
         Left            =   -73560
         TabIndex        =   8
         Tag             =   "3"
         ToolTipText     =   "Select Country From The List"
         Top             =   2060
         Width           =   2055
      End
      Begin VB.ComboBox cmbBtSte 
         Height          =   315
         Left            =   -70920
         Sorted          =   -1  'True
         TabIndex        =   41
         Tag             =   "3"
         Top             =   2265
         Width           =   735
      End
      Begin VB.ComboBox cmbStSte 
         Height          =   315
         Left            =   -70680
         TabIndex        =   27
         Tag             =   "3"
         Top             =   2145
         Width           =   735
      End
      Begin VB.ComboBox cmbSte 
         Height          =   315
         Left            =   -70560
         TabIndex        =   6
         Tag             =   "3"
         Top             =   1740
         Width           =   735
      End
      Begin MSMask.MaskEdBox txtCph 
         Height          =   285
         Left            =   -70080
         TabIndex        =   16
         Top             =   2745
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbTrm 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -73560
         TabIndex        =   35
         Tag             =   "8"
         ToolTipText     =   "Select Terms From List"
         Top             =   3585
         Width           =   660
      End
      Begin VB.ComboBox cmbSlp 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -73560
         Sorted          =   -1  'True
         TabIndex        =   18
         Tag             =   "3"
         ToolTipText     =   "Select Salesperson From List"
         Top             =   3090
         Width           =   975
      End
      Begin VB.ComboBox cmbReg 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -70320
         Sorted          =   -1  'True
         TabIndex        =   21
         Tag             =   "8"
         ToolTipText     =   "Select Region From List"
         Top             =   3465
         Width           =   780
      End
      Begin VB.ComboBox cmbDiv 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -73560
         Sorted          =   -1  'True
         TabIndex        =   20
         Tag             =   "8"
         ToolTipText     =   "Select Division From List"
         Top             =   3465
         Width           =   860
      End
      Begin VB.TextBox txtType 
         Height          =   285
         Left            =   -73560
         TabIndex        =   44
         Tag             =   "3"
         Top             =   2985
         Width           =   375
      End
      Begin VB.TextBox txtQaExt 
         Height          =   285
         Left            =   3720
         TabIndex        =   49
         Tag             =   "1"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtBtCol 
         Height          =   1215
         Left            =   -73560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Tag             =   "9"
         ToolTipText     =   "Comments (Max 1024 Chars)"
         Top             =   3345
         Width           =   3495
      End
      Begin VB.TextBox txtArExt 
         Height          =   285
         Left            =   -71160
         TabIndex        =   68
         Tag             =   "1"
         Top             =   4200
         Width           =   495
      End
      Begin VB.TextBox txtArCon 
         Height          =   285
         Left            =   -73560
         TabIndex        =   65
         Tag             =   "2"
         Top             =   3760
         Width           =   2055
      End
      Begin VB.CheckBox optCut 
         Alignment       =   1  'Right Justify
         Caption         =   "Cut Off?"
         Height          =   255
         Left            =   -69840
         TabIndex        =   59
         ToolTipText     =   "Customer's Credit Has Been Halted"
         Top             =   1665
         Width           =   1095
      End
      Begin VB.TextBox txtArDpp 
         Height          =   285
         Left            =   -73560
         TabIndex        =   61
         Tag             =   "3"
         Top             =   2400
         Width           =   3315
      End
      Begin VB.TextBox txtArExm 
         Height          =   285
         Left            =   -73560
         TabIndex        =   60
         Tag             =   "3"
         Top             =   2040
         Width           =   3315
      End
      Begin VB.TextBox txtArPct 
         Height          =   285
         Left            =   -71040
         TabIndex        =   58
         Tag             =   "1"
         Top             =   1665
         Width           =   735
      End
      Begin VB.TextBox txtArNet 
         Height          =   285
         Left            =   -71340
         TabIndex        =   56
         Tag             =   "1"
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox txtArDay 
         Height          =   285
         Left            =   -72480
         TabIndex        =   55
         Tag             =   "1"
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox txtArDis 
         Height          =   285
         Left            =   -73560
         TabIndex        =   54
         Tag             =   "1"
         Top             =   1020
         Width           =   735
      End
      Begin VB.TextBox txtStDis 
         Height          =   285
         Left            =   -69960
         TabIndex        =   36
         Tag             =   "1"
         Top             =   3585
         Width           =   615
      End
      Begin VB.TextBox txtBtCty 
         Height          =   285
         Left            =   -73560
         TabIndex        =   40
         Tag             =   "2"
         Top             =   2265
         Width           =   1935
      End
      Begin VB.TextBox txtStFra 
         Height          =   285
         Left            =   -69960
         TabIndex        =   34
         Tag             =   "1"
         Top             =   3225
         Width           =   615
      End
      Begin VB.TextBox txtStFrd 
         Height          =   285
         Left            =   -73560
         TabIndex        =   33
         Tag             =   "1"
         Top             =   3225
         Width           =   615
      End
      Begin VB.TextBox txtBtZip 
         Height          =   285
         Left            =   -69720
         TabIndex        =   42
         Tag             =   "3"
         Top             =   2265
         Width           =   1335
      End
      Begin VB.TextBox txtBtAdr 
         Height          =   765
         Left            =   -73560
         MultiLine       =   -1  'True
         TabIndex        =   39
         Tag             =   "9"
         Top             =   1425
         Width           =   3495
      End
      Begin VB.TextBox txtBtNme 
         Height          =   285
         Left            =   -73560
         TabIndex        =   38
         Tag             =   "2"
         Top             =   1065
         Width           =   3475
      End
      Begin VB.TextBox txtStVia 
         Height          =   285
         Left            =   -69960
         TabIndex        =   32
         Tag             =   "3"
         Top             =   2865
         Width           =   1935
      End
      Begin VB.TextBox txtStFob 
         Height          =   285
         Left            =   -73560
         TabIndex        =   31
         Tag             =   "3"
         Top             =   2865
         Width           =   1575
      End
      Begin VB.TextBox txtStSel 
         Height          =   885
         Left            =   -73560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Tag             =   "9"
         ToolTipText     =   "Comments (Max 1024 Chars)"
         Top             =   3945
         Width           =   3495
      End
      Begin VB.TextBox txtZip 
         Height          =   285
         Left            =   -69000
         TabIndex        =   7
         Tag             =   "3"
         Top             =   1740
         Width           =   1335
      End
      Begin VB.TextBox txtStZip 
         Height          =   285
         Left            =   -69360
         TabIndex        =   28
         Tag             =   "3"
         Top             =   2145
         Width           =   1335
      End
      Begin VB.TextBox txtStCty 
         Height          =   285
         Left            =   -73560
         TabIndex        =   26
         Tag             =   "2"
         Top             =   2145
         Width           =   1935
      End
      Begin VB.TextBox txtStAdr 
         Height          =   645
         Left            =   -73560
         MultiLine       =   -1  'True
         TabIndex        =   25
         Tag             =   "9"
         Top             =   1425
         Width           =   3495
      End
      Begin VB.TextBox txtStNme 
         Height          =   285
         Left            =   -73560
         TabIndex        =   24
         Tag             =   "2"
         Top             =   1065
         Width           =   3475
      End
      Begin VB.TextBox txtSln 
         Height          =   285
         Left            =   -73560
         TabIndex        =   19
         Tag             =   "1"
         Top             =   5040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtExt 
         Height          =   285
         Left            =   -68280
         TabIndex        =   17
         Top             =   2745
         Width           =   615
      End
      Begin VB.TextBox txtAcd 
         Height          =   285
         Left            =   -71760
         TabIndex        =   11
         Tag             =   "1"
         Top             =   2420
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtWeb 
         Height          =   285
         Left            =   -73560
         TabIndex        =   23
         Tag             =   "2"
         ToolTipText     =   "You May Double Click To Open A Page That Has Been Entered"
         Top             =   4185
         Width           =   5220
      End
      Begin VB.TextBox txtCnt 
         Height          =   285
         Left            =   -73560
         TabIndex        =   14
         Top             =   2745
         Width           =   2205
      End
      Begin VB.TextBox txtQaEml 
         Height          =   285
         Left            =   1440
         TabIndex        =   52
         Tag             =   "2"
         ToolTipText     =   "Click Here To Send E-Mail (Requires An Entry)"
         Top             =   1785
         Width           =   5340
      End
      Begin VB.TextBox txtEml 
         Height          =   285
         Left            =   -73560
         TabIndex        =   22
         Tag             =   "2"
         ToolTipText     =   "Double Click Here To Send E-Mail (Requires An Entry)"
         Top             =   3825
         Width           =   5220
      End
      Begin VB.TextBox txtCty 
         Height          =   285
         Left            =   -73560
         TabIndex        =   5
         Tag             =   "2"
         Top             =   1725
         Width           =   2085
      End
      Begin VB.TextBox txtAdr 
         Height          =   615
         Left            =   -73560
         MultiLine       =   -1  'True
         TabIndex        =   4
         Tag             =   "9"
         Top             =   1065
         Width           =   3475
      End
      Begin VB.TextBox txtQaRep 
         Height          =   285
         Left            =   1440
         TabIndex        =   46
         Tag             =   "2"
         Top             =   1065
         Width           =   2085
      End
      Begin MSMask.MaskEdBox txtPhn 
         Height          =   285
         Left            =   -73065
         TabIndex        =   10
         Top             =   2420
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFax 
         Height          =   285
         Left            =   -70080
         TabIndex        =   13
         Top             =   2420
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtQaPhn 
         Height          =   285
         Left            =   1920
         TabIndex        =   48
         Top             =   1425
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   14
         Mask            =   """###-###-####"""
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtArPhn 
         Height          =   285
         Left            =   -73080
         TabIndex        =   67
         Top             =   4200
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   14
         Mask            =   """###-###-####"""
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtArFax 
         Height          =   285
         Left            =   -69600
         TabIndex        =   70
         Top             =   4200
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   14
         Mask            =   """###-###-####"""
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtQaFax 
         Height          =   285
         Left            =   5400
         TabIndex        =   51
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   14
         Mask            =   """###-###-####"""
         PromptChar      =   "_"
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "E Invoicing"
         Height          =   285
         Index           =   71
         Left            =   -74880
         TabIndex        =   154
         ToolTipText     =   "Electronic Invoicing"
         Top             =   4620
         Width           =   1185
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "AP EMail"
         Height          =   285
         Index           =   72
         Left            =   -74880
         TabIndex        =   153
         Top             =   4980
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit"
         Height          =   285
         Index           =   69
         Left            =   -74880
         TabIndex        =   150
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Number"
         Height          =   285
         Index           =   68
         Left            =   -69900
         TabIndex        =   149
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "EOM"
         Height          =   285
         Index           =   67
         Left            =   -71280
         TabIndex        =   148
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Since"
         Height          =   285
         Index           =   51
         Left            =   -74880
         TabIndex        =   147
         Top             =   700
         Width           =   1425
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         Height          =   285
         Index           =   66
         Left            =   -74880
         TabIndex        =   144
         Top             =   2625
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         Height          =   285
         Index           =   65
         Left            =   -74880
         TabIndex        =   143
         Top             =   2505
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         Height          =   285
         Index           =   64
         Left            =   -74880
         TabIndex        =   142
         Top             =   2025
         Width           =   915
      End
      Begin VB.Label lblTrm 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -72840
         TabIndex        =   140
         Top             =   3585
         Width           =   1815
      End
      Begin VB.Label lblReg 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -69480
         TabIndex        =   139
         Top             =   3465
         Width           =   1455
      End
      Begin VB.Label lblDiv 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -72525
         TabIndex        =   138
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label lblSlp 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -72525
         TabIndex        =   137
         Top             =   3090
         Width           =   3060
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Type"
         Height          =   285
         Index           =   63
         Left            =   -74880
         TabIndex        =   135
         Top             =   2985
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext"
         Height          =   285
         Index           =   62
         Left            =   3360
         TabIndex        =   134
         Top             =   1425
         Width           =   465
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes:"
         Height          =   285
         Index           =   61
         Left            =   -74880
         TabIndex        =   133
         Top             =   3345
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Days"
         Height          =   285
         Index           =   60
         Left            =   -70560
         TabIndex        =   132
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   285
         Index           =   59
         Left            =   -70560
         TabIndex        =   131
         Top             =   4200
         Width           =   465
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext"
         Height          =   285
         Index           =   58
         Left            =   -71520
         TabIndex        =   130
         Top             =   4200
         Width           =   465
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   285
         Index           =   57
         Left            =   -74520
         TabIndex        =   129
         Top             =   4200
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         Height          =   285
         Index           =   56
         Left            =   -74880
         TabIndex        =   128
         Top             =   3825
         Width           =   1185
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Wholesale"
         Height          =   285
         Index           =   55
         Left            =   -71520
         TabIndex        =   127
         Top             =   3360
         Width           =   945
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Retail"
         Height          =   285
         Index           =   54
         Left            =   -74520
         TabIndex        =   126
         Top             =   3360
         Width           =   945
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "B and O Tax Codes:"
         Height          =   285
         Index           =   53
         Left            =   -74880
         TabIndex        =   125
         Top             =   3120
         Width           =   1905
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Code"
         Height          =   285
         Index           =   52
         Left            =   -74880
         TabIndex        =   124
         Top             =   2760
         Width           =   945
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Direct Pay Per #"
         Height          =   285
         Index           =   50
         Left            =   -74880
         TabIndex        =   123
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Exempt #"
         Height          =   285
         Index           =   49
         Left            =   -74880
         TabIndex        =   122
         Top             =   2040
         Width           =   1185
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   285
         Index           =   48
         Left            =   -70200
         TabIndex        =   121
         Top             =   1665
         Width           =   315
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   285
         Index           =   47
         Left            =   -71520
         TabIndex        =   120
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Account"
         Height          =   285
         Index           =   46
         Left            =   -74520
         TabIndex        =   119
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Service Chgs:"
         Height          =   285
         Index           =   45
         Left            =   -74880
         TabIndex        =   118
         Top             =   1425
         Width           =   1185
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Net"
         Height          =   285
         Index           =   44
         Left            =   -71760
         TabIndex        =   117
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "% "
         Height          =   285
         Index           =   40
         Left            =   -72720
         TabIndex        =   116
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   285
         Index           =   43
         Left            =   -69840
         TabIndex        =   115
         Top             =   3585
         Width           =   315
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   285
         Index           =   42
         Left            =   -69840
         TabIndex        =   114
         Top             =   3225
         Width           =   315
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   285
         Index           =   41
         Left            =   -70920
         TabIndex        =   113
         Top             =   3585
         Width           =   1095
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping Terms"
         Height          =   285
         Index           =   39
         Left            =   -74880
         TabIndex        =   112
         Top             =   3585
         Width           =   1275
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Terms"
         Height          =   285
         Index           =   38
         Left            =   -74880
         TabIndex        =   111
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Frgt Allow"
         Height          =   285
         Index           =   37
         Left            =   -71280
         TabIndex        =   110
         Top             =   3225
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Freight Days"
         Height          =   285
         Index           =   36
         Left            =   -74880
         TabIndex        =   109
         Top             =   3225
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   285
         Index           =   35
         Left            =   -74880
         TabIndex        =   108
         Top             =   1065
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   285
         Index           =   34
         Left            =   -74880
         TabIndex        =   107
         Top             =   1425
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   285
         Index           =   33
         Left            =   -74880
         TabIndex        =   106
         Top             =   2265
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         Height          =   285
         Index           =   32
         Left            =   -71520
         TabIndex        =   105
         Top             =   2265
         Width           =   795
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         Height          =   285
         Index           =   31
         Left            =   -70080
         TabIndex        =   104
         Top             =   2265
         Width           =   555
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill To:"
         Height          =   285
         Index           =   30
         Left            =   -74880
         TabIndex        =   103
         Top             =   825
         Width           =   795
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship Via"
         Height          =   285
         Index           =   29
         Left            =   -71280
         TabIndex        =   102
         Top             =   2865
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "FOB"
         Height          =   285
         Index           =   28
         Left            =   -74880
         TabIndex        =   101
         Top             =   2865
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Notes:"
         Height          =   285
         Index           =   27
         Left            =   -74880
         TabIndex        =   100
         Top             =   3945
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         Height          =   285
         Index           =   26
         Left            =   -71040
         TabIndex        =   99
         Top             =   3465
         Width           =   705
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         Height          =   285
         Index           =   25
         Left            =   -74880
         TabIndex        =   98
         Top             =   3465
         Width           =   1065
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To:"
         Height          =   285
         Index           =   24
         Left            =   -74880
         TabIndex        =   97
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         Height          =   285
         Index           =   23
         Left            =   -69840
         TabIndex        =   96
         Top             =   2145
         Width           =   555
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         Height          =   285
         Index           =   22
         Left            =   -69720
         TabIndex        =   95
         Top             =   1740
         Width           =   555
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         Height          =   285
         Index           =   21
         Left            =   -71280
         TabIndex        =   94
         Top             =   2145
         Width           =   795
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   285
         Index           =   20
         Left            =   -74880
         TabIndex        =   93
         Top             =   2145
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   285
         Index           =   19
         Left            =   -74880
         TabIndex        =   92
         Top             =   1425
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   285
         Index           =   18
         Left            =   -74880
         TabIndex        =   91
         Top             =   1065
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Num"
         Height          =   285
         Index           =   17
         Left            =   -74520
         TabIndex        =   90
         Top             =   5040
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Salesperson"
         Height          =   285
         Index           =   16
         Left            =   -74880
         TabIndex        =   89
         Top             =   3105
         Width           =   1065
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   285
         Index           =   15
         Left            =   -71160
         TabIndex        =   88
         Top             =   2745
         Width           =   705
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext"
         Height          =   285
         Index           =   14
         Left            =   -68640
         TabIndex        =   87
         Top             =   2745
         Width           =   465
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Web Site"
         Height          =   285
         Index           =   13
         Left            =   -74880
         TabIndex        =   86
         Top             =   4185
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         Height          =   285
         Index           =   12
         Left            =   -74880
         TabIndex        =   85
         Top             =   2745
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
         Height          =   285
         Index           =   5
         Left            =   -74880
         TabIndex        =   84
         Top             =   3825
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         Height          =   285
         Index           =   4
         Left            =   -71160
         TabIndex        =   83
         Top             =   1785
         Width           =   675
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   285
         Index           =   3
         Left            =   -74880
         TabIndex        =   82
         Top             =   1725
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   285
         Index           =   7
         Left            =   -71160
         TabIndex        =   81
         Top             =   2420
         Width           =   705
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   285
         Index           =   6
         Left            =   -74880
         TabIndex        =   80
         Top             =   2420
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   165
         Index           =   2
         Left            =   -74880
         TabIndex        =   79
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
         Height          =   285
         Index           =   11
         Left            =   120
         TabIndex        =   78
         Top             =   1785
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   77
         Top             =   1425
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   285
         Index           =   9
         Left            =   4440
         TabIndex        =   76
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quality Rep"
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   75
         Top             =   1080
         Width           =   1185
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   136
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
      PictureUp       =   "diaCcust.frx":0396
      PictureDn       =   "diaCcust.frx":04DC
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   6660
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6525
      FormDesignWidth =   7620
   End
   Begin VB.Label lblBook 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Price Book"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   145
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   73
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   72
      Top             =   280
      Width           =   1425
   End
End
Attribute VB_Name = "diaCcust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Accounts (SO/B&O) 2/13/03
'10/8/03 International Calling Codes
'11/5/03 Customer Since
'2/4/04 Email
'6/2/04 Reformated entire form
Option Explicit
Dim rdoCst As ADODB.Recordset
Dim bCancel As Byte
Dim bGoodCust As Byte
Dim bNewCust As Byte
Dim bOnLoad As Byte
Dim iTimer As Integer
Dim sState As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtChange() As New EsiKeyBd

Private Sub ClearCombos()
   Dim iControl As Integer
   For iControl = 0 To Controls.count - 1
      If TypeOf Controls(iControl) Is ComboBox Then _
                         Controls(iControl).SelLength = 0
   Next
   
End Sub

Private Sub FillSales()
   On Error GoTo DiaErr1
   sSql = "SELECT SPNUMBER FROM SprsTable "
   LoadComboBox cmbSlp, -1
   Exit Sub
   
DiaErr1:
   sProcName = "fillsales"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub ManageBoxes(bEnabled As Byte)
   tab1.enabled = bEnabled
   
End Sub

Private Sub cmbArSvc_LostFocus()
   cmbArSvc = CheckLen(cmbArSvc, 12)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUSVCACCT = "" & cmbArSvc
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub cmbBtSte_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub cmbBtSte_LostFocus()
   cmbBtSte = CheckLen(cmbBtSte, 4)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUBSTATE = "" & cmbBtSte
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub cmbCst_Click()
   bGoodCust = GetCustomer(0)
   'ManageBoxes False
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If bCancel Then Exit Sub
   If Len(cmbCst) = 0 Then
      cmdCan.SetFocus
      Exit Sub
   End If
   bGoodCust = GetCustomer(1)
   If bGoodCust = 0 Then
      ManageBoxes False
      AddCustomer
   Else
      ManageBoxes True
   End If
   
End Sub

Private Sub cmbDiv_Click()
   GetDivision
   
End Sub

Private Sub cmbDiv_LostFocus()
   cmbDiv = CheckLen(cmbDiv, 4)
   If Trim(cmbDiv) = "" Then
      If cmbDiv.ListCount > 0 Then
         Beep
         cmbDiv = cmbDiv.List(0)
      End If
   End If
   GetDivision
   If Len(cmbDiv) Then
      If bGoodCust = 1 Then
         On Error Resume Next
         rdoCst!CUDIVISION = "" & cmbDiv
         rdoCst.Update
         If Err > 0 Then bGoodCust = GetCustomer(0)
      End If
   End If
   
End Sub


Private Sub cmbReg_Click()
   GetRegion
   
End Sub

Private Sub cmbReg_LostFocus()
   cmbReg = CheckLen(cmbReg, 2)
   If Trim(cmbReg) = "" Then
      If cmbReg.ListCount > 0 Then
         Beep
         cmbReg = cmbReg.List(0)
      End If
   End If
   GetRegion
   If Len(cmbReg) Then
      If bGoodCust = 1 Then
         On Error Resume Next
         rdoCst!CUREGION = "" & cmbReg
         rdoCst.Update
         If Err > 0 Then bGoodCust = GetCustomer(0)
      End If
   End If
   
End Sub


Private Sub cmbSlp_Click()
   GetSalesPerson
   
End Sub

Private Sub cmbSlp_LostFocus()
   Dim b As Byte
   Dim iRow As Integer
   cmbSlp = CheckLen(cmbSlp, 4)
   b = 1
   If Trim(cmbSlp) = "" Then
      If cmbSlp.ListCount > 0 Then
         Beep
         cmbSlp = cmbSlp.List(0)
      End If
   Else
      b = 0
      If cmbSlp.ListCount > 0 Then
         For iRow = 0 To cmbSlp.ListCount - 1
            If cmbSlp = cmbSlp.List(iRow) Then b = 1
         Next
      Else
         b = 1
      End If
   End If
   If b = 0 Then
      Beep
      cmbSlp = cmbSlp.List(0)
   End If
   GetSalesPerson
   If bGoodCust = 1 Then
      If Len(Trim(cmbSlp)) Then
         On Error Resume Next
         rdoCst!CUSALESMAN = "" & cmbSlp
         rdoCst.Update
         If Err > 0 Then bGoodCust = GetCustomer(0)
      End If
   End If
   
End Sub


Private Sub cmbSte_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub cmbSte_LostFocus()
   cmbSte = CheckLen(cmbSte, 4)
   If bNewCust = 1 Then
      cmbStSte = cmbSte
      cmbBtSte = cmbSte
   End If
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUSTATE = "" & cmbSte
      rdoCst!CUSTSTATE = "" & cmbStSte
      rdoCst!CUBSTATE = "" & cmbBtSte
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub cmbStSte_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub cmbStSte_LostFocus()
   cmbStSte = CheckLen(cmbStSte, 4)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUSTSTATE = "" & cmbStSte
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub cmbTrm_Click()
   GetTerms
   
End Sub

Private Sub cmbTrm_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub cmbTrm_LostFocus()
   cmbTrm = CheckLen(cmbTrm, 2)
   If Trim(cmbTrm) = "" Then
      If cmbTrm.ListCount > 0 Then
         Beep
         cmbTrm = cmbTrm.List(0)
      End If
   End If
   GetTerms
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUSTERMS = "" & cmbTrm
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub





Private Sub cmbTxr_LostFocus()
   cmbTxr = CheckLen(cmbTxr, 8)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUBORTAXCODE = "" & cmbTxr
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub cmbTxw_LostFocus()
   cmbTxw = CheckLen(cmbTxw, 8)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUBOWTAXCODE = "" & cmbTxw
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub cmdCan_Click()
   Timer1(0).enabled = False
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = True
   
End Sub


Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs2103"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub




Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      MouseCursor 13
      bOnLoad = 0
      bNewCust = 0
      FillScAccount
      FillTaxCodes
      FillCountries
      FillStates Me
      cmbSlp = ""
      sState = cmbSte
      FillSales
      FillCustomers Me
      If cUR.CurrentCustomer <> "" Then cmbCst = cUR.CurrentCustomer
      FillDivisions Me
      FillRegions Me
      FillTerms Me
      ManageBoxes False
      bGoodCust = GetCustomer(0)
   End If
   If optVewAcct.Value = vbChecked Then Unload VewAcct
   MouseCursor 0
   
End Sub

Private Sub Form_Initialize()
   tab1.Tab = 0
   
End Sub

Private Sub Form_Load()
   Dim cNewSize As Currency
   MouseCursor 13
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bDataHasChanged And bGoodCust = 1 Then UpdateCustomerSet True
   If Len(cmbCst) Then
      If bGoodCust = 1 Then cUR.CurrentCustomer = cmbCst
      SaveCurrentSelections
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)

   If optVewAcct.Value = vbChecked Then Unload VewAcct
   Set rdoCst = Nothing
   FormUnload
   Set diaCcust = Nothing
   
End Sub

Private Function GetCustomer(bSelect As Byte) As Byte
   Dim sCustRef As String
   Dim sPhone As String
   If bDataHasChanged And bGoodCust = 1 Then UpdateCustomerSet
   MouseCursor 13
   sCustRef = Compress(cmbCst)
   On Error GoTo DiaErr1
   If bSelect Then
      tab1.Tab = 0
      ManageBoxes False
   End If
   GetCustomer = 0
   bGoodCust = 0
   ClearBoxes
   sSql = "SELECT * FROM CustTable WHERE CUREF='" & sCustRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst, ES_KEYSET)
   If bSqlRows Then
      With rdoCst
         cmbCst = "" & Trim(!CUNICKNAME)
         'lblNum = Format(!CUNUMBER, "#####0")
         txtNumber = Format(!CUNUMBER, "#####0")
         'General
         txtNme = "" & Trim(!CUNAME)
         If Len(Trim(!CUADR)) = 0 And Len(Trim(!CUCITY)) = 0 Then bNewCust = 1 _
                Else bNewCust = 0
         txtAdr = "" & Trim(!CUADR)
         txtCty = "" & Trim(!CUCITY)
         If bNewCust = 1 Then
            If Trim(!CUSTATE) = "" Then cmbSte = sState
         Else
            cmbSte = "" & Trim(!CUSTATE)
         End If
         txtZip = "" & Trim(!CUZIP)
         If Len(Trim("" & !CUTELEPHONE)) Then txtPhn = "" & Trim(!CUTELEPHONE)
         sPhone = "" & Trim(!CUFAX)
         If Len(sPhone) Then txtFax = "" & Trim(!CUFAX)
         txtCnt = "" & Trim(!CUCCONTACT)
         sPhone = "" & Trim(!CUCPHONE)
         If Not IsNull(!CUSINCE) Then txtBeg = Format(!CUSINCE, "mm/dd/yy")
         If Len(sPhone) Then txtCph = "" & Trim(!CUCPHONE)
         If !CUCEXT > 0 Then
            txtExt = Format(!CUCEXT, "###0")
         Else
            txtExt = ""
         End If
         cmbSlp = "" & Trim(!CUSALESMAN)
         If !CUSLSMAN > 0 Then
            txtSln = Format(0 + !CUSLSMAN, "000000")
         Else
            txtSln = ""
         End If
         cmbDiv = "" & Trim(!CUDIVISION)
         cmbReg = "" & Trim(!CUREGION)
         txtEml = "" & Trim(!CUEMAIL)
         txtWeb = "" & Trim(!CUWEB)
         '5/24/04
         optEom.Value = !CUBEOM
         'Shipping
         txtStNme = "" & Trim(!CUSTNAME)
         txtStAdr = "" & Trim(!CUSTADR)
         txtStCty = "" & Trim(!CUSTCITY)
         cmbStSte = "" & Trim(!CUSTSTATE)
         txtStZip = "" & Trim(!CUSTZIP)
         txtStFob = "" & Trim(!CUFOB)
         txtStVia = "" & Trim(!CUVIA)
         txtStFrd = Format(!CUFRTDAYS, "##0")
         txtStFra = Format(!CUFRTALLOW, "##0.000")
         cmbTrm = "" & Trim(!CUSTERMS)
         txtStDis = Format(!CUDISCOUNT, "##0.000")
         txtStSel = "" & Trim(!CUSEL)
         'If bOnload Then
         'Billing
         txtBtNme = "" & Trim(!CUBTNAME)
         txtBtAdr = "" & Trim(!CUBTADR)
         txtBtCty = "" & Trim(!CUBCITY)
         cmbBtSte = "" & Trim(!CUBSTATE)
         txtBtZip = "" & Trim(!CUBZIP)
         txtType = "" & Trim(!CUTYPE)
         txtBtCol = "" & Trim(!CUCOL)
         'QA
         txtQaRep = "" & Trim(!CUQAREP)
         If Len(Trim(!CUQAPHONE)) Then txtQaPhn = "" & Trim(!CUQAPHONE)
         If !CUQAPHONEEXT > 0 Then
            txtQaExt = Format(!CUQAPHONEEXT, "###0")
         Else
            txtQaExt = ""
         End If
         If Len(Trim(!CUQAFAX)) Then txtQaFax = "" & Trim(!CUQAFAX)
         txtQaEml = "" & Trim(!CUQAEMAIL)
         'Receivables
         txtArDis = Format(!CUARDISC, "##0.000")
         txtArDay = Format(!CUDAYS, "##0")
         txtArNet = Format(!CUNETDAYS, "##0")
         cmbArSvc = "" & Trim(!CUSVCACCT)
         txtArPct = Format(!CUARPCT, "##0.000")
         optCut.Value = !CUCUTOFF
         txtArExm = "" & Trim(!CUTAXEXEMPT)
         txtArDpp = "" & Trim(!CUDPP)
         'Removed 10/10/02
         'txtArTxs = "" & Trim(!CUTAXSTATE)
         txtArTxc = "" & Trim(!CUTAXCODE)
         cmbTxr = "" & Trim(!CUBORTAXCODE)
         cmbTxw = "" & Trim(!CUBOWTAXCODE)
         txtArCon = "" & Trim(!CUAPCONTACT)
         If Len(Trim(!CUAPPHONE)) Then txtArPhn = "" & Trim(!CUAPPHONE)
         txtArExt = Format(!CUAPPHONEEXT, "###0")
         If Len(Trim(!CUAPFAX)) Then txtArFax = "" & Trim(!CUAPFAX)
         bOnLoad = 0
         'country 2/15/00
         txtCntr = "" & Trim(!CUCOUNTRY)
         txtStcntr = "" & Trim(!CUSTCOUNTRY)
         txtBtcntr = "" & Trim(!CUBCOUNTRY)
         lblBook = "" & Trim(!CUPRICEBOOK)
         '10/8/03 International Phones
         txtCint = "" & Trim(!CUINTPHONE)
         txtCintf = "" & Trim(!CUINTFAX)
         txtCcInt = "" & Trim(!CUCINTPHONE)
         txtQaIntp = "" & Trim(!CUQAINTPHONE)
         txtQaIntf = "" & Trim(!CUQAINTFAX)
         txtArIntp = "" & Trim(!CUAPINTPHONE)
         txtArIntf = "" & Trim(!CUAPINTFAX)
         
         txtCreditLimit = Format(!CUCREDITLIMIT, "###,###,##0.00")
         
         '12/8/17
         optEInv.Value = !CUEINVOICING
         txtApEmail = "" & Trim(!CUAPEMAIL)

      End With
      txtPhn.Mask = "###-###-####"
      txtFax.Mask = "###-###-####"
      txtCph.Mask = "###-###-####"
      txtQaPhn.Mask = "###-###-####"
      txtQaFax.Mask = "###-###-####"
      txtArPhn.Mask = "###-###-####"
      txtArFax.Mask = "###-###-####"
      If bSelect = 0 Then ManageBoxes True
      GetDivision
      GetRegion
      GetSalesPerson
      GetTerms
      Err.Clear
      GetCustomer = 1
   Else
      ManageBoxes False
      ClearBoxes
      GetCustomer = 0
   End If
   bDataHasChanged = False
   tab1.Tab = 0
   iTimer = 0
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getcustom"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddCustomer()
   Dim bEntry As Boolean
   Dim bResponse As Byte
   Dim iNewNumber As Long
   Dim sNewCust As String
   
   sNewCust = Compress(cmbCst)
   If sNewCust = "ALL" Then
      Beep
      MsgBox "ALL Is An Illegal Customer Nickname.", vbExclamation, Caption
      Exit Sub
   End If
   bResponse = IllegalCharacters(cmbCst)
   If bResponse > 0 Then
      MsgBox "The Nickname Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   bResponse = MsgBox(cmbCst & " Wasn't Found. Add It?", ES_YESQUESTION, Caption)
   If bResponse <> vbYes Then
      bGoodCust = False
      MouseCursor 0
      CancelTrans
      cmbCst = ""
      On Error Resume Next
      cmdCan.SetFocus
      Exit Sub
   Else
      bEntry = CheckValidColumn(cmbCst)
      bGoodCust = False
      If Not bEntry Then Exit Sub
      MouseCursor 13
      On Error Resume Next
      sSql = "SELECT MAX(CUNUMBER) FROM CustTable "
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst, ES_KEYSET)
      If bSqlRows Then
         With rdoCst
            If Not IsNull(.Fields(0)) Then
               iNewNumber = .Fields(0) + 1
            Else
               iNewNumber = 1
            End If
            .Cancel
         End With
      Else
         iNewNumber = 1
      End If
      On Error GoTo DiaErr1
      sSql = "INSERT CustTable (CUREF,CUNICKNAME,CUNUMBER) " _
             & "VALUES('" & sNewCust & "','" _
             & cmbCst & "'," _
             & Trim(str(iNewNumber)) & ")"
      clsADOCon.ExecuteSql sSql
      cmbSlp = ""
      On Error Resume Next
      'lblNum = iNewNumber
      txtNumber = iNewNumber
      cmbCst.AddItem cmbCst
      SysMsg cmbCst & " Added.", True, Me
      bNewCust = 1
      bDataHasChanged = False
      bGoodCust = GetCustomer(0)
      txtNme.SetFocus
   End If
   Exit Sub
   
DiaErr1:
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   bGoodCust = False
   ManageBoxes False
   MouseCursor 0
   MsgBox CurrError.Description & vbCrLf & "Couldn't Add Customer.", vbExclamation, Caption
   
End Sub


Private Sub lblTrm_Change()
   lblTrm.ToolTipText = lblTrm
   
End Sub

Private Sub optCut_Click()
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUCUTOFF = optCut.Value
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub optCut_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optEInv_Click()
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUEINVOICING = optEInv.Value
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
End Sub


Private Sub optEom_Click()
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUBEOM = optEom.Value
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub optEom_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub tab1_Click(PreviousTab As Integer)
   On Error Resume Next
   iTimer = 0
   ClearCombos
   If bOnLoad Then Exit Sub
   
   Select Case tab1.Tab
      Case 0
         txtBeg.SetFocus
      Case 1
         txtStNme.SetFocus
      Case 2
         txtBtNme.SetFocus
      Case 3
         txtQaRep.SetFocus
      Case 4
         txtCreditLimit.SetFocus
   End Select
   
End Sub

Private Sub Timer1_Timer(Index As Integer)
   Dim bReponse As Byte
   Dim sMsg As String
   iTimer = iTimer + 1
   If iTimer > 300 Then
      '6000/300 sec 30 min
      On Error Resume Next
      rdoCst.Close
      sMsg = Format(Now, "mm/dd/yy hh:mm AM/PM") _
             & " This Normal Time Out." & vbCr _
             & "All Current Data Has Been Saved."
      MsgBox sMsg, vbInformation, Caption
      Timer1(0).enabled = False
      Unload Me
   End If
   
End Sub

Private Sub txtAcd_Change()
   If Len(txtAcd) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtAcd_LostFocus()
   txtAcd = CheckLen(txtAcd, 3)
   If Len(txtAcd) > 1 Then txtAcd = Format(Abs(Val(txtAcd)), "000")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUAREACODE = Val(txtAcd)
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtAdr_LostFocus()
   txtAdr = CheckLen(txtAdr, 160)
   If bNewCust = 1 Then
      txtStAdr = txtAdr
      txtBtAdr = txtAdr
   End If
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUADR = "" & txtAdr
      rdoCst!CUSTADR = "" & txtStAdr
      rdoCst!CUBTADR = "" & txtBtAdr
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub








Private Sub txtApEmail_LostFocus()
   txtApEmail = CheckLen(txtApEmail, 60)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUAPEMAIL = "" & txtApEmail
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
End Sub

Private Sub txtArCon_LostFocus()
   txtArCon = CheckLen(txtArCon, 20)
   txtArCon = StrCase(txtArCon)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUAPCONTACT = "" & txtArCon
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArDay_LostFocus()
   txtArDay = CheckLen(txtArDay, 3)
   If Val(txtArDay) > 31 Then txtArDay = "31"
   txtArDay = Format(Abs(Val(txtArDay)), "##0")
   iTimer = 0
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUDAYS = Val(txtArDay)
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArDis_LostFocus()
   txtArDis = CheckLen(txtArDis, 7)
   If Val(txtArDis) > 50 Then txtArDis = "50"
   txtArDis = Format(Abs(Val(txtArDis)), "##0.000")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUARDISC = Val(txtArDis)
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArDpp_LostFocus()
   txtArDpp = CheckLen(txtArDpp, 30)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUDPP = "" & txtArDpp
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArExm_LostFocus()
   txtArExm = CheckLen(txtArExm, 30)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUTAXEXEMPT = "" & txtArExm
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArExt_LostFocus()
   txtArExt = CheckLen(txtArExt, 4)
   txtArExt = Format(Abs(Val(txtArExt)), "###0")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUAPPHONEEXT = Val(txtArExt)
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArFax_LostFocus()
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUAPFAX = "" & txtArFax
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArIntf_Change()
   If Len(txtArIntf) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtArIntf_LostFocus()
   txtArIntf = CheckLen(txtArIntf, 4)
   txtArIntf = Format(Abs(Val(txtArIntf)), "###")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUAPINTFAX = txtArIntf
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtArIntp_Change()
   If Len(txtArIntp) > 3 Then SendKeys "{tab}"
   
End Sub


Private Sub txtArIntp_LostFocus()
   txtArIntp = CheckLen(txtArIntp, 4)
   txtArIntp = Format(Abs(Val(txtArIntp)), "###")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUAPINTPHONE = txtArIntp
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtArNet_LostFocus()
   txtArNet = CheckLen(txtArNet, 3)
   'If Val(txtArNet) > 31 Then txtArNet = "31"
   txtArNet = Format(Abs(Val(txtArNet)), "##0")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUNETDAYS = Val(txtArNet)
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArPct_LostFocus()
   txtArPct = CheckLen(txtArPct, 7)
   txtArPct = Format(Abs(Val(txtArPct)), "##0.000")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUARPCT = Val(txtArPct)
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArPhn_LostFocus()
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUAPPHONE = "" & txtArPhn
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub



Private Sub txtArTxc_LostFocus()
   txtArTxc = CheckLen(txtArTxc, 12)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUTAXCODE = txtArTxc
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub




Private Sub txtBeg_DropDown()
   bDataHasChanged = True
   ShowCalendar Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) > 0 Then txtBeg = CheckDate(txtBeg)
   If bGoodCust = 1 Then
      On Error Resume Next
      If Len(txtBeg) Then rdoCst!CUSINCE = txtBeg Else _
             rdoCst!CUSINCE = Null
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtBtAdr_LostFocus()
   txtBtAdr = CheckLen(txtBtAdr, 160)
   iTimer = 0
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUBTADR = "" & txtBtAdr
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtBtcntr_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub txtBtCntr_LostFocus()
   txtBtcntr = CheckLen(txtBtcntr, 15)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUBCOUNTRY = "" & txtBtcntr
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtBtCol_LostFocus()
   txtBtCol = CheckLen(txtBtCol, 1024)
   txtBtCol = StrCase(txtBtCol, ES_FIRSTWORD)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUCOL = "" & txtBtCol
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtBtCty_LostFocus()
   txtBtCty = CheckLen(txtBtCty, 18)
   txtBtCty = StrCase(txtBtCty)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUBCITY = "" & txtBtCty
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtBtNme_LostFocus()
   txtBtNme = CheckLen(txtBtNme, 40)
   txtBtNme = StrCase(txtBtNme)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUBTNAME = "" & txtBtNme
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtBtZip_LostFocus()
   txtBtZip = CheckLen(txtBtZip, 10)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUBZIP = "" & txtBtZip
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtCcInt_Change()
   If Len(txtCcInt) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtCcInt_LostFocus()
   txtCcInt = CheckLen(txtCcInt, 4)
   txtCcInt = Format(Abs(Val(txtCcInt)), "###")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUCINTPHONE = txtCcInt
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtCint_Change()
   If Len(txtCint) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtCint_LostFocus()
   txtCint = CheckLen(txtCint, 4)
   txtCint = Format(Abs(Val(txtCint)), "###")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUINTPHONE = txtCint
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtCintf_Change()
   If Len(txtCintf) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtCintf_LostFocus()
   txtCintf = CheckLen(txtCintf, 4)
   txtCintf = Format(Abs(Val(txtCintf)), "###")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUINTFAX = txtCintf
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtCnt_LostFocus()
   txtCnt = CheckLen(txtCnt, 20)
   txtCnt = StrCase(txtCnt)
   iTimer = 0
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUCCONTACT = "" & txtCnt
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtCntr_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub txtCntr_LostFocus()
   txtCntr = CheckLen(txtCntr, 15)
   If bNewCust = 1 Then
      txtStcntr = txtCntr
      txtBtcntr = txtCntr
   End If
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUCOUNTRY = "" & txtCntr
      rdoCst!CUSTCOUNTRY = "" & txtStcntr
      rdoCst!CUBCOUNTRY = "" & txtBtcntr
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   bNewCust = 0
   
End Sub


Private Sub txtCph_LostFocus()
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUCPHONE = "" & txtCph
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtCreditLimit_LostFocus()
   On Error Resume Next
   txtCreditLimit = Format(CCur(txtCreditLimit), "###,###,##0.00")
   If Err Then
      txtCreditLimit = "0.00"
   End If
   On Error GoTo 0
   
   If bGoodCust = 1 Then
      rdoCst!CUCREDITLIMIT = CCur(txtCreditLimit)
      rdoCst.Update
   End If
End Sub

Private Sub txtCty_LostFocus()
   txtCty = CheckLen(txtCty, 18)
   txtCty = StrCase(txtCty)
   If bNewCust = 1 Then
      txtStCty = txtCty
      txtBtCty = txtCty
   End If
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUCITY = "" & txtCty
      rdoCst!CUSTCITY = "" & txtStCty
      rdoCst!CUBCITY = "" & txtBtCty
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtEml_DblClick()
   If Trim(txtEml) <> "" Then SendEMail Trim(txtEml)
   
End Sub


Private Sub txtEml_LostFocus()
   txtEml = CheckLen(txtEml, 60)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUEMAIL = "" & txtEml
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtExt_LostFocus()
   txtExt = CheckLen(txtExt, 4)
   txtExt = Format(Abs(Val(txtExt)), "###0")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUCEXT = Val(txtExt)
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtFax_LostFocus()
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUFAX = "" & txtFax
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtNme_LostFocus()
   txtNme = CheckLen(txtNme, 40)
   txtNme = StrCase(txtNme)
   If bNewCust = 1 Then
      txtStNme = txtNme
      txtBtNme = txtNme
   End If
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUNAME = "" & txtNme
      rdoCst!CUBTNAME = "" & txtBtNme
      rdoCst!CUSTNAME = "" & txtStNme
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtNumber_LostFocus()
   
   'check for valid customer number
   Dim ok As Boolean
   If IsValidCustomerNumber(txtNumber.Text, cmbCst.Text) Then
      If bGoodCust = 1 Then
         On Error Resume Next
         rdoCst!CUNUMBER = CLng("0" & txtNumber)
         rdoCst.Update
         If Err > 0 Then bGoodCust = GetCustomer(0)
      End If
   Else
      txtNumber.Text = rdoCst!CUNUMBER
   End If
   
End Sub

Private Function IsValidCustomerNumber(sNumber As String, sNickName As String) As Boolean
   'don't allow a customer number assigned to another customer
   Dim rdo As ADODB.Recordset
   Dim num As Long
   num = CLng("0" & sNumber)
   
   sSql = "SELECT CUREF FROM CustTable" & vbCrLf _
          & "WHERE CUNUMBER = " + CStr(num) & vbCrLf _
          & "AND CUREF <> '" & Compress(sNickName) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If bSqlRows Then
      MsgBox "Customer number " & CStr(num) _
         & " is assigned to customer " + rdo!CUREF
      IsValidCustomerNumber = False
   Else
      IsValidCustomerNumber = True
   End If
   rdo.Close
   Set rdo = Nothing
   
End Function

Private Sub txtPhn_LostFocus()
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUTELEPHONE = "" & txtPhn
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

































Private Sub txtQaEml_DblClick()
   If Trim(txtQaEml) <> "" Then SendEMail Trim(txtQaEml)
   
End Sub


Private Sub txtQaEml_LostFocus()
   txtQaEml = CheckLen(txtQaEml, 60)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUQAEMAIL = "" & txtQaEml
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtQaExt_LostFocus()
   txtQaExt = CheckLen(txtQaExt, 4)
   txtQaExt = Format(Abs(Val(txtQaExt)), "###0")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUQAPHONEEXT = Val(txtQaExt)
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtQaFax_LostFocus()
   txtQaFax = CheckLen(txtQaFax, 12)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUQAFAX = "" & txtQaFax
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtQaIntf_Change()
   If Len(txtQaIntf) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtQaIntf_LostFocus()
   txtQaIntf = CheckLen(txtQaIntf, 4)
   txtQaIntf = Format(Abs(Val(txtQaIntf)), "###")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUQAINTFAX = txtQaIntf
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtQaIntp_Change()
   If Len(txtQaIntp) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtQaIntp_LostFocus()
   txtQaIntp = CheckLen(txtQaIntp, 4)
   txtQaIntp = Format(Abs(Val(txtQaIntp)), "###")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUQAINTPHONE = txtQaIntp
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtQaPhn_LostFocus()
   iTimer = 0
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUQAPHONE = "" & txtQaPhn
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtQaRep_LostFocus()
   txtQaRep = CheckLen(txtQaRep, 20)
   txtQaRep = StrCase(txtQaRep)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUQAREP = "" & txtQaRep
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtSln_LostFocus()
   txtSln = CheckLen(txtSln, 6)
   txtSln = Format(Abs(Val(txtSln)), "000000")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUSLSMAN = Val(txtSln)
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub





Private Sub txtStAdr_LostFocus()
   txtStAdr = CheckLen(txtStAdr, 160)
   iTimer = 0
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUSTADR = "" & txtStAdr
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtStCntr_LostFocus()
   txtStcntr = CheckLen(txtStcntr, 15)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUSTCOUNTRY = "" & txtStcntr
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtStCty_LostFocus()
   txtStCty = CheckLen(txtStCty, 18)
   txtStCty = StrCase(txtStCty)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUSTCITY = "" & txtStCty
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtStDis_LostFocus()
   txtStDis = CheckLen(txtStDis, 7)
   txtStDis = Format(Abs(Val(txtStDis)), "##0.000")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUDISCOUNT = Val(txtStDis)
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtStFob_LostFocus()
   txtStFob = CheckLen(txtStFob, 12)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUFOB = "" & txtStFob
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtStFra_LostFocus()
   txtStFra = CheckLen(txtStFra, 7)
   txtStFra = Format(Abs(Val(txtStFra)), "##0.000")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUFRTALLOW = Val(txtStFra)
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtStFrd_LostFocus()
   txtStFrd = CheckLen(txtStFrd, 4)
   txtStFrd = Format(Abs(Val(txtStFrd)), "##0")
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUFRTDAYS = Val(txtStFrd)
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtStNme_LostFocus()
   txtStNme = CheckLen(txtStNme, 40)
   txtStNme = StrCase(txtStNme)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUSTNAME = "" & txtStNme
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtStSel_LostFocus()
   txtStSel = CheckLen(txtStSel, 1024)
   txtStSel = StrCase(txtStSel, ES_FIRSTWORD)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUSEL = "" & txtStSel
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub



Private Sub txtStVia_LostFocus()
   txtStVia = CheckLen(txtStVia, 30)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUVIA = "" & txtStVia
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtStZip_LostFocus()
   txtStZip = CheckLen(txtStZip, 10)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUSTZIP = "" & txtStZip
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtType_LostFocus()
   txtType = CheckLen(txtType, 2)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUTYPE = "" & txtType
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtWeb_DblClick()
   If Trim(txtWeb) <> "" Then OpenWebPage txtWeb
   
End Sub


Private Sub txtWeb_LostFocus()
   txtWeb = CheckLen(txtWeb, 60)
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUWEB = "" & txtWeb
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtZip_LostFocus()
   txtZip = CheckLen(txtZip, 10)
   If bNewCust = 1 Then
      txtStZip = txtZip
      txtBtZip = txtZip
   End If
   If bGoodCust = 1 Then
      On Error Resume Next
      rdoCst!CUZIP = "" & txtZip
      rdoCst!CUSTZIP = "" & txtStZip
      rdoCst!CUBZIP = "" & txtBtZip
      rdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub



Private Sub ClearBoxes()
   Dim iList As Integer
   For iList = 0 To Controls.count - 1
      If TypeOf Controls(iList) Is TextBox Then Controls(iList).Text = ""
      If TypeOf Controls(iList) Is MaskEdBox Then
         Controls(iList).Mask = ""
         Controls(iList).Text = ""
      End If
   Next
   
   'optCut.Value = vbUnchecked
   txtStFrd = "0"
   txtStFra = "0.000"
   txtStDis = "0.000"
   txtArDis = "0.000"
   txtArDay = "0"
   txtArNet = "0"
   txtArPct = "0.000"
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   b = AutoFormatChange(Me, txtChange())
   
End Sub

Private Sub GetSalesPerson()
   Dim rdoSlp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT SPNUMBER,SPFIRST,SPLAST,SPREGION FROM SprsTable " _
          & "WHERE SPNUMBER='" & cmbSlp & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoSlp)
   If bSqlRows Then
      cmbSlp = "" & Trim(rdoSlp!SPNumber)
      lblSlp = "" & Trim(rdoSlp!SPFIRST) & " " & Trim(rdoSlp!SPLAST)
      If Trim(cmbReg) = "" Then cmbReg = "" & Trim(rdoSlp!SPREGION)
   Else
      lblSlp = ""
   End If
   On Error Resume Next
   Set rdoSlp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsalesper"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetDivision()
   Dim RdoDiv As ADODB.Recordset
   Dim sDivision As String
   sDivision = cmbDiv
   On Error Resume Next
   sSql = "SELECT DIVDESC FROM CdivTable WHERE DIVREF='" & sDivision & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDiv)
   If bSqlRows Then
      lblDiv = "" & Trim(RdoDiv!DIVDESC)
   Else
      lblDiv = ""
   End If
   On Error Resume Next
   RdoDiv.Close
   Set RdoDiv = Nothing
   
End Sub

Private Sub GetRegion()
   Dim RdoReg As ADODB.Recordset
   Dim sRegion As String
   On Error Resume Next
   sRegion = cmbReg
   sSql = "SELECT REGDESC FROM CregTable WHERE REGREF='" & sRegion & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoReg)
   If bSqlRows Then
      lblReg = "" & Trim(RdoReg!REGDESC)
   Else
      lblReg = ""
   End If
   RdoReg.Close
   Set RdoReg = Nothing
   
End Sub

Private Sub GetTerms()
   Dim RdoTrm As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT TRMDESC FROM StrmTable WHERE TRMREF='" _
          & Compress(cmbTrm) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTrm)
   If bSqlRows Then
      lblTrm = "" & Trim(RdoTrm!TRMDESC)
   Else
      lblTrm = ""
   End If
   RdoTrm.Close
   Set RdoTrm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getterms"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillCountries()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_FillCountries"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            AddComboStr txtCntr.hWnd, "" & Trim(!COUCOUNTRY)
            AddComboStr txtStcntr.hWnd, "" & Trim(!COUCOUNTRY)
            AddComboStr txtBtcntr.hWnd, "" & Trim(!COUCOUNTRY)
            .MoveNext
         Loop
         .Cancel
      End With
   Else
      sSql = "SELECT DISTINCT CUCOUNTRY FROM CustTable " _
             & "ORDER BY CUCOUNTRY"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
      If bSqlRows Then
         With RdoCmb
            Do Until .EOF
               AddComboStr txtCntr.hWnd, "" & Trim(!CUCOUNTRY)
               AddComboStr txtStcntr.hWnd, "" & Trim(!CUCOUNTRY)
               AddComboStr txtBtcntr.hWnd, "" & Trim(!CUCOUNTRY)
               .MoveNext
            Loop
            .Cancel
         End With
      End If
   End If
   If txtCntr.ListCount > 0 Then
      txtCntr = "USA"
      txtStcntr = "USA"
      txtBtcntr = "USA"
   Else
      txtCntr.AddItem "USA"
      txtCntr = "USA"
      txtStcntr.AddItem "USA"
      txtStcntr = "USA"
      txtBtcntr.AddItem "USA"
      txtBtcntr = "USA"
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcountries"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FillTaxCodes()
   Dim RdoTax As ADODB.Recordset
   sSql = "SELECT TAXREF FROM TxcdTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTax, ES_FORWARD)
   If bSqlRows Then
      With RdoTax
         Do Until .EOF
            AddComboStr txtArTxc.hWnd, "" & Trim(!TAXREF)
            AddComboStr cmbTxr.hWnd, "" & Trim(!TAXREF)
            AddComboStr cmbTxw.hWnd, "" & Trim(!TAXREF)
            .MoveNext
         Loop
      End With
   End If
   Set RdoTax = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "filltaxcod"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillScAccount()
   Dim RdoRns As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT GLACCTREF,GLACCTNO FROM GlacTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns, ES_FORWARD)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            cmbArSvc.AddItem "" & Trim(!GLACCTNO)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   On Error Resume Next
   Set RdoRns = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillscacct"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub UpdateCustomerSet(Optional ShowMouse As Boolean)
   If ShowMouse Then MouseCursor 13
   On Error GoTo DiaErr1
   
   'don't allow customer number change to a customer number already in use
   
   
   With rdoCst
      !CUSVCACCT = "" & cmbArSvc
      !CUBSTATE = "" & cmbBtSte
      !CUDIVISION = "" & cmbDiv
      !CUREGION = "" & cmbReg
      !CUSALESMAN = "" & cmbSlp
      !CUSTATE = "" & cmbSte
      !CUSTSTATE = "" & cmbStSte
      !CUSTERMS = "" & cmbTrm
      !CUBORTAXCODE = "" & cmbTxr
      !CUBOWTAXCODE = "" & cmbTxw
      !CUCUTOFF = optCut.Value
      !CUBEOM = optEom.Value
      !CUAREACODE = Val(txtAcd)
      !CUBSTATE = "" & cmbBtSte
      !CUDIVISION = "" & cmbDiv
      !CUREGION = "" & cmbReg
      !CUSALESMAN = "" & cmbSlp
      !CUSTATE = "" & cmbSte
      !CUSTSTATE = "" & cmbStSte
      !CUSTERMS = "" & cmbTrm
      !CUBORTAXCODE = "" & cmbTxr
      !CUBOWTAXCODE = "" & cmbTxw
      !CUCUTOFF = optCut.Value
      !CUBEOM = optEom.Value
      !CUAREACODE = Val(txtAcd)
      !CUADR = "" & txtAdr
      !CUAPCONTACT = "" & txtArCon
      !CUDAYS = Val(txtArDay)
      !CUARDISC = Val(txtArDis)
      !CUDPP = "" & txtArDpp
      !CUTAXEXEMPT = "" & txtArExm
      !CUAPPHONEEXT = Val(txtArExt)
      !CUAPINTFAX = txtArIntf
      !CUAPINTPHONE = txtArIntp
      !CUNETDAYS = Val(txtArNet)
      !CUARPCT = Val(txtArPct)
      !CUAPPHONE = "" & txtArPhn
      !CUTAXCODE = txtArTxc
      !CUBTADR = "" & txtBtAdr
      !CUBCOUNTRY = "" & txtBtcntr
      !CUCOL = "" & txtBtCol
      !CUBCITY = "" & txtBtCty
      !CUBTNAME = "" & txtBtNme
      !CUBZIP = "" & txtBtZip
      !CUCINTPHONE = txtCcInt
      !CUINTPHONE = txtCint
      !CUINTFAX = txtCintf
      !CUCCONTACT = "" & txtCnt
      !CUCOUNTRY = "" & txtCntr
      !CUCPHONE = "" & txtCph
      !CUCITY = "" & txtCty
      !CUEMAIL = "" & txtEml
      !CUCEXT = Val(txtExt)
      !CUFAX = "" & txtFax
      !CUNAME = "" & txtNme
      !CUTELEPHONE = "" & txtPhn
      !CUQAEMAIL = "" & txtQaEml
      !CUQAPHONEEXT = Val(txtQaExt)
      !CUQAFAX = "" & txtQaFax
      !CUQAINTFAX = txtQaIntf
      !CUQAINTPHONE = txtQaIntp
      !CUQAPHONE = "" & txtQaPhn
      !CUQAREP = "" & txtQaRep
      !CUSLSMAN = Val(txtSln)
      !CUSTADR = "" & txtStAdr
      !CUSTCOUNTRY = "" & txtStcntr
      !CUSTCITY = "" & txtStCty
      !CUDISCOUNT = Val(txtStDis)
      !CUFOB = "" & txtStFob
      !CUFRTALLOW = Val(txtStFra)
      !CUFRTDAYS = Val(txtStFrd)
      !CUSTNAME = "" & txtStNme
      !CUSEL = "" & txtStSel
      !CUVIA = "" & txtStVia
      !CUSTZIP = "" & txtStZip
      !CUZIP = "" & txtZip
      .Update
   End With
   bDataHasChanged = False
   MouseCursor 0
   Exit Sub
DiaErr1:
   On Error GoTo 0
   
End Sub
