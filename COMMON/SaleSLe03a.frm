VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form SaleSLe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customers"
   ClientHeight    =   7245
   ClientLeft      =   1950
   ClientTop       =   750
   ClientWidth     =   7695
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   2103
   Icon            =   "SaleSLe03a.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7245
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame tabFrame 
      Enabled         =   0   'False
      Height          =   5175
      Index           =   4
      Left            =   7980
      TabIndex        =   137
      Top             =   1680
      Width           =   7452
      Begin VB.TextBox txtCreditLimit 
         Height          =   285
         Left            =   1440
         TabIndex        =   53
         Tag             =   "1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtApEmail 
         Height          =   285
         Left            =   1440
         TabIndex        =   72
         Tag             =   "2"
         ToolTipText     =   "Double Click Here To Send E-Mail (Requires An Entry)"
         Top             =   4080
         Width           =   5220
      End
      Begin VB.CheckBox optEInv 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   252
         Left            =   1440
         TabIndex        =   71
         ToolTipText     =   "Electronic Invoicing"
         Top             =   3720
         Width           =   735
      End
      Begin VB.ComboBox cmbArSvc 
         Height          =   288
         Left            =   1440
         TabIndex        =   57
         Tag             =   "3"
         Top             =   852
         Width           =   1935
      End
      Begin VB.ComboBox txtArTxc 
         Height          =   288
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   62
         Tag             =   "3"
         ToolTipText     =   "Select Or Enter Customer (10 Char Max)"
         Top             =   1932
         Width           =   1680
      End
      Begin VB.TextBox txtArIntf 
         Height          =   285
         Left            =   4920
         TabIndex        =   69
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   3372
         Width           =   475
      End
      Begin VB.TextBox txtArIntp 
         Height          =   285
         Left            =   1440
         TabIndex        =   66
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   3372
         Width           =   475
      End
      Begin VB.ComboBox cmbTxw 
         Height          =   288
         Left            =   4920
         Sorted          =   -1  'True
         TabIndex        =   64
         Tag             =   "3"
         ToolTipText     =   "Select Or Enter Customer (10 Char Max)"
         Top             =   2532
         Width           =   1680
      End
      Begin VB.ComboBox cmbTxr 
         Height          =   288
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   63
         Tag             =   "3"
         ToolTipText     =   "Select Or Enter Customer (10 Char Max)"
         Top             =   2532
         Width           =   1680
      End
      Begin VB.TextBox txtArExt 
         Height          =   285
         Left            =   3840
         TabIndex        =   68
         Tag             =   "1"
         Top             =   3372
         Width           =   495
      End
      Begin VB.TextBox txtArCon 
         Height          =   285
         Left            =   1440
         TabIndex        =   65
         Tag             =   "2"
         Top             =   2940
         Width           =   2055
      End
      Begin VB.CheckBox optCut 
         Alignment       =   1  'Right Justify
         Caption         =   "Cut Off?"
         Height          =   255
         Left            =   5160
         TabIndex        =   59
         ToolTipText     =   "Customer's Credit Has Been Halted"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtArDpp 
         Height          =   285
         Left            =   1440
         TabIndex        =   61
         Tag             =   "3"
         Top             =   1572
         Width           =   3315
      End
      Begin VB.TextBox txtArExm 
         Height          =   285
         Left            =   1440
         TabIndex        =   60
         Tag             =   "3"
         Top             =   1212
         Width           =   3315
      End
      Begin VB.TextBox txtArPct 
         Height          =   285
         Left            =   3960
         TabIndex        =   58
         Tag             =   "1"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtArNet 
         Height          =   285
         Left            =   6000
         TabIndex        =   56
         Tag             =   "1"
         Top             =   495
         Width           =   615
      End
      Begin VB.TextBox txtArDay 
         Height          =   285
         Left            =   4800
         TabIndex        =   55
         Tag             =   "1"
         Top             =   495
         Width           =   615
      End
      Begin VB.TextBox txtArDis 
         Height          =   285
         Left            =   3720
         TabIndex        =   54
         Tag             =   "1"
         Top             =   480
         Width           =   735
      End
      Begin MSMask.MaskEdBox txtArPhn 
         Height          =   288
         Left            =   1920
         TabIndex        =   67
         Top             =   3372
         Width           =   1308
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   14
         Mask            =   """###-###-####"""
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtArFax 
         Height          =   288
         Left            =   5400
         TabIndex        =   70
         Top             =   3372
         Width           =   1308
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   14
         Mask            =   """###-###-####"""
         PromptChar      =   "_"
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit"
         Height          =   285
         Index           =   45
         Left            =   120
         TabIndex        =   163
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Edit Receivables Fields in the Finance Module"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   162
         Top             =   180
         Width           =   5715
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "AP EMail"
         Height          =   288
         Index           =   72
         Left            =   120
         TabIndex        =   161
         Top             =   4080
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "E Invoicing"
         Height          =   288
         Index           =   71
         Left            =   120
         TabIndex        =   160
         ToolTipText     =   "Electronic Invoicing"
         Top             =   3720
         Width           =   1188
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Days"
         Height          =   285
         Index           =   60
         Left            =   6720
         TabIndex        =   154
         Top             =   540
         Width           =   585
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   288
         Index           =   59
         Left            =   4440
         TabIndex        =   153
         Top             =   3372
         Width           =   468
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext"
         Height          =   288
         Index           =   58
         Left            =   3480
         TabIndex        =   152
         Top             =   3372
         Width           =   468
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   288
         Index           =   57
         Left            =   480
         TabIndex        =   151
         Top             =   3372
         Width           =   828
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         Height          =   288
         Index           =   56
         Left            =   120
         TabIndex        =   150
         Top             =   3000
         Width           =   1188
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Wholesale"
         Height          =   288
         Index           =   55
         Left            =   3480
         TabIndex        =   149
         Top             =   2532
         Width           =   948
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Retail"
         Height          =   288
         Index           =   54
         Left            =   480
         TabIndex        =   148
         Top             =   2532
         Width           =   948
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "B and O Tax Codes:"
         Height          =   288
         Index           =   53
         Left            =   120
         TabIndex        =   147
         Top             =   2292
         Width           =   1908
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Code"
         Height          =   288
         Index           =   52
         Left            =   120
         TabIndex        =   146
         Top             =   1932
         Width           =   948
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Direct Pay Per #"
         Height          =   288
         Index           =   50
         Left            =   120
         TabIndex        =   145
         Top             =   1572
         Width           =   1308
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Exempt #"
         Height          =   288
         Index           =   49
         Left            =   120
         TabIndex        =   144
         Top             =   1212
         Width           =   1188
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   288
         Index           =   48
         Left            =   4800
         TabIndex        =   143
         Top             =   840
         Width           =   312
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   288
         Index           =   47
         Left            =   3480
         TabIndex        =   142
         Top             =   852
         Width           =   588
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Svc Chgs: Acct"
         Height          =   285
         Index           =   46
         Left            =   120
         TabIndex        =   141
         Top             =   855
         Width           =   1305
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Net"
         Height          =   285
         Index           =   44
         Left            =   5520
         TabIndex        =   140
         Top             =   540
         Width           =   435
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "% "
         Height          =   285
         Index           =   40
         Left            =   4560
         TabIndex        =   139
         Top             =   540
         Width           =   315
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Terms"
         Height          =   285
         Index           =   38
         Left            =   3060
         TabIndex        =   138
         Top             =   540
         Width           =   645
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   5175
      Index           =   3
      Left            =   8040
      TabIndex        =   131
      Top             =   2280
      Width           =   7452
      Begin VB.TextBox txtQaIntf 
         Height          =   285
         Left            =   4920
         TabIndex        =   50
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   852
         Width           =   475
      End
      Begin VB.TextBox txtQaIntp 
         Height          =   285
         Left            =   1440
         TabIndex        =   47
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   840
         Width           =   475
      End
      Begin VB.TextBox txtQaExt 
         Height          =   285
         Left            =   3720
         TabIndex        =   49
         Tag             =   "1"
         Top             =   852
         Width           =   615
      End
      Begin VB.TextBox txtQaEml 
         Height          =   285
         Left            =   1440
         TabIndex        =   52
         Tag             =   "2"
         ToolTipText     =   "Click Here To Send E-Mail (Requires An Entry)"
         Top             =   1200
         Width           =   5340
      End
      Begin VB.TextBox txtQaRep 
         Height          =   285
         Left            =   1440
         TabIndex        =   46
         Tag             =   "2"
         Top             =   480
         Width           =   2085
      End
      Begin MSMask.MaskEdBox txtQaPhn 
         Height          =   288
         Left            =   1920
         TabIndex        =   48
         Top             =   840
         Width           =   1308
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   14
         Mask            =   """###-###-####"""
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtQaFax 
         Height          =   288
         Left            =   5400
         TabIndex        =   51
         Top             =   852
         Width           =   1332
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   14
         Mask            =   """###-###-####"""
         PromptChar      =   "_"
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext"
         Height          =   288
         Index           =   62
         Left            =   3360
         TabIndex        =   136
         Top             =   840
         Width           =   468
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
         Height          =   288
         Index           =   11
         Left            =   120
         TabIndex        =   135
         Top             =   1200
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   288
         Index           =   10
         Left            =   120
         TabIndex        =   134
         Top             =   840
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   288
         Index           =   9
         Left            =   4440
         TabIndex        =   133
         Top             =   852
         Width           =   588
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quality Rep"
         Height          =   288
         Index           =   8
         Left            =   120
         TabIndex        =   132
         Top             =   492
         Width           =   1188
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   5175
      Index           =   2
      Left            =   8040
      TabIndex        =   121
      Top             =   1680
      Width           =   7452
      Begin VB.ComboBox txtBtcntr 
         Height          =   288
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   43
         Tag             =   "3"
         ToolTipText     =   "Select Country From The List"
         Top             =   2148
         Width           =   1935
      End
      Begin VB.ComboBox cmbBtSte 
         Height          =   288
         Left            =   4080
         Sorted          =   -1  'True
         TabIndex        =   41
         Tag             =   "3"
         Top             =   1788
         Width           =   735
      End
      Begin VB.TextBox txtType 
         Height          =   285
         Left            =   1440
         TabIndex        =   44
         Tag             =   "3"
         Top             =   2508
         Width           =   375
      End
      Begin VB.TextBox txtBtCol 
         Height          =   1215
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Tag             =   "9"
         ToolTipText     =   "Comments (Max 1024 Chars)"
         Top             =   2868
         Width           =   3495
      End
      Begin VB.TextBox txtBtCty 
         Height          =   285
         Left            =   1440
         TabIndex        =   40
         Tag             =   "2"
         Top             =   1788
         Width           =   1935
      End
      Begin VB.TextBox txtBtZip 
         Height          =   285
         Left            =   5280
         TabIndex        =   42
         Tag             =   "3"
         Top             =   1788
         Width           =   1335
      End
      Begin VB.TextBox txtBtAdr 
         Height          =   765
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   39
         Tag             =   "9"
         Top             =   948
         Width           =   3495
      End
      Begin VB.TextBox txtBtNme 
         Height          =   285
         Left            =   1440
         TabIndex        =   38
         Tag             =   "2"
         Top             =   588
         Width           =   3475
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         Height          =   288
         Index           =   66
         Left            =   120
         TabIndex        =   130
         Top             =   2148
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Type"
         Height          =   288
         Index           =   63
         Left            =   120
         TabIndex        =   129
         Top             =   2508
         Width           =   1212
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes:"
         Height          =   288
         Index           =   61
         Left            =   120
         TabIndex        =   128
         Top             =   2868
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   288
         Index           =   35
         Left            =   120
         TabIndex        =   127
         Top             =   588
         Width           =   1152
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   288
         Index           =   34
         Left            =   120
         TabIndex        =   126
         Top             =   948
         Width           =   1152
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   288
         Index           =   33
         Left            =   120
         TabIndex        =   125
         Top             =   1788
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         Height          =   288
         Index           =   32
         Left            =   3480
         TabIndex        =   124
         Top             =   1788
         Width           =   792
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         Height          =   288
         Index           =   31
         Left            =   4920
         TabIndex        =   123
         Top             =   1788
         Width           =   552
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill To:"
         Height          =   288
         Index           =   30
         Left            =   120
         TabIndex        =   122
         Top             =   240
         Width           =   792
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   5655
      Index           =   1
      Left            =   120
      TabIndex        =   101
      Top             =   1440
      Width           =   7455
      Begin VB.TextBox txtDoor 
         Height          =   285
         Left            =   4200
         TabIndex        =   169
         Tag             =   "3"
         Top             =   1680
         Width           =   1440
      End
      Begin VB.TextBox txtBuild 
         Height          =   285
         Left            =   1440
         TabIndex        =   167
         Tag             =   "2"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox cbPaccar 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   2400
         TabIndex        =   165
         Top             =   5160
         Width           =   855
      End
      Begin VB.CheckBox cbKanBan 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   1440
         TabIndex        =   37
         ToolTipText     =   "Is this a KANBAN Customer (Controls Printing of KANBAN Labels)"
         Top             =   4800
         Width           =   855
      End
      Begin VB.CheckBox optTrans 
         Caption         =   "__"
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   5040
         TabIndex        =   28
         ToolTipText     =   "Allow Transfers (consignments) For This Customer"
         Top             =   2400
         Width           =   495
      End
      Begin VB.CheckBox optEom 
         Caption         =   "__"
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   6840
         TabIndex        =   29
         Top             =   2400
         Width           =   495
      End
      Begin VB.ComboBox txtStcntr 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   27
         Tag             =   "3"
         ToolTipText     =   "Select Country From The List"
         Top             =   2385
         Width           =   1935
      End
      Begin VB.ComboBox cmbStSte 
         Height          =   315
         Left            =   4320
         TabIndex        =   25
         Tag             =   "3"
         Top             =   2025
         Width           =   735
      End
      Begin VB.ComboBox cmbTrm 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1440
         TabIndex        =   34
         Tag             =   "8"
         ToolTipText     =   "Select Terms From List"
         Top             =   3465
         Width           =   660
      End
      Begin VB.TextBox txtStDis 
         Height          =   285
         Left            =   5040
         TabIndex        =   35
         Tag             =   "1"
         Top             =   3465
         Width           =   615
      End
      Begin VB.TextBox txtStFra 
         Height          =   285
         Left            =   5040
         TabIndex        =   33
         Tag             =   "1"
         Top             =   3105
         Width           =   615
      End
      Begin VB.TextBox txtStFrd 
         Height          =   285
         Left            =   1440
         TabIndex        =   32
         Tag             =   "1"
         Top             =   3105
         Width           =   615
      End
      Begin VB.TextBox txtStVia 
         Height          =   285
         Left            =   5040
         TabIndex        =   31
         Tag             =   "3"
         Top             =   2745
         Width           =   1935
      End
      Begin VB.TextBox txtStFob 
         Height          =   285
         Left            =   1440
         TabIndex        =   30
         Tag             =   "3"
         Top             =   2745
         Width           =   1575
      End
      Begin VB.TextBox txtStSel 
         Height          =   885
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Tag             =   "9"
         ToolTipText     =   "Comments (Max 1024 Chars)"
         Top             =   3825
         Width           =   3495
      End
      Begin VB.TextBox txtStZip 
         Height          =   285
         Left            =   5640
         TabIndex        =   26
         Tag             =   "3"
         Top             =   2025
         Width           =   1440
      End
      Begin VB.TextBox txtStCty 
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Tag             =   "2"
         Top             =   2025
         Width           =   1935
      End
      Begin VB.TextBox txtStAdr 
         Height          =   645
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   23
         Tag             =   "9"
         Top             =   948
         Width           =   3495
      End
      Begin VB.TextBox txtStNme 
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         Tag             =   "2"
         Top             =   588
         Width           =   3475
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Door"
         Height          =   285
         Index           =   74
         Left            =   3720
         TabIndex        =   170
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Building Num"
         Height          =   285
         Index           =   73
         Left            =   120
         TabIndex        =   168
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblCustomShip 
         Caption         =   "Print Custom Shipping Labels"
         Height          =   255
         Left            =   120
         TabIndex        =   166
         Top             =   5160
         Width           =   2295
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "KANBAN Cust?"
         Height          =   285
         Index           =   70
         Left            =   120
         TabIndex        =   164
         Top             =   4800
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transfers"
         Height          =   285
         Index           =   68
         Left            =   3720
         TabIndex        =   120
         ToolTipText     =   "Allow Transfers (consignments) For This Customer"
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "EOM"
         Height          =   285
         Index           =   67
         Left            =   5640
         TabIndex        =   119
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         Height          =   285
         Index           =   65
         Left            =   120
         TabIndex        =   118
         Top             =   2385
         Width           =   915
      End
      Begin VB.Label lblTrm 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2160
         TabIndex        =   117
         Top             =   3465
         Width           =   1815
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   285
         Index           =   41
         Left            =   4080
         TabIndex        =   116
         Top             =   3465
         Width           =   1095
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping Terms"
         Height          =   285
         Index           =   39
         Left            =   120
         TabIndex        =   115
         Top             =   3465
         Width           =   1275
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Freight Allow"
         Height          =   285
         Index           =   37
         Left            =   3720
         TabIndex        =   114
         Top             =   3105
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Freight Days"
         Height          =   285
         Index           =   36
         Left            =   120
         TabIndex        =   113
         Top             =   3105
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship Via"
         Height          =   285
         Index           =   29
         Left            =   3720
         TabIndex        =   112
         Top             =   2745
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "FOB"
         Height          =   285
         Index           =   28
         Left            =   120
         TabIndex        =   111
         Top             =   2760
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Notes:"
         Height          =   285
         Index           =   27
         Left            =   120
         TabIndex        =   110
         Top             =   4185
         Width           =   1155
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To:"
         Height          =   288
         Index           =   24
         Left            =   120
         TabIndex        =   109
         Top             =   240
         Width           =   1152
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         Height          =   285
         Index           =   23
         Left            =   5160
         TabIndex        =   108
         Top             =   2025
         Width           =   555
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         Height          =   285
         Index           =   21
         Left            =   3720
         TabIndex        =   107
         Top             =   2025
         Width           =   795
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   285
         Index           =   20
         Left            =   120
         TabIndex        =   106
         Top             =   2025
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   288
         Index           =   19
         Left            =   120
         TabIndex        =   105
         Top             =   948
         Width           =   1152
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   288
         Index           =   18
         Left            =   120
         TabIndex        =   104
         Top             =   588
         Width           =   1152
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   285
         Index           =   43
         Left            =   5640
         TabIndex        =   103
         Top             =   3480
         Width           =   315
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   285
         Index           =   42
         Left            =   5640
         TabIndex        =   102
         Top             =   3120
         Width           =   315
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   5175
      Index           =   0
      Left            =   8000
      TabIndex        =   81
      Top             =   1560
      Width           =   7452
      Begin VB.TextBox txtBook 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   157
         TabStop         =   0   'False
         Tag             =   "1"
         ToolTipText     =   "Customer Assigned Price Book"
         Top             =   4080
         Width           =   1656
      End
      Begin VB.TextBox BookExpires 
         Height          =   285
         Left            =   4164
         Locked          =   -1  'True
         TabIndex        =   156
         TabStop         =   0   'False
         Tag             =   "1"
         ToolTipText     =   "Price Book Expiration Date"
         Top             =   4080
         Width           =   800
      End
      Begin VB.ComboBox txtBeg 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Tag             =   "4"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCcInt 
         Height          =   285
         Left            =   4440
         TabIndex        =   14
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   2280
         Width           =   475
      End
      Begin VB.TextBox txtCintf 
         Height          =   285
         Left            =   4440
         TabIndex        =   11
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   1956
         Width           =   475
      End
      Begin VB.TextBox txtCint 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   1956
         Width           =   475
      End
      Begin VB.ComboBox txtCntr 
         Height          =   288
         Left            =   1440
         TabIndex        =   7
         Tag             =   "3"
         ToolTipText     =   "Select Country From The List"
         Top             =   1596
         Width           =   2055
      End
      Begin VB.ComboBox cmbSte 
         Height          =   288
         Left            =   4440
         TabIndex        =   5
         Tag             =   "3"
         Top             =   1284
         Width           =   735
      End
      Begin VB.ComboBox cmbSlp 
         ForeColor       =   &H00800000&
         Height          =   288
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   17
         Tag             =   "3"
         ToolTipText     =   "Select Salesperson From List"
         Top             =   2628
         Width           =   860
      End
      Begin VB.ComboBox cmbReg 
         ForeColor       =   &H00800000&
         Height          =   288
         Left            =   4680
         Sorted          =   -1  'True
         TabIndex        =   19
         Tag             =   "8"
         ToolTipText     =   "Select Region From List"
         Top             =   3000
         Width           =   780
      End
      Begin VB.ComboBox cmbDiv 
         ForeColor       =   &H00800000&
         Height          =   288
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   18
         Tag             =   "8"
         ToolTipText     =   "Select Division From List"
         Top             =   3000
         Width           =   860
      End
      Begin VB.TextBox txtZip 
         Height          =   285
         Left            =   6000
         TabIndex        =   6
         Tag             =   "3"
         Top             =   1284
         Width           =   1335
      End
      Begin VB.TextBox txtExt 
         Height          =   285
         Left            =   6720
         TabIndex        =   16
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtAcd 
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Tag             =   "1"
         Top             =   1956
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtWeb 
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Tag             =   "2"
         ToolTipText     =   "You May Double Click To Open A Page That Has Been Entered"
         Top             =   3720
         Width           =   5220
      End
      Begin VB.TextBox txtCnt 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   2280
         Width           =   2205
      End
      Begin VB.TextBox txtEml 
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Tag             =   "2"
         ToolTipText     =   "Double Click Here To Send E-Mail (Requires An Entry)"
         Top             =   3360
         Width           =   5220
      End
      Begin VB.TextBox txtCty 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Tag             =   "2"
         Top             =   1260
         Width           =   2085
      End
      Begin VB.TextBox txtAdr 
         Height          =   615
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   3
         Tag             =   "9"
         Top             =   600
         Width           =   3475
      End
      Begin MSMask.MaskEdBox txtCph 
         Height          =   288
         Left            =   4920
         TabIndex        =   15
         Top             =   2280
         Width           =   1308
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPhn 
         Height          =   288
         Left            =   1932
         TabIndex        =   9
         Top             =   1956
         Width           =   1308
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFax 
         Height          =   288
         Left            =   4920
         TabIndex        =   12
         Top             =   1956
         Width           =   1308
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Price Book"
         Height          =   252
         Index           =   69
         Left            =   120
         TabIndex        =   159
         Top             =   4080
         Width           =   1452
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Expires"
         Height          =   252
         Index           =   17
         Left            =   3140
         TabIndex        =   158
         Top             =   4080
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "More >>>>"
         Height          =   252
         Left            =   6240
         TabIndex        =   155
         Top             =   480
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Since"
         Height          =   288
         Index           =   51
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   1428
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         Height          =   288
         Index           =   64
         Left            =   120
         TabIndex        =   99
         Top             =   1560
         Width           =   912
      End
      Begin VB.Label lblReg 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5520
         TabIndex        =   98
         Top             =   3000
         Width           =   1452
      End
      Begin VB.Label lblDiv 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   2472
         TabIndex        =   97
         Top             =   3024
         Width           =   1332
      End
      Begin VB.Label lblSlp 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   2472
         TabIndex        =   96
         Top             =   2628
         Width           =   3060
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         Height          =   288
         Index           =   26
         Left            =   3960
         TabIndex        =   95
         Top             =   3000
         Width           =   708
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         Height          =   288
         Index           =   25
         Left            =   120
         TabIndex        =   94
         Top             =   3000
         Width           =   1068
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         Height          =   288
         Index           =   22
         Left            =   5280
         TabIndex        =   93
         Top             =   1284
         Width           =   552
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Salesperson"
         Height          =   288
         Index           =   16
         Left            =   120
         TabIndex        =   92
         Top             =   2640
         Width           =   1068
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   288
         Index           =   15
         Left            =   3840
         TabIndex        =   91
         Top             =   2280
         Width           =   708
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext"
         Height          =   288
         Index           =   14
         Left            =   6360
         TabIndex        =   90
         Top             =   2280
         Width           =   468
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Web Site"
         Height          =   288
         Index           =   13
         Left            =   120
         TabIndex        =   89
         Top             =   3720
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         Height          =   288
         Index           =   12
         Left            =   120
         TabIndex        =   88
         Top             =   2280
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
         Height          =   288
         Index           =   5
         Left            =   120
         TabIndex        =   87
         Top             =   3360
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         Height          =   288
         Index           =   4
         Left            =   3840
         TabIndex        =   86
         Top             =   1320
         Width           =   672
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   288
         Index           =   3
         Left            =   120
         TabIndex        =   85
         Top             =   1260
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   288
         Index           =   7
         Left            =   3840
         TabIndex        =   84
         Top             =   1956
         Width           =   708
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   288
         Index           =   6
         Left            =   120
         TabIndex        =   83
         Top             =   1956
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   168
         Index           =   2
         Left            =   120
         TabIndex        =   82
         Top             =   624
         Width           =   1032
      End
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   80
      Top             =   960
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   11033
      TabWidthStyle   =   2
      TabFixedWidth   =   2469
      TabFixedHeight  =   473
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Shipping/Selling"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Billing"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Quality Assurance"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Receivables"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame z2 
      Height          =   40
      Left            =   0
      TabIndex        =   79
      Top             =   1080
      Width           =   7452
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLe03a.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   78
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
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
      TabIndex        =   77
      Top             =   6240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtNme 
      Height          =   285
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "2"
      Top             =   600
      Width           =   3475
   End
   Begin VB.ComboBox cmbCst 
      Height          =   288
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Customer (10 Char Max)"
      Top             =   240
      Width           =   1560
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6600
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   6120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7245
      FormDesignWidth =   7695
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3120
      TabIndex        =   76
      Top             =   240
      Width           =   600
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   75
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   74
      Top             =   280
      Width           =   1425
   End
End
Attribute VB_Name = "SaleSLe03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Accounts (SO/B&O) 2/13/03
'10/8/03 International Calling Codes
'11/5/03 Customer Since
'2/4/04 Email
'6/2/04 Reformated entire form
'8/23/05 Added Allow Transfers
'8/9/06 Replaced Tab with TabStrip
'10/19/06 E Invoicing
'1/8/07 AP Email 7.2.6
Option Explicit
Dim RdoCst As ADODB.Recordset
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
   For iControl = 0 To Controls.Count - 1
      If TypeOf Controls(iControl) Is ComboBox Then _
                         Controls(iControl).SelLength = 0
   Next
   
End Sub

Private Sub FillSales()
   On Error GoTo DiaErr1
   sSql = "Qry_FillSalesPersons"
   LoadComboBox cmbSlp, -1
   Exit Sub
   
DiaErr1:
   sProcName = "fillsales"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub cbKanBan_LostFocus()
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUPRINTKANBAN = cbKanBan.Value
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
End Sub



Private Sub cbPaccar_LostFocus()
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUPRINTPACCAR = cbPaccar.Value
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
End Sub

Private Sub cmbArSvc_LostFocus()
   cmbArSvc = CheckLen(cmbArSvc, 12)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUSVCACCT = "" & cmbArSvc
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUBSTATE = "" & cmbBtSte
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub cmbCst_Click()
   bGoodCust = GetCustomer(0)
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If bCancel Then Exit Sub
   If Len(cmbCst) = 0 Then
      cmdCan.SetFocus
      Exit Sub
   End If
   bGoodCust = GetCustomer(1)
   If bGoodCust = 0 Then AddCustomer
   
End Sub

Private Sub cmbDiv_Change()
   If Len(cmbDiv) > 4 Then cmbDiv = Left(cmbDiv, 4)
   
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
         'RdoCst'.Edit
         RdoCst!CUDIVISION = "" & cmbDiv
         RdoCst.Update
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
    ' 2/25/2009 - find if the region exists.
    If CheckExistRegion() = True Then
        GetRegion
    Else
        MsgBox "Region not found in system setting", vbInformation
        cmbReg = ""
    End If
    
'   GetRegion
'    If Len(cmbReg) Then
'        If bGoodCust = 1 Then
'            On Error Resume Next
'            'RdoCst'.Edit
'            RdoCst!CUREGION = "" & cmbReg
'            RdoCst.Update
'            If Err > 0 Then bGoodCust = GetCustomer(0)
'        End If
'    End If
   
End Sub

Private Sub cmbSlp_Change()
   If Len(cmbSlp) > 4 Then cmbSlp = Left(cmbSlp, 4)
   
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
         'RdoCst'.Edit
         RdoCst!CUSALESMAN = "" & cmbSlp
         RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUSTATE = "" & cmbSte
      RdoCst!CUSTSTATE = "" & cmbStSte
      RdoCst!CUBSTATE = "" & cmbBtSte
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUSTSTATE = "" & cmbStSte
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUSTERMS = "" & cmbTrm
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub





Private Sub cmbTxr_LostFocus()
   cmbTxr = CheckLen(cmbTxr, 8)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUBORTAXCODE = "" & cmbTxr
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub cmbTxw_LostFocus()
   cmbTxw = CheckLen(cmbTxw, 8)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUBOWTAXCODE = "" & cmbTxw
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub cmdCan_Click()
   Timer1(0).Enabled = False
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = True
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2103
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
      FillCustomers
      If cUR.CurrentCustomer <> "" Then cmbCst = cUR.CurrentCustomer
      FillDivisions
      FillRegions
      FillTerms
      bGoodCust = GetCustomer(0)
      HideNonApplicableScreenPrompts
   End If
   'If optVewAcct.Value = vbChecked Then Unload VewAcct
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim b As Byte
   Dim cNewSize As Currency
   
   MouseCursor 13
   FormLoad Me
   For b = 0 To 4
      With tabFrame(b)
         .Visible = False
         .BorderStyle = 0
         .Left = 40
      End With
   Next
   tabFrame(0).Visible = True
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
   'If optVewAcct.Value = vbChecked Then Unload VewAcct
   Set RdoCst = Nothing
   FormUnload
   Set SaleSLe03a = Nothing
   
End Sub




Private Function GetCustomer(bSelect As Byte) As Byte
   Dim sCustRef As String
   Dim sPhone As String
   If bDataHasChanged And bGoodCust = 1 Then UpdateCustomerSet
   MouseCursor 13
   sCustRef = Compress(cmbCst)
   On Error GoTo DiaErr1
   GetCustomer = 0
   bGoodCust = 0
   ClearBoxes
   sSql = "SELECT * FROM CustTable WHERE CUREF='" & sCustRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_KEYSET)
   If bSqlRows Then
      With RdoCst
         cmbCst = "" & Trim(!CUNICKNAME)
         lblNum = Format(!CUNUMBER, "#####0")
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
         If Not IsNull(!CUSINCE) Then txtBeg = Format(!CUSINCE, "mm/dd/yyyy")
         If Len(sPhone) Then txtCph = "" & Trim(!CUCPHONE)
         If !CUCEXT > 0 Then
            txtExt = Format(!CUCEXT, "###0")
         Else
            txtExt = ""
         End If
         cmbSlp = "" & Trim(!CUSALESMAN)
         cmbDiv = "" & Trim(!CUDIVISION)
         cmbReg = "" & Trim(!CUREGION)
         txtEml = "" & Trim(!CUEMAIL)
         txtWeb = "" & Trim(!CUWEB)
         '5/24/04
         optEom.Value = !CUBEOM
         'Shipping
         txtStNme = "" & Trim(!CUSTNAME)
         txtStAdr = "" & Trim(!CUSTADR)
         txtBuild = "" & Trim(!CUSTBLDG)
         txtDoor = "" & Trim(!CUSTDOOR)
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
         txtBook = "" & Trim(!CUPRICEBOOK)
         '10/8/03 International Phones
         txtCint = "" & Trim(!CUINTPHONE)
         txtCintf = "" & Trim(!CUINTFAX)
         txtCcInt = "" & Trim(!CUCINTPHONE)
         txtQaIntp = "" & Trim(!CUQAINTPHONE)
         txtQaIntf = "" & Trim(!CUQAINTFAX)
         txtArIntp = "" & Trim(!CUAPINTPHONE)
         txtArIntf = "" & Trim(!CUAPINTFAX)
         '8/23/05
         optTrans.Value = !CUALLOWTRANSFERS
         '10/19/06
         optEInv.Value = !CUEINVOICING
         txtApEmail = "" & Trim(!CUAPEMAIL)
      
         txtCreditLimit = Format(!CUCREDITLIMIT, "###,###,##0.00")
         cbKanBan.Value = IIf(IsNull(!CUPRINTKANBAN), 0, !CUPRINTKANBAN)
         cbPaccar.Value = IIf(IsNull(!CUPRINTPACCAR), 0, !CUPRINTPACCAR)
      End With
      txtPhn.Mask = "###-###-####"
      txtFax.Mask = "###-###-####"
      txtCph.Mask = "###-###-####"
      txtQaPhn.Mask = "###-###-####"
      txtQaFax.Mask = "###-###-####"
      txtArPhn.Mask = "###-###-####"
      txtArFax.Mask = "###-###-####"
      GetDivision
      GetRegion
      GetSalesPerson
      GetTerms
      GetPriceBook Me
      Err.Clear
      GetCustomer = 1
   Else
      ClearBoxes
      GetCustomer = 0
   End If
   bDataHasChanged = False
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
   Dim lNewNumber As Long
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
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_KEYSET)
      If bSqlRows Then
         With RdoCst
            If Not IsNull(.Fields(0)) Then
               lNewNumber = .Fields(0) + 1
            Else
               lNewNumber = 1
            End If
            ClearResultSet RdoCst
         End With
      Else
         lNewNumber = 1
      End If
      On Error GoTo DiaErr1
      sSql = "INSERT CustTable (CUREF,CUNICKNAME,CUNUMBER) " _
             & "VALUES('" & sNewCust & "','" _
             & cmbCst & "'," _
             & Trim(str(lNewNumber)) & ")"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      cmbSlp = ""
      On Error Resume Next
      lblNum = lNewNumber
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
   MouseCursor 0
   MsgBox CurrError.Description & vbCrLf & "Couldn't Add Customer.", vbExclamation, Caption
   
End Sub


Private Sub lblTrm_Change()
   lblTrm.ToolTipText = lblTrm
   
End Sub

Private Sub optCut_Click()
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUCUTOFF = optCut.Value
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub optCut_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optEInv_Click()
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUEINVOICING = optEInv.Value
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub optEom_Click()
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUBEOM = optEom.Value
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub optEom_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optTrans_Click()
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUALLOWTRANSFERS = optTrans.Value
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub tab1_Click()
   Dim b As Byte
   On Error Resume Next
   iTimer = 0
   ClearCombos
   For b = 0 To 4
      tabFrame(b).Visible = False
   Next
   tabFrame(tab1.SelectedItem.Index - 1).Visible = True
   Select Case tab1.SelectedItem.Index
      Case 1
         txtBeg.SetFocus
      Case 2
         txtStNme.SetFocus
      Case 3
         txtBtNme.SetFocus
      Case 4
         txtQaRep.SetFocus
      Case 5
         txtArDis.SetFocus
   End Select
   
End Sub





Private Sub Timer1_Timer(Index As Integer)
   Dim bReponse As Byte
   Dim sMsg As String
   iTimer = iTimer + 1
   If iTimer > 300 Then
      '6000/300 sec 30 min
      On Error Resume Next
      RdoCst.Close
      sMsg = Format(Now, "mm/dd/yy hh:mm AM/PM") _
             & " This Normal Time Out." & vbCrLf _
             & "All Current Data Has Been Saved."
      MsgBox sMsg, vbInformation, Caption
      Timer1(0).Enabled = False
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
      'RdoCst'.Edit
      RdoCst!CUAREACODE = Val(txtAcd)
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUADR = "" & txtAdr
      RdoCst!CUSTADR = "" & txtStAdr
      RdoCst!CUBTADR = "" & txtBtAdr
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub








Private Sub txtApEmail_LostFocus()
   txtApEmail = CheckLen(txtApEmail, 60)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUAPEMAIL = "" & txtApEmail
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtArCon_LostFocus()
   txtArCon = CheckLen(txtArCon, 20)
   txtArCon = StrCase(txtArCon)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUAPCONTACT = "" & txtArCon
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUDAYS = Val(txtArDay)
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArDis_LostFocus()
   txtArDis = CheckLen(txtArDis, 7)
   If Val(txtArDis) > 50 Then txtArDis = "50"
   txtArDis = Format(Abs(Val(txtArDis)), "##0.000")
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUARDISC = Val(txtArDis)
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArDpp_LostFocus()
   txtArDpp = CheckLen(txtArDpp, 30)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUDPP = "" & txtArDpp
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArExm_LostFocus()
   txtArExm = CheckLen(txtArExm, 30)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUTAXEXEMPT = "" & txtArExm
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArExt_LostFocus()
   txtArExt = CheckLen(txtArExt, 4)
   txtArExt = Format(Abs(Val(txtArExt)), "###0")
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUAPPHONEEXT = Val(txtArExt)
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArFax_LostFocus()
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUAPFAX = "" & txtArFax
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUAPINTFAX = txtArIntf
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUAPINTPHONE = txtArIntp
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtArNet_LostFocus()
   txtArNet = CheckLen(txtArNet, 3)
   If Val(txtArNet) > 31 Then txtArNet = "31"
   txtArNet = Format(Abs(Val(txtArNet)), "##0")
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUNETDAYS = Val(txtArNet)
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArPct_LostFocus()
   txtArPct = CheckLen(txtArPct, 7)
   txtArPct = Format(Abs(Val(txtArPct)), "##0.000")
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUARPCT = Val(txtArPct)
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtArPhn_LostFocus()
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUAPPHONE = "" & txtArPhn
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub



Private Sub txtArTxc_LostFocus()
   txtArTxc = CheckLen(txtArTxc, 12)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUTAXCODE = txtArTxc
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub




Private Sub txtBeg_DropDown()
   ShowCalendarEx Me, 1440 + txtBeg.Height
   bDataHasChanged = True
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) > 0 Then txtBeg = CheckDateEx(txtBeg)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      If Len(txtBeg) Then RdoCst!CUSINCE = txtBeg Else _
             RdoCst!CUSINCE = Null
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtBtAdr_LostFocus()
   txtBtAdr = CheckLen(txtBtAdr, 160)
   iTimer = 0
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUBTADR = "" & txtBtAdr
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUBCOUNTRY = "" & txtBtcntr
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtBtCol_LostFocus()
   txtBtCol = CheckLen(txtBtCol, 1024)
   txtBtCol = StrCase(txtBtCol, ES_FIRSTWORD)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUCOL = "" & txtBtCol
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtBtCty_LostFocus()
   txtBtCty = CheckLen(txtBtCty, 18)
   txtBtCty = StrCase(txtBtCty)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUBCITY = "" & txtBtCty
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtBtNme_LostFocus()
   txtBtNme = CheckLen(txtBtNme, 40)
   txtBtNme = StrCase(txtBtNme)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUBTNAME = "" & txtBtNme
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtBtZip_LostFocus()
   txtBtZip = CheckLen(txtBtZip, 10)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUBZIP = "" & txtBtZip
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUCINTPHONE = txtCcInt
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUINTPHONE = txtCint
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUINTFAX = txtCintf
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtCnt_LostFocus()
   txtCnt = CheckLen(txtCnt, 20)
   txtCnt = StrCase(txtCnt)
   iTimer = 0
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUCCONTACT = "" & txtCnt
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUCOUNTRY = "" & txtCntr
      RdoCst!CUSTCOUNTRY = "" & txtStcntr
      RdoCst!CUBCOUNTRY = "" & txtBtcntr
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   bNewCust = 0
   
End Sub


Private Sub txtCph_LostFocus()
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUCPHONE = "" & txtCph
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
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
      'RdoCst'.Edit
      RdoCst!CUCITY = "" & txtCty
      RdoCst!CUSTCITY = "" & txtStCty
      RdoCst!CUBCITY = "" & txtBtCty
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUEMAIL = "" & txtEml
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtExt_LostFocus()
   txtExt = CheckLen(txtExt, 4)
   txtExt = Format(Abs(Val(txtExt)), "###0")
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUCEXT = Val(txtExt)
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtFax_LostFocus()
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUFAX = "" & txtFax
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUNAME = "" & txtNme
      RdoCst!CUBTNAME = "" & txtBtNme
      RdoCst!CUSTNAME = "" & txtStNme
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtPhn_LostFocus()
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUTELEPHONE = "" & txtPhn
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUQAEMAIL = "" & txtQaEml
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtQaExt_LostFocus()
   txtQaExt = CheckLen(txtQaExt, 4)
   txtQaExt = Format(Abs(Val(txtQaExt)), "###0")
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUQAPHONEEXT = Val(txtQaExt)
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtQaFax_LostFocus()
   txtQaFax = CheckLen(txtQaFax, 12)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUQAFAX = "" & txtQaFax
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUQAINTFAX = txtQaIntf
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUQAINTPHONE = txtQaIntp
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtQaPhn_LostFocus()
   iTimer = 0
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUQAPHONE = "" & txtQaPhn
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtQaRep_LostFocus()
   txtQaRep = CheckLen(txtQaRep, 20)
   txtQaRep = StrCase(txtQaRep)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUQAREP = "" & txtQaRep
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub







Private Sub txtStAdr_LostFocus()
   txtStAdr = CheckLen(txtStAdr, 160)
   iTimer = 0
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUSTADR = "" & txtStAdr
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtStCntr_LostFocus()
   txtStcntr = CheckLen(txtStcntr, 15)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUSTCOUNTRY = "" & txtStcntr
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtBuild_LostFocus()
   txtBuild = CheckLen(txtBuild, 12)
   txtBuild = StrCase(txtBuild)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUSTBLDG = "" & txtBuild
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtDoor_LostFocus()
   txtDoor = CheckLen(txtDoor, 12)
   txtDoor = StrCase(txtDoor)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUSTDOOR = "" & txtStCty
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub



Private Sub txtStCty_LostFocus()
   txtStCty = CheckLen(txtStCty, 18)
   txtStCty = StrCase(txtStCty)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUSTCITY = "" & txtStCty
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtStDis_LostFocus()
   txtStDis = CheckLen(txtStDis, 7)
   txtStDis = Format(Abs(Val(txtStDis)), "##0.000")
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUDISCOUNT = Val(txtStDis)
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub


Private Sub txtStFob_LostFocus()
   txtStFob = CheckLen(txtStFob, 12)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUFOB = "" & txtStFob
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtStFra_LostFocus()
   txtStFra = CheckLen(txtStFra, 7)
   txtStFra = Format(Abs(Val(txtStFra)), "##0.000")
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUFRTALLOW = Val(txtStFra)
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtStFrd_LostFocus()
   txtStFrd = CheckLen(txtStFrd, 4)
   txtStFrd = Format(Abs(Val(txtStFrd)), "##0")
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUFRTDAYS = Val(txtStFrd)
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtStNme_LostFocus()
   txtStNme = CheckLen(txtStNme, 40)
   txtStNme = StrCase(txtStNme)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUSTNAME = "" & txtStNme
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtStSel_LostFocus()
   txtStSel = CheckLen(txtStSel, 1024)
   txtStSel = StrCase(txtStSel, ES_FIRSTWORD)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUSEL = "" & txtStSel
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub



Private Sub txtStVia_LostFocus()
   txtStVia = CheckLen(txtStVia, 30)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUVIA = "" & txtStVia
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtStZip_LostFocus()
   txtStZip = CheckLen(txtStZip, 10)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUSTZIP = "" & txtStZip
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub

Private Sub txtType_LostFocus()
   txtType = CheckLen(txtType, 2)
   If bGoodCust = 1 Then
      On Error Resume Next
      'RdoCst'.Edit
      RdoCst!CUTYPE = "" & txtType
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUWEB = "" & txtWeb
      RdoCst.Update
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
      'RdoCst'.Edit
      RdoCst!CUZIP = "" & txtZip
      RdoCst!CUSTZIP = "" & txtStZip
      RdoCst!CUBZIP = "" & txtBtZip
      RdoCst.Update
      If Err > 0 Then bGoodCust = GetCustomer(0)
   End If
   
End Sub



Private Sub ClearBoxes()
   Dim iList As Integer
   For iList = 0 To Controls.Count - 1
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
   Dim RdoRows As ADODB.Recordset
   Dim b As Byte
   Dim iRows As Integer
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   b = AutoFormatChange(Me, txtChange())
   BookExpires.BackColor = Es_TextDisabled
   txtBook.BackColor = Es_TextDisabled
   On Error Resume Next
   
   '9/12/06
   sSql = "SELECT COUNT(PBHREF) as RowsCounted FROM PbhdTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRows, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoRows!RowsCounted) Then iRows = RdoRows!RowsCounted
   End If
   If iRows > 0 Then
      BookExpires.BackColor = Es_TextDisabled
      txtBook.BackColor = Es_TextDisabled
   Else
      BookExpires.Visible = False
      txtBook.Visible = False
      z1(17).Visible = False
      z1(69).Visible = False
   End If
   Set RdoRows = Nothing
   
End Sub

Private Sub GetSalesPerson()
   Dim rdoSlp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetSalesPerson '" & cmbSlp & "'"
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
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDiv, ES_FORWARD)
   If bSqlRows Then
      lblDiv = "" & Trim(RdoDiv!DIVDESC)
      ClearResultSet RdoDiv
   Else
      lblDiv = ""
   End If
   On Error Resume Next
   Set RdoDiv = Nothing
   
End Sub

Private Sub GetRegion()
   Dim RdoReg As ADODB.Recordset
   Dim sRegion As String
   On Error Resume Next
   sRegion = cmbReg
   sSql = "SELECT REGDESC FROM CregTable WHERE REGREF='" & sRegion & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoReg, ES_FORWARD)
   If bSqlRows Then
      lblReg = "" & Trim(RdoReg!REGDESC)
      ClearResultSet RdoReg
   Else
      lblReg = ""
   End If
   Set RdoReg = Nothing
   
End Sub

Function CheckExistRegion() As Boolean
   Dim RdoReg As ADODB.Recordset
   Dim sRegion As String
   On Error Resume Next
   sRegion = cmbReg
   sSql = "SELECT * FROM CregTable WHERE REGREF='" & sRegion & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoReg, ES_FORWARD)
   If bSqlRows Then
        CheckExistRegion = True
   Else
        CheckExistRegion = False
   End If
   Set RdoReg = Nothing
   
End Function


Private Sub GetTerms()
   Dim RdoTrm As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT TRMDESC FROM StrmTable WHERE TRMREF='" _
          & Compress(cmbTrm) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTrm)
   If bSqlRows Then
      lblTrm = "" & Trim(RdoTrm!TRMDESC)
      ClearResultSet RdoTrm
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
         ClearResultSet RdoCmb
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
            ClearResultSet RdoCmb
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
   With RdoTax
      Do Until .EOF
         AddComboStr txtArTxc.hWnd, "" & Trim(!TAXREF)
         AddComboStr cmbTxr.hWnd, "" & Trim(!TAXREF)
         AddComboStr cmbTxw.hWnd, "" & Trim(!TAXREF)
         .MoveNext
      Loop
      ClearResultSet RdoTax
   End With
   Set RdoTax = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "filltaxcod"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillScAccount()
   On Error GoTo DiaErr1
   sSql = "Qry_FillAccountCombo"
   LoadComboBox cmbArSvc
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
   With RdoCst
      '.Edit
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
      !CUSTBLDG = "" & txtBuild
      !CUSTDOOR = "" & txtDoor
      
      .Update
   End With
   bDataHasChanged = False
   MouseCursor 0
   Exit Sub
DiaErr1:
   On Error GoTo 0
   
End Sub



Private Sub HideNonApplicableScreenPrompts()
    Dim rdoShow As ADODB.Recordset
    
    lblCustomShip.Visible = False
    cbPaccar.Visible = False
    cbPaccar.Value = vbUnchecked
    
    sSql = "SELECT COCUSTOMSHIPLABEL FROM ComnTable WHERE COREF = 1 "
    bSqlRows = clsADOCon.GetDataSet(sSql, rdoShow, ES_FORWARD)
    If bSqlRows Then
        If rdoShow!COCUSTOMSHIPLABEL = 1 Then
            lblCustomShip.Visible = True
            cbPaccar.Visible = True
        End If
    
    End If
    
    If bGoodCust = 1 And cbPaccar.Visible = False Then
        'RdoCst'.Edit
        RdoCst!CUPRINTPACCAR = 0
        RdoCst.Update
    End If
End Sub
