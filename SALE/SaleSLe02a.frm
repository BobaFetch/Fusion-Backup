VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form SaleSLe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise Sales Order"
   ClientHeight    =   5970
   ClientLeft      =   1845
   ClientTop       =   975
   ClientWidth     =   7635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   2102
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox OptSoEDI 
      Caption         =   "FromEDI"
      Height          =   195
      Left            =   5040
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Height          =   375
      Left            =   2880
      Picture         =   "SaleSLe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   97
      ToolTipText     =   "Find Sales Order by PO Number/Customer"
      Top             =   360
      Width           =   375
   End
   Begin VB.CheckBox OptSoXml 
      Caption         =   "FromXMLSO"
      Height          =   195
      Left            =   4080
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdStatCode 
      DisabledPicture =   "SaleSLe02a.frx":043A
      DownPicture     =   "SaleSLe02a.frx":0DAC
      Height          =   315
      Left            =   6000
      MaskColor       =   &H8000000F&
      Picture         =   "SaleSLe02a.frx":123B
      Style           =   1  'Graphical
      TabIndex        =   95
      TabStop         =   0   'False
      ToolTipText     =   "Add Internal Status Code"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.Frame tabFrame 
      Height          =   3705
      Index           =   2
      Left            =   180
      TabIndex        =   88
      Top             =   1980
      Visible         =   0   'False
      Width           =   7296
      Begin VB.CheckBox chkITAREAR 
         Caption         =   "This is an ITAR/EAR SO"
         Height          =   195
         Left            =   240
         TabIndex        =   101
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txtJob 
         Height          =   285
         Left            =   1320
         TabIndex        =   36
         Tag             =   "3"
         ToolTipText     =   "Click to insert Sales Order Number"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtJnm 
         Height          =   285
         Left            =   3960
         TabIndex        =   37
         Tag             =   "2"
         ToolTipText     =   "Job Name or Number (20 char)"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton cmdComments 
         DisabledPicture =   "SaleSLe02a.frx":16CA
         DownPicture     =   "SaleSLe02a.frx":203C
         Height          =   350
         Index           =   0
         Left            =   6720
         Picture         =   "SaleSLe02a.frx":29AE
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Standard Comments"
         Top             =   360
         Width           =   350
      End
      Begin VB.TextBox txtCmt 
         Height          =   1575
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Tag             =   "9"
         ToolTipText     =   "Comment"
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Number"
         Height          =   252
         Index           =   14
         Left            =   240
         TabIndex        =   91
         Top             =   2160
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Name"
         Height          =   252
         Index           =   19
         Left            =   2760
         TabIndex        =   90
         Top             =   2160
         Width           =   1452
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Order Remarks:"
         Height          =   252
         Index           =   29
         Left            =   120
         TabIndex        =   89
         Top             =   120
         Width           =   2052
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   3705
      Index           =   1
      Left            =   7680
      TabIndex        =   77
      Top             =   2040
      Visible         =   0   'False
      Width           =   7296
      Begin VB.ComboBox cmbShpVia 
         Height          =   315
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   100
         ToolTipText     =   "Enter/Revise Terms (2 char)"
         Top             =   480
         Width           =   2100
      End
      Begin VB.CommandButton cmdComments 
         DisabledPicture =   "SaleSLe02a.frx":2FB0
         DownPicture     =   "SaleSLe02a.frx":3922
         Height          =   350
         Index           =   1
         Left            =   5200
         Picture         =   "SaleSLe02a.frx":4294
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Standard Comments"
         Top             =   1680
         Width           =   350
      End
      Begin VB.ComboBox cmbTrm 
         Height          =   288
         Left            =   1620
         Sorted          =   -1  'True
         TabIndex        =   25
         ToolTipText     =   "Enter/Revise Terms (2 char)"
         Top             =   120
         Width           =   660
      End
      Begin VB.TextBox txtBto 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   735
         Left            =   1620
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   2655
         Width           =   3495
      End
      Begin VB.TextBox txtFrd 
         Height          =   285
         Left            =   1620
         TabIndex        =   28
         Tag             =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtFra 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1620
         TabIndex        =   26
         Tag             =   "1"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtStNme 
         Height          =   285
         Left            =   1620
         TabIndex        =   30
         Tag             =   "2"
         Top             =   1320
         Width           =   3475
      End
      Begin VB.TextBox txtStAdr 
         BackColor       =   &H00FFFFFF&
         Height          =   795
         Left            =   1620
         MultiLine       =   -1  'True
         TabIndex        =   31
         Tag             =   "9"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtVia 
         Height          =   285
         Left            =   5280
         TabIndex        =   27
         Tag             =   "3"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.TextBox txtFob 
         Height          =   285
         Left            =   3960
         TabIndex        =   29
         Tag             =   "3"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtStr 
         Height          =   285
         Left            =   6900
         TabIndex        =   78
         Tag             =   "3"
         Top             =   3240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblTrm 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   2400
         TabIndex        =   87
         Top             =   120
         Width           =   2172
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill To:"
         Height          =   255
         Index           =   30
         Left            =   120
         TabIndex        =   86
         Top             =   2655
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Freight Days"
         Height          =   252
         Index           =   28
         Left            =   120
         TabIndex        =   85
         Top             =   840
         Width           =   1452
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Freight Allowance"
         Height          =   252
         Index           =   27
         Left            =   120
         TabIndex        =   84
         Top             =   480
         Width           =   1692
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To Name"
         Height          =   288
         Index           =   26
         Left            =   120
         TabIndex        =   83
         Top             =   1320
         Width           =   1632
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To Address"
         Height          =   288
         Index           =   25
         Left            =   120
         TabIndex        =   82
         Top             =   1680
         Width           =   1632
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship Via"
         Height          =   252
         Index           =   23
         Left            =   3000
         TabIndex        =   81
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "FOB"
         Height          =   252
         Index           =   22
         Left            =   3000
         TabIndex        =   80
         Top             =   840
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping Terms"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   79
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   3825
      Index           =   0
      Left            =   7920
      TabIndex        =   56
      Top             =   1920
      Width           =   7296
      Begin VB.TextBox BookExpires 
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   94
         TabStop         =   0   'False
         Tag             =   "1"
         ToolTipText     =   "Price Book Expiration Date (based on order date)"
         Top             =   3360
         Width           =   800
      End
      Begin VB.TextBox txtExt 
         Height          =   285
         Left            =   3876
         TabIndex        =   21
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtCintFax 
         Height          =   285
         Left            =   5316
         TabIndex        =   22
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   3000
         Width           =   475
      End
      Begin VB.TextBox txtCcInt 
         Height          =   285
         Left            =   1596
         TabIndex        =   19
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   3000
         Width           =   475
      End
      Begin VB.TextBox txtBook 
         Height          =   285
         Left            =   1596
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "1"
         ToolTipText     =   "Customer Assigned Price Book (based on order date)"
         Top             =   3360
         Width           =   1656
      End
      Begin VB.ComboBox txtBok 
         Height          =   315
         Left            =   1596
         TabIndex        =   3
         Tag             =   "4"
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtCnt 
         Height          =   285
         Left            =   1596
         TabIndex        =   18
         Tag             =   "2"
         ToolTipText     =   "Job Name or Number (20 char)"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.ComboBox txtShp 
         Height          =   315
         Left            =   1596
         TabIndex        =   6
         Tag             =   "4"
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox txtDdt 
         Height          =   315
         Left            =   4716
         TabIndex        =   4
         Tag             =   "4"
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox cmbSlp 
         ForeColor       =   &H00800000&
         Height          =   288
         Left            =   1596
         Sorted          =   -1  'True
         TabIndex        =   17
         Tag             =   "8"
         ToolTipText     =   "Select Salesperson From List"
         Top             =   2268
         Width           =   975
      End
      Begin VB.ComboBox cmbReg 
         Height          =   288
         Left            =   4716
         Sorted          =   -1  'True
         TabIndex        =   14
         Tag             =   "3"
         ToolTipText     =   "Select Region From List"
         Top             =   1560
         Width           =   780
      End
      Begin VB.ComboBox cmbDiv 
         Height          =   288
         Left            =   1596
         Sorted          =   -1  'True
         TabIndex        =   13
         Tag             =   "3"
         ToolTipText     =   "Select Division From List"
         Top             =   1560
         Width           =   860
      End
      Begin VB.TextBox txtDis 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4716
         TabIndex        =   16
         Tag             =   "1"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtBus 
         Height          =   285
         Left            =   1596
         TabIndex        =   15
         Tag             =   "3"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtTax 
         Height          =   285
         Left            =   4716
         TabIndex        =   12
         Tag             =   "3"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox optTax 
         Alignment       =   1  'Right Justify
         Caption         =   "Taxable      "
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1660
      End
      Begin VB.TextBox txtNet 
         Height          =   285
         Left            =   3636
         TabIndex        =   10
         Tag             =   "1"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtDay 
         Height          =   285
         Left            =   2556
         TabIndex        =   9
         Tag             =   "1"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtDsc 
         Height          =   285
         Left            =   1596
         TabIndex        =   8
         Tag             =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtCpo 
         Height          =   285
         Left            =   4716
         TabIndex        =   7
         Tag             =   "3"
         Top             =   480
         Width           =   2085
      End
      Begin MSMask.MaskEdBox txtCph 
         Height          =   288
         Left            =   2076
         TabIndex        =   20
         Top             =   3000
         Width           =   1308
         _ExtentX        =   2328
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCfax 
         Height          =   288
         Left            =   5796
         TabIndex        =   23
         Top             =   3000
         Width           =   1308
         _ExtentX        =   2328
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin VB.Label z1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Expires"
         Height          =   252
         Index           =   10
         Left            =   3360
         TabIndex        =   93
         Top             =   3360
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "More >>>"
         Height          =   252
         Left            =   6120
         TabIndex        =   92
         Top             =   120
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext"
         Height          =   288
         Index           =   34
         Left            =   3516
         TabIndex        =   76
         Top             =   3000
         Width           =   468
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   288
         Index           =   33
         Left            =   4716
         TabIndex        =   75
         Top             =   3000
         Width           =   1188
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Phone"
         Height          =   288
         Index           =   32
         Left            =   156
         TabIndex        =   74
         Top             =   3000
         Width           =   1188
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Price Book"
         Height          =   252
         Index           =   31
         Left            =   156
         TabIndex        =   73
         Top             =   3360
         Width           =   1452
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         Height          =   252
         Index           =   24
         Left            =   156
         TabIndex        =   72
         Top             =   2640
         Width           =   1452
      End
      Begin VB.Label lblSlp 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   2748
         TabIndex        =   71
         Top             =   2268
         Width           =   3060
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   252
         Index           =   21
         Left            =   5436
         TabIndex        =   70
         Top             =   1920
         Width           =   132
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   252
         Index           =   20
         Left            =   2796
         TabIndex        =   69
         Top             =   1920
         Width           =   1452
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Business Unit"
         Height          =   252
         Index           =   13
         Left            =   156
         TabIndex        =   68
         Top             =   1920
         Width           =   1452
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         Height          =   252
         Index           =   12
         Left            =   2796
         TabIndex        =   67
         Top             =   1560
         Width           =   1452
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         Height          =   252
         Index           =   11
         Left            =   156
         TabIndex        =   66
         Top             =   1560
         Width           =   1452
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Exempt Number"
         Height          =   252
         Index           =   9
         Left            =   2796
         TabIndex        =   65
         Top             =   1200
         Width           =   2052
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Net"
         Height          =   252
         Index           =   8
         Left            =   3156
         TabIndex        =   64
         Top             =   840
         Width           =   372
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   252
         Index           =   7
         Left            =   2316
         TabIndex        =   63
         Top             =   840
         Width           =   132
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Terms:"
         Height          =   252
         Index           =   6
         Left            =   156
         TabIndex        =   62
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Order Number"
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   61
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Default  Delivery Date"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   60
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Booking Date"
         Height          =   252
         Index           =   3
         Left            =   156
         TabIndex        =   59
         Top             =   120
         Width           =   1452
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Default Ship Date"
         Height          =   252
         Index           =   2
         Left            =   156
         TabIndex        =   58
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Salesperson"
         Height          =   252
         Index           =   15
         Left            =   156
         TabIndex        =   57
         Top             =   2280
         Width           =   1452
      End
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   4275
      Left            =   0
      TabIndex        =   55
      Top             =   1560
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7541
      TabWidthStyle   =   2
      TabFixedWidth   =   1587
      TabFixedHeight  =   473
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Shipping"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Other"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLe02a.frx":4896
      Style           =   1  'Graphical
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optNotes 
      Caption         =   "Notes"
      Height          =   255
      Left            =   5400
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdNte 
      DisabledPicture =   "SaleSLe02a.frx":5044
      DownPicture     =   "SaleSLe02a.frx":54C1
      Height          =   315
      Left            =   6000
      MaskColor       =   &H8000000F&
      Picture         =   "SaleSLe02a.frx":593E
      Style           =   1  'Graphical
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Selling And Credit Info"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.CommandButton optPrn 
      DownPicture     =   "SaleSLe02a.frx":5DBB
      Height          =   315
      Left            =   6000
      MaskColor       =   &H8000000F&
      Picture         =   "SaleSLe02a.frx":634D
      Style           =   1  'Graphical
      TabIndex        =   51
      TabStop         =   0   'False
      ToolTipText     =   "Print This Sales Order"
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "SaleSLe02a.frx":68DF
      Height          =   320
      Left            =   4080
      MaskColor       =   &H8000000F&
      Picture         =   "SaleSLe02a.frx":6DB9
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Show Existing Sales Orders"
      Top             =   1200
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CheckBox optSvw 
      Caption         =   "SoView"
      Height          =   195
      Left            =   2640
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox optLst 
      Caption         =   "List"
      Height          =   255
      Left            =   1920
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbCst 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select New Customer From List"
      Top             =   840
      Width           =   1555
   End
   Begin VB.CheckBox optItm 
      Caption         =   "Items"
      Height          =   195
      Left            =   4440
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optNew 
      Caption         =   "New"
      Height          =   255
      Left            =   3600
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdItm 
      Caption         =   "&Items"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6480
      TabIndex        =   38
      ToolTipText     =   "Edit Sales Order Items"
      Top             =   480
      Width           =   915
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   315
      Left            =   6480
      TabIndex        =   39
      ToolTipText     =   "Enter New Sales Order"
      Top             =   840
      Width           =   915
   End
   Begin VB.ComboBox cmbSon 
      Height          =   288
      Left            =   1800
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Select or Enter Sales Order Number (List Contains Last 3 Years Up To 500 Enties)"
      Top             =   360
      Width           =   975
   End
   Begin VB.ComboBox cmbPre 
      Height          =   288
      Left            =   1320
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "S"
      ToolTipText     =   "Select or Enter Type A thru Z"
      Top             =   360
      Width           =   520
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6480
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Find by PO"
      Height          =   495
      Index           =   16
      Left            =   3360
      TabIndex        =   98
      Top             =   360
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtCan 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4800
      TabIndex        =   50
      Top             =   360
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblCan 
      BackStyle       =   0  'Transparent
      Caption         =   "Canceled"
      Height          =   255
      Left            =   4080
      TabIndex        =   49
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   2880
      TabIndex        =   44
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label lblCst 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   360
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   42
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   41
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "SaleSLe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'9/27/03Removed Commissions
'6/25/04 Added cmdNte
'7/21/04 Contact Phone and fax
'1/18/05 Added Address Comments
'2/24/05 Added Contact Extension
'4/12/05 Update infor on Customer change (GetBillTo)
'7/28/05 Customer PO Checking (sOldPO)
'8/9/06 Replaced Tab with TabStrip
'9/8/06 Added BookExpires and colors
'10/5/86 Added SOCANCELED to the SohdTable trap (GetSalesOrder) query
Option Explicit
'Dim rdoQry As rdoQuery
Dim cmdObj As ADODB.Command
Dim RdoRso As ADODB.Recordset

Dim bCanceled As Boolean
Dim bCutOff As Byte
Dim bGoodSO As Byte
Dim bItems As Byte
Dim bOnLoad As Byte
Dim bPrint As Byte

Dim lCurrSo As Long

Dim sStName As String
Dim sStAdr As String
Dim sDivision As String
Dim sRegion As String
Dim sSterms As String
Dim sVia As String
Dim sFob As String
Dim sSlsMan As String
Dim sContact As String
Dim sConIntPhone As String
Dim sConPhone As String
Dim sConIntFax As String
Dim sConFax As String
Dim sConExt As String
Dim cDiscount As String
Dim iDays As String
Dim iNetDays As String
Dim iFrtDays As String
Dim sTaxExempt As String


Dim sOldPrefix As String
Dim sOldCustomer As String
Dim sOldSpo As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function CheckCustomerPO() As Byte
   Dim RdoCpo As ADODB.Recordset
   If Trim(txtCpo) = "" Then
      CheckCustomerPO = 0
   Else
      sSql = "SELECT DISTINCT SOCUST,SOPO FROM SohdTable WHERE " _
             & "(SOCUST='" & Compress(cmbCst) & "' AND SOPO='" _
             & Trim(txtCpo) & "' AND SONUMBER<>" & Val(cmbSon) & " )"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCpo, ES_FORWARD)
      If bSqlRows Then
         With RdoCpo
            CheckCustomerPO = 1
            ClearResultSet RdoCpo
         End With
      End If
   End If
   Set RdoCpo = Nothing
   
End Function


Private Sub ClearCombos()
   Dim iControl As Integer
   For iControl = 0 To Controls.Count - 1
      If TypeOf Controls(iControl) Is ComboBox Then _
                         Controls(iControl).SelLength = 0
   Next
   
End Sub

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

Private Sub GetSalesPerson()
   Dim rdoSlp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetSalesPerson '" & cmbSlp & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoSlp)
   If bSqlRows Then
      cmbSlp = "" & Trim(rdoSlp!SPNumber)
      lblSlp = "" & Trim(rdoSlp!SPFIRST) & " " & Trim(rdoSlp!SPLAST)
      If Trim(cmbReg) = "" Then cmbReg = "" & Trim(rdoSlp!SPREGION)
      ClearResultSet rdoSlp
   Else
      lblSlp = ""
   End If
   On Error Resume Next
   Set rdoSlp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsalesp"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
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

Private Sub FillShipVia()
   On Error GoTo DiaErr1
   sSql = "Qry_FillSPVia"
   LoadComboBox cmbShpVia, 0
   Exit Sub
   
DiaErr1:
   sProcName = "FillShipVia"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FormatControls()
   Dim RdoRows As ADODB.Recordset
   Dim b As Byte
   Dim iRows As Integer
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtBto.BackColor = Es_TextDisabled
   txtDdt.ToolTipText = "Includes Default Freight Days"
   
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
      z1(10).Visible = False
      z1(31).Visible = False
   End If
   Set RdoRows = Nothing
   
End Sub

Private Sub chkITAREAR_Click()
   Dim checked As Integer
   checked = chkITAREAR.Value
   If bGoodSO Then
       RdoRso!SOITAREAR = checked
         RdoRso.Update
   End If
End Sub

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst, False
   GetPriceBook Me
   GetBillTo
      
End Sub

Private Sub cmbCst_LostFocus()
   Dim sNewCust As String
   Dim bGoodCus As Boolean
   cmbCst = CheckLen(cmbCst, 10)
   If bPrint = 1 Then Exit Sub
   
   FindCustomer Me, cmbCst, False
   sNewCust = Compress(cmbCst)
   
   If Trim(sNewCust) <> sOldCustomer Then
      If bGoodSO Then
         On Error Resume Next
         
         Dim bResponse As Byte
         bResponse = MsgBox("Do you wish to change the Customer name ?", _
                  ES_NOQUESTION, Caption)
         If bResponse = vbNo Then
            cmbCst = sOldCustomer
            FindCustomer Me, cmbCst, False
            Exit Sub
         End If
   
         'RdoRso.Edit
         bGoodCus = GetCustomerData(sNewCust)
         
         If (bGoodCus = True) Then
            
'            Dim strTableName As String
'            Dim strTransID As String
'            strTableName = "AuditSoTable"
'            strTransID = GetAuditTransID(strTableName)
'            ' Audit SOhdTable
'            SaveSOHistory strTransID, "UPDATE", "SOSTNAME", sStName, "" & RdoRso!SOSTNAME
'            SaveSOHistory strTransID, "UPDATE", "SOSTADR", sStAdr, "" & RdoRso!SOSTADR
'            SaveSOHistory strTransID, "UPDATE", "SODIVISION", sDivision, "" & RdoRso!SODIVISION
'            SaveSOHistory strTransID, "UPDATE", "SOCUST", sNewCust, "" & RdoRso!SOCUST
'            SaveSOHistory strTransID, "UPDATE", "SOREGION", sRegion, "" & RdoRso!SOREGION
'            SaveSOHistory strTransID, "UPDATE", "SOSALESMAN", sSlsMan, "" & RdoRso!SOSALESMAN
'            SaveSOHistory strTransID, "UPDATE", "SOARDISC", cDiscount, "" & RdoRso!SOARDISC
'            SaveSOHistory strTransID, "UPDATE", "SODAYS", CStr(iNetDays), "" & RdoRso!SODAYS
'
'            SaveSOHistory strTransID, "UPDATE", "SOCCONTACT", sContact, "" & RdoRso!SOCCONTACT
'            SaveSOHistory strTransID, "UPDATE", "SOCPHONE", sConPhone, "" & RdoRso!SOCPHONE
'            SaveSOHistory strTransID, "UPDATE", "SOCINTPHONE", sConIntPhone, "" & RdoRso!SOCINTPHONE
            
            RdoRso!SOCUST = sNewCust
            RdoRso!SOSTNAME = sStName
            RdoRso!SOSTADR = sStAdr
            RdoRso!SODIVISION = sDivision
            RdoRso!SOREGION = sRegion
            RdoRso!SOSTERMS = sSterms
            RdoRso!SOVIA = sVia
            RdoRso!SOFOB = sFob
            RdoRso!SOSALESMAN = sSlsMan
            RdoRso!SOARDISC = cDiscount
            RdoRso!SODAYS = iDays
            RdoRso!SONETDAYS = iNetDays
            RdoRso!SOFREIGHTDAYS = iFrtDays
            RdoRso!SOTAXEXEMPT = sTaxExempt
            RdoRso!SOCCONTACT = sContact
            RdoRso!SOCPHONE = sConPhone
            RdoRso!SOCINTPHONE = sConIntPhone
            RdoRso!SOCINTFAX = sConIntFax
            RdoRso!SOCFAX = sConFax
            RdoRso!SOCEXT = sConFax
         
            RdoRso.Update
            
            GetSalesOrder (False)
            
            ' Check for custom credit limit warning
            CheckCusCreditLmt sNewCust
            
            
         End If
         
         If Err > 0 Then
            If Err = 40088 Then
               RefreshSoCursor
               'RdoRso.Edit
               RdoRso!SOCUST = sNewCust
               RdoRso.Update
            Else
               ValidateEdit
            End If
         End If
      
      If bCanceled Then Exit Sub
      GetPriceBook Me
      GetBillTo
      
      End If
      
   End If
   
End Sub

Private Function GetCustomerData(ByVal strCust As String) As Byte
   Dim RdoCst As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT CUREF,CUSTNAME,CUSTNAME,CUSTADR,CUARDISC," _
          & "CUDAYS,CUNETDAYS,CUDIVISION,CUREGION,CUSTERMS," _
          & "CUVIA,CUFOB,CUSALESMAN,CUDISCOUNT,CUSTSTATE," _
          & "CUSTCITY,CUSTZIP,CUCCONTACT,CUCPHONE,CUCEXT,CUCINTPHONE," _
          & "CUFRTDAYS,CUINTFAX,CUFAX,CUTAXEXEMPT,CUCUTOFF " _
          & "FROM CustTable WHERE CUREF='" & strCust & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst)
   If bSqlRows Then
      With RdoCst
         bCutOff = "" & !CUCUTOFF
         sStName = "" & Trim(!CUSTNAME)
         sStAdr = "" & Trim(!CUSTADR) & vbCrLf _
                  & "" & Trim(!CUSTCITY) & " " & Trim(!CUSTSTATE) _
                  & "  " & Trim(!CUSTZIP)
         sDivision = "" & Trim(!CUDIVISION)
         sRegion = "" & Trim(!CUREGION)
         sSterms = "" & Trim(!CUSTERMS)
         sVia = "" & Trim(!CUVIA)
         sFob = "" & Trim(!CUFOB)
         sSlsMan = "" & Trim(!CUSALESMAN)
         sContact = "" & Trim(!CUCCONTACT)
         sConIntPhone = "" & Trim(!CUCINTPHONE)
         sConPhone = "" & Trim(!CUCPHONE)
         sConIntFax = "" & Trim(!CUINTFAX)
         sConFax = "" & Trim(!CUFAX)
         sConFax = "" & Trim(str$(!CUCEXT))
         cDiscount = Format(0 + !CUARDISC, "##0.000")
         iDays = Format("" & !CUDAYS, "###0")
         iNetDays = Format("" & !CUNETDAYS, "###0")
         iFrtDays = Format("" & !CUFRTDAYS, "##0")
         sTaxExempt = "" & Trim("" & !CUTAXEXEMPT)
         ClearResultSet RdoCst
      End With
      GetCustomerData = True
   Else
      sStName = ""
      sStAdr = ""
      sDivision = ""
      sRegion = ""
      sSterms = ""
      sVia = ""
      sFob = ""
      sSlsMan = ""
      iFrtDays = 0
      MsgBox "Couldn't Retrieve Customer.", vbExclamation, Caption
      GetCustomerData = False
   End If
   Set RdoCst = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcustda"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbDiv_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub cmbDiv_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbDiv = CheckLen(cmbDiv, 4)
   If cmbDiv.ListCount > 0 Then
      For iList = 0 To cmbDiv.ListCount - 1
         If cmbDiv.List(iList) = cmbDiv Then bByte = 1
      Next
      If bByte = 0 Then
         Beep
         cmbDiv = cmbDiv.List(0)
      End If
   Else
      If Len(cmbDiv) > 0 Then
         Beep
         cmbDiv = ""
      End If
   End If
   If bGoodSO Then
      On Error Resume Next
'      ' Audit the column
'      SaveSOHistory "", "UPDATE", "SODIVISION", CStr(cmbDiv), RdoRso!SODIVISION
      
      'RdoRso.Edit
      RdoRso!SODIVISION = "" & cmbDiv
      RdoRso.Update
      
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SODIVISION = "" & cmbDiv
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub


'Private Sub cmbPONumber_Click()
'    Dim sPONumber As String
'    sPONumber = cmbPONumber & ""
'
'    GetPONumberForSalesOrder (sPONumber)
'End Sub

'Private Sub cmbPONumber_LostFocus()
'    Dim sPONumber As String
'    sPONumber = cmbPONumber & ""
'
'    GetPONumberForSalesOrder (sPONumber)
'End Sub

Private Sub cmbPre_LostFocus()
   Dim a As Integer
   cmbPre = CheckLen(cmbPre, 1)
   On Error Resume Next
   a = Asc(Left(cmbPre, 1))
   If a < 65 Or a > 90 Then
      MsgBox "Must Be Between A and Z..", vbInformation, Caption
      cmbPre = sOldPrefix
   End If
   If Len(Trim(cmbPre)) = 0 Then cmbPre = "S"
   
   If bGoodSO Then
      On Error Resume Next
'      ' Aduit the column
'      SaveSOHistory "", "UPDATE", "SOPREFIX", CStr(sOldPrefix), "" & RdoRso!SOPREFIX
'      SaveSOHistory "", "UPDATE", "SOPREFDATE", Format(ES_SYSDATE, "mm/dd/yyyy"), "" & RdoRso!SOPREFDATE
'      Debug.Print Err
      
      'RdoRso.Edit
      RdoRso!SOTYPE = cmbPre
      If sOldPrefix <> cmbPre Then
         RdoRso!SOPREFIX = sOldPrefix
         RdoRso!SOPREFDATE = Format(ES_SYSDATE, "mm/dd/yyyy")
      End If
      RdoRso.Update
      Debug.Print Err
      
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub

Private Sub cmbReg_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub cmbReg_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbReg = CheckLen(cmbReg, 2)
   If cmbReg.ListCount > 0 Then
      For iList = 0 To cmbReg.ListCount - 1
         If cmbReg.List(iList) = cmbReg Then bByte = 1
      Next
      If bByte = 0 Then
         Beep
         cmbReg = cmbReg.List(0)
      End If
   Else
      If Len(cmbReg) > 0 Then
         Beep
         cmbReg = ""
      End If
   End If
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOREGION", CStr(cmbReg), "" & RdoRso!SOREGION
      
      'RdoRso.Edit
      RdoRso!SOREGION = "" & cmbReg
      RdoRso.Update
      
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOREGION = "" & cmbReg
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub




Private Sub cmbSlp_Click()
   GetSalesPerson
   
End Sub

Private Sub cmbSlp_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub cmbSlp_LostFocus()
   cmbSlp = CheckLen(cmbSlp, 4)
   If bGoodSO Then
      If Len(Trim(cmbSlp)) Then
         On Error Resume Next
         'RdoRso.Edit
         RdoRso!SOSALESMAN = "" & cmbSlp
         RdoRso.Update
         If Err > 0 Then
            If Err = 40088 Then
               RefreshSoCursor
               ' Aduit the column
               'SaveSOHistory "", "UPDATE", "SOSALESMAN", CStr(cmbSlp), "" & RdoRso!SOSALESMAN
               
               'RdoRso.Edit
               RdoRso!SOSALESMAN = "" & cmbSlp
               RdoRso.Update
            Else
               ValidateEdit
            End If
         End If
      End If
   End If
   
End Sub


Private Sub cmbSon_Click()
   If lCurrSo <> Val(cmbSon) Then
      bGoodSO = GetSalesOrder(False)
      ' Check to see if the SO customer could be changed.
      cmbCst.Enabled = CheckCustToModify(Val(cmbSon))
   End If
End Sub

Private Sub cmbSon_GotFocus()
   
   cmbSon_Click
   
End Sub

Private Function CheckCustToModify(ByVal strSoNum As String)
   If (strSoNum <> "") Then
      Dim RdoSO As ADODB.Recordset
      
      sSql = "SELECT ITACTUAL, ITPSNUMBER FROM soitTable " & _
               " WHERE ITSO = '" & strSoNum & "' AND (ITACTUAL <> NULL " & _
               " OR ITPSNUMBER <> '' OR ITINVOICE <> 0)"
      
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoSO, ES_FORWARD)
      If bSqlRows Then
         CheckCustToModify = False
      Else
         CheckCustToModify = True
      End If
   End If

End Function

Private Sub cmbSon_LostFocus()
   cmbSon = CheckLen(cmbSon, SO_NUM_SIZE)
   cmbSon = Format(Abs(Val(cmbSon)), SO_NUM_FORMAT)
   If bCanceled Then Exit Sub
   If bPrint Then Exit Sub
   If bItems = 0 Then
      bGoodSO = GetSalesOrder(True)
      ' Save the last sales order revised so we can use it later (elsewhere)
      SaveSetting "Esi2000", "EsiSale", "LastRevisedSO", cmbSon
      ' Check to see if the SO customer could be changed.
      cmbCst.Enabled = CheckCustToModify(Val(cmbSon))
   End If
   bItems = 0
   
End Sub

Private Sub cmbTrm_Click()
   GetTerms
   
End Sub

Private Sub cmbTrm_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub cmbTrm_LostFocus()
   cmbTrm = CheckLen(cmbTrm, 2)
   If bGoodSO Then
      On Error Resume Next
      'RdoRso.Edit
      RdoRso!SOSTERMS = "" & cmbTrm
      RdoRso.Update
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            ' Aduit the column
            'SaveSOHistory "", "UPDATE", "SOSTERMS", CStr(cmbTrm), "" & RdoRso!SOSTERMS
            
            'RdoRso.Edit
            RdoRso!SOSTERMS = "" & cmbTrm
            RdoRso.Update
            
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdComments_Click(Index As Integer)
   If Index = (0) Then
      If cmdComments(0) Then
         txtCmt.SetFocus
         'See List For Index
         SysComments.lblListIndex = 2
         SysComments.Show
         cmdComments(0) = False
      End If
   Else
      If cmdComments(1) Then
         txtStAdr.SetFocus
         'See List For Index
         SysComments.lblListIndex = 2
         SysComments.Show
         cmdComments(1) = False
      End If
   End If
End Sub

Private Sub cmdFind_Click()
    SOLookup.lblControl = "cmbSon"
    SOLookup.lblSONumber = cmbSon.Text
    SOLookup.lblSoType = cmbPre.Text
    SOLookup.Show
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2102
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdItm_Click()
   If bCutOff Then
      MsgBox "This Customer's Credit Is On Hold.", _
         vbInformation, Caption
   Else
      UpdateSalesOrder
      bGoodSO = GetSalesOrder(True)
      If bGoodSO Then
         ' Aduit the column
         'SaveSOHistory "", "UPDATE", "SOREVISED", Format(ES_SYSDATE, "mm/dd/yyyy"), "" & RdoRso!SOREVISED
         ' On Error Resume Next
         'RdoRso.Edit
         RdoRso!SOREVISED = Format(ES_SYSDATE, "mm/dd/yyyy")
         RdoRso.Update
         
         optItm.Value = vbChecked
         optNew.Value = vbUnchecked
         SaleSLe02b.Caption = SaleSLe02b.Caption & " - " & cmbPre & cmbSon
         SaleSLe02b.Show
      End If
   End If
   bItems = 0
   
End Sub

Private Sub cmdItm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bItems = 1
   
End Sub


Private Sub cmdNew_Click()
   optNew = vbChecked
   SaleSLe01a.optRev = vbChecked
   SaleSLe01a.Show
   
End Sub


Private Sub cmdNte_Click()
   optNotes.Value = vbChecked
   SaleSLe02d.lblCustomer = cmbCst
   SaleSLe02d.Show
   
End Sub


Private Sub cmdStatCode_Click()
 StatusCode.lblSCTypeRef = "SO Number"
 StatusCode.txtSCTRef = cmbSon.Text
 StatusCode.lblSCTRef1 = ""
 StatusCode.lblSCTRef2 = ""
 StatusCode.lblStatType = "SO"
 StatusCode.lblSysCommIndex = 2 ' The index in the Sys Comment "SO"
 StatusCode.txtCurUser = cUR.CurrentUser

 StatusCode.Show
End Sub

Private Sub cmdVew_Click()
   If cmdVew.Value = True Then
      '        SoTree.Show
      '        cmdVew.Value = False
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bDataHasChanged = False
      CheckPriceBookDates
      cmdComments(0).Enabled = True
      cmdComments(1).Enabled = True
      MouseCursor 13
      bPrint = 0
      
      FillSales
      ' Ship via
      FillShipVia
      ' 2/15/2009 Passing the frm to fill the combo box
      FillTerms Me
      If cmbTrm.ListCount > 0 Then cmbTrm = cmbTrm.List(0)
      ' 2/15/2009 Passing the frm to fill the combo box
      FillCustomers Me
      ' 2/15/2009 Passing the frm to fill the combo box
      FillDivisions Me
      ' 2/15/2009 Passing the frm to fill the combo box
      FillRegions Me
      FillCombo
      bOnLoad = 0
   End If
   If optNew = vbChecked And optNotes.Value = vbUnchecked Then
      If SaleSLe01a.optNew = vbChecked Then
         cmbSon = SaleSLe01a.txtSon
         cmbPre = SaleSLe01a.cmbPre
      End If
      Unload SaleSLe01a
   End If
   

   If OptSoXml = vbChecked Then
      cmbSon = SaleSLf12a.txtSon
      cmbPre = SaleSLf12a.cmbPre
      OptSoXml = vbUnchecked
      SaleSLf12a.OptSoXml = vbChecked
      Unload SaleSLf12a
   End If
   
   If OptSoEDI = vbChecked Then
      cmbSon = SaleSLf14a.txtSon
      cmbPre = SaleSLf14a.cmbPre
      OptSoEDI = vbUnchecked
      SaleSLf14a.OptSoEDI = vbChecked
      Unload SaleSLf14a
   End If
   
   
   cmbCst.Enabled = CheckCustToModify(Val(cmbSon))
   optNotes.Value = vbUnchecked
   If optItm = vbChecked Then
      optItm = vbUnchecked
      Unload SaleSLe02b
      bGoodSO = GetSalesOrder(True)
   End If
   If optLst = vbChecked Then
      cmbPre = SaleSLp07a.txtPre
      cmbSon = SaleSLp07a.txtSon
      optLst = vbUnchecked
      Unload SaleSLp07a
   End If
   
End Sub

Private Sub Form_Load()
   Dim iList As Integer
   FormLoad Me
   FormatControls
   sSql = "SELECT * FROM SohdTable WHERE SONUMBER= ? "
   'Set rdoQry = RdoCon.CreateQuery("", sSql)
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql
   
   Dim prmObj As ADODB.Parameter
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adInteger
   cmdObj.parameters.Append prmObj
   
   For iList = 0 To 2
      With tabFrame(iList)
         .BorderStyle = 0
         .Visible = False
         .Left = 40
      End With
   Next
   tabFrame(0).Visible = True
   bOnLoad = 1
   txtCan.ForeColor = ES_RED
   cmbPre = "S"
   For iList = 65 To 90
      AddComboStr cmbPre.hWnd, Chr$(iList)
   Next
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bGoodSO = 1 And bDataHasChanged Then UpdateSalesOrder
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload bPrint
   Set cmdObj = Nothing
   Set RdoRso = Nothing
   Set SaleSLe02a = Nothing
   
End Sub



Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim iList As Integer
   'Dim sYear As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   cmbSon.Clear
   'doesn't work on leap day
   'iList = Format(Now, "yyyy")
   'iList = iList - 3
   'sYear = Trim$(iList) & "-" & Format(Now, "mm-dd")
   'sSql = "Qry_FillSalesOrders '" & sYear & "'"
   
'   'this is valid sql but doesn't work through rdo
'   sSql = "declare @dt as smalldatetime;" & vbCrLf _
'      & "set @dt = dateadd(year,-3,getdate());" & vbCrLf _
'      & "Qry_FillSalesOrders @dt;"
      
   sSql = "Qry_FillSalesOrders '" & DateAdd("yyyy", -3, Now) & "'"
      
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      iList = -1
      With RdoCmb
         cmbPre = "" & Trim(!SOTYPE)
         cmbSon = Format(!SoNumber, SO_NUM_FORMAT)
         Do Until .EOF
            iList = iList + 1
            If iList > 999 Then Exit Do
            AddComboStr cmbSon.hWnd, Format$(!SoNumber, SO_NUM_FORMAT)
           ' AddComboStr Me.cmbPONumber.hwnd, Trim(!SOPO) 'BBS Added on 03/22/2010 for Ticket #28317
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   Else
      MouseCursor 0
      MsgBox "No Sales Orders Where Found.", vbInformation, Caption
      Exit Sub
   End If
   Set RdoCmb = Nothing
   MouseCursor 0
   If cmbSon.ListCount > 0 Then bGoodSO = GetSalesOrder(False)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Function GetSalesOrder(Optional bOpen As Byte) As Byte
   Dim iDelDays As Integer
   Dim vDelDate As Date
   
   If bGoodSO = 1 And bDataHasChanged Then UpdateSalesOrder
   bGoodSO = 0
   On Error GoTo DiaErr1
   cmdItm.Enabled = False
   GetSalesOrder = 0
   MouseCursor 13
   txtCph.Mask = ""
   txtCfax.Mask = ""
   'rdoQry.RowsetSize = 1
   'rdoQry(0) = Val(cmbSon)
   
   cmdObj.parameters(0).Value = Val(cmbSon)
   bSqlRows = clsADOCon.GetQuerySet(RdoRso, cmdObj, ES_KEYSET, True)
   If bSqlRows Then
      With RdoRso
         If !SOLOCKED = 1 Then
            MsgBox "This Sales Order Is Locked.", _
               vbInformation, Caption
            FillCombo
            Exit Function
         End If
         cmbSon = Format("" & !SoNumber, SO_NUM_FORMAT)
         
         
        ' cmbPONumber = Trim(!SOPO) 'BBS Added on 03/22/2010 for Ticket #28317
         
         cmbPre = "" & Trim(!SOTYPE)
         sOldPrefix = "" & Trim(!SOTYPE)
         txtBok = Format("" & !SODATE, "mm/dd/yyyy")
         txtShp = Format("" & !SOSHIPDATE, "mm/dd/yyyy")
         If Len(Trim(txtShp)) = 0 Then txtShp = Format(ES_SYSDATE, "mm/dd/yyyy")
         txtCpo = "" & Trim(!SOPO)
         sOldSpo = txtCpo
         txtDdt = Format("" & !SODELDATE, "mm/dd/yyyy")
         If Len(Trim(txtDdt)) = 0 Then
            If Not IsNull(!SOFREIGHTDAYS) Then _
                          iDelDays = !SOFREIGHTDAYS
            vDelDate = Format(ES_SYSDATE, "mm/dd/yyyy")
            vDelDate = vDelDate + iDelDays
            txtDdt = Format(vDelDate, "mm/dd/yyyy")
         End If
         txtDsc = Format("" & !SOARDISC, "#0.00")
         txtDay = Format("" & !SODAYS, "##0")
         txtNet = Format("" & !SONETDAYS, "#0")
         optTax.Value = "" & !SOTAXABLE
         txtTax = "" & Trim(!SOTAXEXEMPT)
         'optCom.Value = !SOCYN
         'If optCom.Value = vbChecked Then txtCms.Enabled = True Else txtCms.Enabled = False
         'txtCms = Format(!SOCOMMISSION, "#####0.00")
         cmbDiv = "" & Trim(!SODIVISION)
         cmbReg = "" & Trim(!SOREGION)
         txtBus = "" & Trim(!SOBUSUNIT)
         cmbSlp = "" & Trim(!SOSALESMAN)
         'JA 6/20/99 SOREP is no longer used by MCS (SOSALESMAN)
         'txtRep = "" & Trim(!SOREP)
         txtJob = "" & Trim(!SOJOBNO)
         txtJnm = "" & Trim(!SONAME)
         txtStr = "" & Trim(!SOTERMS)
         cmbTrm = "" & Trim(!SOSTERMS)
         txtFob = "" & Trim(!SOFOB)
         txtFrd = Format("" & !SOFREIGHTDAYS, "##0")
         txtVia = "" & Trim(!SOVIA)
         cmbShpVia = "" & Trim(!SOVIA)
         
         txtStNme = "" & Trim(!SOSTNAME)
         txtStAdr = "" & Trim(!SOSTADR)
         txtCmt = "" & Trim(!SOREMARKS)
         txtExt = Format("" & !SOCEXT, "###0")
         If Trim(!SOCUST) <> sOldCustomer Then FindCustomer Me, "" & Trim(!SOCUST), True
         sOldCustomer = "" & Trim(!SOCUST)
         txtCnt = "" & Trim(!SOCCONTACT)
         ' 2/25/2009 Uncommented this as the BillTo field is not populated.
         GetBillTo
         
         GetSalesPerson
         If !SOCANCELED = 0 Then
            txtCan = ""
            lblCan.Visible = False
            txtCan.Visible = False
            cmdItm.Enabled = True
         Else
            cmdItm.Enabled = False
            txtCan = Format(!SOCANDATE, "mm/dd/yyyy")
            lblCan.Visible = True
            txtCan.Visible = True
         End If
         If bOpen Then
            If lblCan.Visible = False Then cmdItm.Enabled = True
            On Error Resume Next
            '                    If cmbCst.Enabled Then
            '                        cmbCst.SetFocus
            '                    Else
            '                        txtBok.SetFocus
            '                    End If
         End If
         '7/21/04 ***
         txtCcInt = "" & Trim(!SOCINTPHONE)
         txtCph = "" & Trim(!SOCPHONE)
         txtCintFax = "" & Trim(!SOCINTFAX)
         txtCfax = "" & Trim(!SOCFAX)
         txtCph.Mask = "###-###-####"
         txtCfax.Mask = "###-###-####"
         '****
         
         Me.chkITAREAR.Value = IIf(!SOITAREAR, 1, 0)
         
         GetTerms
         GetPriceBook Me
         GetSalesOrder = 1
         lCurrSo = Val(cmbSon)
      End With
   Else
      On Error Resume Next
      lCurrSo = 0
      cmdItm.Enabled = False
      GetSalesOrder = 0
      cmdItm.Enabled = False
      If Not bOnLoad Then MsgBox "Sales Order Wasn't Found Or Was Canceled.", vbExclamation, Caption
      cmbSon.ListIndex = 0
   End If
   bDataHasChanged = False
   If GetSalesOrder = 1 Then GetItems
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getsaleso"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



Private Sub lblcst_Change()
   cmbCst = lblCst
   
End Sub


Private Sub lblTrm_Change()
   lblTrm.ToolTipText = lblTrm
   
End Sub






Private Sub optItm_Click()
   'never visible-SaleSLe02b is loaded
   
End Sub

Private Sub optLst_Click()
   'never visible-SaleSLp07a is loaded
   
End Sub


Private Sub optNew_Click()
   'never visible-SaleSLe01a is loaded
   
End Sub

Private Sub optPrn_Click()
   If Val(cmbSon) > 0 Then
      WindowState = vbMinimized
      bPrint = True
      SaleSLp01a.Show
      SaleSLp01a.txtBeg = cmbSon
      Unload Me
   End If
   
End Sub


Private Sub optSvw_Click()
   If optSvw.Value = vbChecked Then
      cmbSon_Click
      optSvw.Value = vbUnchecked
      On Error Resume Next
      cmbSon.SetFocus
   End If
   
End Sub

Private Sub optTax_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optTax_LostFocus()
   If bGoodSO Then
      On Error Resume Next
      'RdoRso.Edit
      RdoRso!SOTAXABLE = optTax.Value
      RdoRso.Update
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            
            ' Aduit the column
            'SaveSOHistory "", "UPDATE", "SOTAXABLE", CStr(optTax.Value), "" & RdoRso!SOTAXABLE
         
            'RdoRso.Edit
            RdoRso!SOTAXABLE = optTax.Value
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub

Private Sub tab1_Click()
   Dim b As Byte
   On Error Resume Next
   ClearCombos
   For b = 0 To 2
      tabFrame(b).Visible = False
   Next
   tabFrame(tab1.SelectedItem.Index - 1).Visible = True
   Select Case tab1.SelectedItem.Index
      Case 1
         txtBok.SetFocus
      Case 2
         cmbTrm.SetFocus
      Case 3
         txtCmt.SetFocus
   End Select
   
End Sub

Private Sub txtBok_DropDown()
   Debug.Print "Visible=" & txtShp.Visible & " Left=" & txtShp.Left & " top=" & txtShp.Top & " frmwidth=" & Me.ScaleWidth & " frmheight=" & Me.ScaleHeight
   Debug.Print txtShp.Parent.Name
   Debug.Print cmbDiv.Parent.Name
   ShowCalendarEx Me, tab1.Top
   bDataHasChanged = True
   
End Sub

Private Sub txtBok_LostFocus()
   txtBok = CheckDateEx(txtBok)
   If bGoodSO Then
      On Error Resume Next
      
      ' Aduit the column
'      SaveSOHistory "", "UPDATE", "SODATE", Format(txtBok, "mm/dd/yyyy"), "" & RdoRso!SODATE
'      SaveSOHistory "", "UPDATE", "SOTEXT", Format(Val(cmbSon), SO_NUM_FORMAT), "" & RdoRso!SOTEXT
            
      'RdoRso.Edit
      RdoRso!SODATE = Format(txtBok, "mm/dd/yyyy")
      '@@@SOTEXT now computed -- RdoRso!SOTEXT = Format(Val(cmbSon), SO_NUM_FORMAT)
      RdoRso.Update
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SODATE = Format(txtBok, "mm/dd/yyyy")
            '@@@SOTEXT now computed -- RdoRso!SOTEXT = Format(Val(cmbSon), SO_NUM_FORMAT)
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub

Private Sub txtBook_Click()
   ' SaleSLe02c.Show
   
End Sub


Private Sub txtBus_LostFocus()
   txtBus = CheckLen(txtBus, 4)
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOBUSUNIT", txtBus, "" & RdoRso!SOBUSUNIT
      
      'RdoRso.Edit
      RdoRso!SOBUSUNIT = "" & txtBus
      RdoRso.Update
      
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOBUSUNIT = "" & txtBus
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub


Private Sub txtCcInt_Change()
   If Len(txtCcInt) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtCcInt_LostFocus()
   txtCcInt = CheckLen(txtCcInt, 4)
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOCINTPHONE", txtCcInt, "" & RdoRso!SOCINTPHONE
            
      'RdoRso.Edit
      RdoRso!SOCINTPHONE = "" & txtCcInt
      RdoRso.Update
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOCINTPHONE = "" & txtCcInt
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub


Private Sub txtCfax_LostFocus()
   txtCfax = CheckLen(txtCfax, 12)
   If bGoodSO Then
      If Len(Trim(txtCnt)) Then
         On Error Resume Next
         ' Aduit the column
         'SaveSOHistory "", "UPDATE", "SOCFAX", txtCfax, "" & RdoRso!SOCFAX
               
         'RdoRso.Edit
         RdoRso!SOCFAX = "" & txtCfax
         RdoRso.Update
         If Err > 0 Then
            If Err = 40088 Then
               RefreshSoCursor
      
               'RdoRso.Edit
               RdoRso!SOCFAX = "" & txtCfax
               RdoRso.Update
            Else
               ValidateEdit
            End If
         End If
      End If
   End If
   
End Sub


Private Sub txtCintFax_Change()
   If Len(txtCintFax) > 3 Then SendKeys "{tab}"
   
End Sub


Private Sub txtCintFax_LostFocus()
   txtCintFax = CheckLen(txtCintFax, 4)
   If bGoodSO Then
      If Len(Trim(txtCnt)) Then
         On Error Resume Next
         'RdoRso.Edit
         RdoRso!SOCINTFAX = "" & txtCintFax
         RdoRso.Update
         If Err > 0 Then
            If Err = 40088 Then
               RefreshSoCursor
      
               ' Aduit the column
               'SaveSOHistory "", "UPDATE", "SOCINTFAX", txtCintFax, "" & RdoRso!SOCINTFAX
               
               'RdoRso.Edit
               RdoRso!SOCINTFAX = "" & txtCintFax
               RdoRso.Update
            Else
               ValidateEdit
            End If
         End If
      End If
   End If
   
End Sub




Private Sub txtCmt_Change()
   If bGoodSO Then bDataHasChanged = True
   
End Sub

Private Sub txtCmt_LostFocus()
    ' 11/3/2009 increased the size
   'txtCmt = CheckLen(txtCmt, 2560)
   'txtCmt = CheckLen(txtCmt, 6000)
      
   txtCmt = CheckComments(txtCmt)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   
   If bGoodSO Then
      On Error Resume Next
      
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOREMARKS", RTrim(txtCmt), "" & RdoRso!SOREMARKS
      
      'RdoRso.Edit
      RdoRso!SOREMARKS = RTrim(txtCmt)
      RdoRso.Update
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOREMARKS = RTrim(txtCmt)
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub


Private Sub txtCnt_LostFocus()
   txtCnt = CheckLen(txtCnt, 20)
   txtCnt = StrCase(txtCnt)
   If bGoodSO Then
      If Len(Trim(txtCnt)) Then
         On Error Resume Next
         ' Aduit the column
         'SaveSOHistory "", "UPDATE", "SOCCONTACT", RTrim(txtCnt), "" & RdoRso!SOCCONTACT
               
         'RdoRso.Edit
         RdoRso!SOCCONTACT = "" & txtCnt
         RdoRso.Update
         If Err > 0 Then
            If Err = 40088 Then
      
               'RdoRso.Edit
               RdoRso!SOCCONTACT = "" & txtCnt
               RdoRso.Update
               RefreshSoCursor
            Else
               ValidateEdit
            End If
         End If
      End If
   End If
   
End Sub




Private Sub txtCph_LostFocus()
   txtCph = CheckLen(txtCph, 12)
   If bGoodSO Then
      If Len(Trim(txtCnt)) Then
         On Error Resume Next
         ' Aduit the column
         'SaveSOHistory "", "UPDATE", "SOCPHONE", RTrim(txtCph), "" & RdoRso!SOCPHONE
               
         'RdoRso.Edit
         RdoRso!SOCPHONE = "" & txtCph
         RdoRso.Update
         If Err > 0 Then
            If Err = 40088 Then
               RefreshSoCursor
      
               'RdoRso.Edit
               RdoRso!SOCPHONE = "" & txtCph
               RdoRso.Update
            Else
               ValidateEdit
            End If
         End If
      End If
   End If
   
End Sub


Private Sub txtCpo_LostFocus()
   Dim bByte As Byte
   txtCpo = CheckLen(txtCpo, 20)
   If sOldSpo <> txtCpo Then bByte = CheckCustomerPO()
   If bByte = 1 Then
      bByte = MsgBox("That Customer PO Is In Use. Proceed Anyway?", _
              ES_NOQUESTION, Caption)
      If bByte = vbNo Then txtCpo = sOldSpo
   End If
   If bGoodSO Then
      On Error Resume Next
      
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOPO", RTrim(txtCpo), "" & RdoRso!SOPO
      
      'RdoRso.Edit
      RdoRso!SOPO = "" & txtCpo
      RdoRso.Update
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOPO = "" & txtCpo
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   sOldSpo = txtCpo
   
End Sub

Private Sub txtDay_LostFocus()
   txtDay = CheckLen(txtDay, 3)
   txtDay = Format(Abs(Val(txtDay)), "##0")
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SODAYS", RTrim(txtDay), "" & RdoRso!SODAYS
            
      'RdoRso.Edit
      RdoRso!SODAYS = Val(txtDay)
      RdoRso.Update
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SODAYS = Val(txtDay)
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub

Private Sub txtDdt_DropDown()
   ShowCalendarEx Me, tab1.Top
   bDataHasChanged = True
   
End Sub

Private Sub txtDdt_LostFocus()
   Dim vDate As Date
   txtDdt = CheckDateEx(txtDdt)
   If bGoodSO Then
      On Error Resume Next
      
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SODELDATE", Format(txtDdt, "mm/dd/yyyy"), "" & RdoRso!SODELDATE
            
      'RdoRso.Edit
      RdoRso!SODELDATE = Format(txtDdt, "mm/dd/yyyy")
      RdoRso.Update
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SODELDATE = Format(txtDdt, "mm/dd/yyyy")
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   On Error GoTo 0
   vDate = txtDdt
   txtShp = Format(vDate - Val(txtFrd), "mm/dd/yyyy")
   
End Sub

Private Sub txtDis_LostFocus()
   txtDis = CheckLen(txtDis, 5)
   txtDis = Format(Abs(Val(txtDis)), "#0.00")
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOARDISC", txtDis, "" & RdoRso!SOARDISC
            
      'RdoRso.Edit
      RdoRso!SOARDISC = Val(txtDis)
      RdoRso.Update
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOARDISC = Val(txtDis)
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 5)
   txtDsc = Format(Abs(Val(txtDsc)), "#0.00")
   If bGoodSO Then
      On Error Resume Next
      
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOARDISC", txtDis, "" & RdoRso!SOARDISC
            
      'RdoRso.Edit
      RdoRso!SOARDISC = Val(txtDsc)
      RdoRso.Update
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOARDISC = Val(txtDis)
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub


Private Sub txtExt_Change()
   txtExt = CheckLen(txtExt, 4)
   txtExt = Format(Abs(Val(txtExt)), "###0")
   If bGoodSO = 1 Then
      On Error Resume Next
      
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOCEXT", txtExt, "" & RdoRso!SOCEXT
      
      'RdoRso.Edit
      RdoRso!SOCEXT = Val(txtExt)
      RdoRso.Update
   End If
   
End Sub

Private Sub txtFob_LostFocus()
   txtFob = CheckLen(txtFob, 12)
   If bGoodSO Then
      On Error Resume Next
      
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOFOB", txtFob, "" & RdoRso!SOFOB
      
      'RdoRso.Edit
      RdoRso!SOFOB = "" & txtFob
      RdoRso.Update
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOFOB = "" & txtFob
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub

Private Sub txtFra_LostFocus()
   txtFra = CheckLen(txtFra, 10)
   txtFra = Format(Abs(Val(txtFra)), "#0.000")
   
End Sub

Private Sub txtFrd_LostFocus()
   txtFrd = CheckLen(txtFrd, 3)
   txtFrd = Format(Abs(Val(txtFrd)), "##0")
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOFREIGHTDAYS", txtFrd, "" & RdoRso!SOFREIGHTDAYS
      
      'RdoRso.Edit
      RdoRso!SOFREIGHTDAYS = Val(txtFrd)
      RdoRso.Update
      
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOFREIGHTDAYS = Val(txtFrd)
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub

Private Sub txtJnm_LostFocus()
   txtJnm = CheckLen(txtJnm, 20)
   txtJnm = StrCase(txtJnm)
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SONAME", txtJnm, "" & RdoRso!SONAME
      
      'RdoRso.Edit
      RdoRso!SONAME = "" & txtJnm
      RdoRso.Update
      
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SONAME = "" & txtJnm
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub

Private Sub txtJob_Click()
   If Trim(txtJob) = "" Then txtJob = cmbPre & cmbSon
End Sub

Private Sub txtJob_LostFocus()
   txtJob = Compress(txtJob, 6, ES_IGNOREDASHES)
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOJOBNO", txtJob, "" & RdoRso!SOJOBNO
      
      'RdoRso.Edit
      RdoRso!SOJOBNO = "" & txtJob
      RdoRso.Update
      
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOJOBNO = "" & txtJob
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub

Private Sub txtNet_LostFocus()
   txtNet = CheckLen(txtNet, 3)
   txtNet = Format(Abs(Val(txtNet)), "##0")
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SONETDAYS", txtNet, "" & RdoRso!SONETDAYS
      
      'RdoRso.Edit
      RdoRso!SONETDAYS = Val(txtNet)
      RdoRso.Update
      
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SONETDAYS = Val(txtNet)
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub


Private Sub txtNme_Change()
   If txtNme = "*** Customer Wasn't Found ***" Then
      txtNme.ForeColor = ES_RED
   Else
      txtNme.ForeColor = Es_TextForeColor
   End If
   
End Sub






Private Sub txtShp_DropDown()
   ShowCalendarEx Me, tab1.Top
   bDataHasChanged = True
End Sub

Private Sub txtShp_LostFocus()
   txtShp = CheckDateEx(txtShp)
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOSHIPDATE", Format(txtShp, "mm/dd/yyyy"), "" & RdoRso!SOSHIPDATE
      
      'RdoRso.Edit
      RdoRso!SOSHIPDATE = Format(txtShp, "mm/dd/yyyy")
      RdoRso.Update
      
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOSHIPDATE = Format(txtShp, "mm/dd/yyyy")
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub


Private Sub txtStAdr_LostFocus()
   txtStAdr = CheckLen(txtStAdr, 255)
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOSTADR", txtStAdr, "" & RdoRso!SOSTADR
      
      'RdoRso.Edit
      RdoRso!SOSTADR = "" & txtStAdr
      RdoRso.Update
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOSTADR = "" & txtStAdr
            RdoRso.Update
      
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub

Private Sub txtStNme_LostFocus()
   txtStNme = CheckLen(txtStNme, 40)
   txtStNme = StrCase(txtStNme)
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOSTNAME", txtStNme, "" & RdoRso!SOSTNAME
      
      'RdoRso.Edit
      RdoRso!SOSTNAME = "" & txtStNme
      RdoRso.Update
      
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOSTNAME = "" & txtStNme
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub

Private Sub txtStr_LostFocus()
   txtStr = CheckLen(txtStr, 2)
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOTERMS", txtStr, "" & RdoRso!SOTERMS
      
      'RdoRso.Edit
      RdoRso!SOTERMS = "" & txtStr
      RdoRso.Update
      
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOTERMS = "" & txtStr
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub

Private Sub txtTax_LostFocus()
   txtTax = CheckLen(txtTax, 30)
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOTAXEXEMPT", txtTax, "" & RdoRso!SOTAXEXEMPT
      
      'RdoRso.Edit
      RdoRso!SOTAXEXEMPT = "" & txtTax
      RdoRso.Update
      
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOTAXEXEMPT = "" & txtTax
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub

Private Sub cmbShpVia_LostFocus()
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOVIA", cmbShpVia, "" & RdoRso!SOVIA
      
      'RdoRso.Edit
      RdoRso!SOVIA = "" & cmbShpVia
      RdoRso.Update
      
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOVIA = "" & txtVia
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
End Sub

Private Sub txtVia_LostFocus()
   txtVia = CheckLen(txtVia, 30)
   If bGoodSO Then
      On Error Resume Next
      ' Aduit the column
      'SaveSOHistory "", "UPDATE", "SOVIA", txtVia, "" & RdoRso!SOVIA
      
      'RdoRso.Edit
      RdoRso!SOVIA = "" & txtVia
      RdoRso.Update
      
      If Err > 0 Then
         If Err = 40088 Then
            RefreshSoCursor
            'RdoRso.Edit
            RdoRso!SOVIA = "" & txtVia
            RdoRso.Update
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub



Private Sub GetItems()
   Dim RdoCst As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT ITSO FROM SoitTable WHERE ITSO=" & Val(cmbSon) _
          & " AND (ITACTUAL IS NOT NULL AND ITPSNUMBER='' AND ITINVOICE=0 " _
          & "AND ITCANCELED=0)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst)
   If bSqlRows Then
      'cmbCst.Enabled = False
      ClearResultSet RdoCst
   Else
      'cmbCst.Enabled = True
   End If
   Set RdoCst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

' 2/25/2009 - Commented this function
' Edited this function to just show the txt Bill To address
'Private Sub GetBillTo()
'   Dim RdoCbt As ADODB.Recordset
'   Dim sCust As String
'   On Error GoTo DiaErr1
'   sCust = Compress(cmbCst)
'   sSql = "SELECT CUBTNAME,CUBTADR,CUBCITY,CUBSTATE,CUBZIP," _
'          & "CUSTNAME,CUSTADR,CUSTCITY,CUSTSTATE,CUSTZIP,CUTAXEXEMPT," _
'          & "CUCCONTACT,CUCPHONE,CUCEXT,CUCINTPHONE,CUINTFAX,CUFAX," _
'          & "CUDIVISION,CUREGION,CUCUTOFF FROM CustTable WHERE " _
'          & "CUREF='" & sCust & "'"
'   bSqlRows = clsADOCon.GetDataSet(sSql,RdoCbt, ES_FORWARD)
'   If bSqlRows Then
'      With RdoCbt
'         If !CUCUTOFF Then
'            txtNme = "*** This Customer's Credit Is On Hold ***"
'            txtNme.ForeColor = ES_RED
'            bCutOff = 1
'            cmdItm.Enabled = False
'         Else
'            txtNme.ForeColor = vbBlack
'            bCutOff = 0
'         End If
'         txtBto = "" & Trim(!CUBTNAME) & vbCrLf _
'                  & "" & Trim(!CUBTADR) & vbCrLf _
'                  & "" & Trim(!CUBCITY) & ", " _
'                  & " " & Trim(!CUBSTATE) & "  " & Trim(!CUBZIP)
'         '4/12/05
'         If sOldCustomer <> Compress(cmbCst) Then
'            txtStNme = "" & Trim(!CUSTNAME)
'            txtStAdr = "" & Trim(!CUSTADR) & vbCrLf _
'                       & "" & Trim(!CUSTCITY) & ", " _
'                       & " " & Trim(!CUSTSTATE) & "  " & Trim(!CUSTZIP)
'            txtTax = "" & Trim(!CUTAXEXEMPT)
'            txtCnt = "" & Trim(!CUCCONTACT)
'            txtCph = "" & Trim(!CUCPHONE)
'            txtExt = !CUCEXT
'            txtCcInt = "" & Trim(!CUCINTPHONE)
'            txtCfax = "" & Trim(!CUFAX)
'            txtCintFax = "" & Trim(!CUINTFAX)
'            cmbDiv = "" & Trim(!CUDIVISION)
'            cmbReg = "" & Trim(!CUREGION)
'            If bGoodSO Then
'               On Error Resume Next
'               With RdoRso
'                  '.Edit
'                  !SOSTNAME = "" & txtStNme
'                  !SOSTADR = "" & txtStAdr
'                  !SOTAXEXEMPT = "" & txtTax
'                  !SOCCONTACT = "" & Trim(txtCnt)
'                  !SOCPHONE = "" & Trim(txtCph)
'                  !SOCEXT = Val(txtExt)
'                  !SOCINTPHONE = "" & Trim(txtCcInt)
'                  !SOCFAX = "" & Trim(txtCfax)
'                  !SOCINTFAX = "" & Trim(txtCintFax)
'                  !SODIVISION = "" & Trim(cmbDiv)
'                  !SOREGION = "" & Trim(cmbReg)
'                  RdoRso.Update
'               End With
'            End If
'            sOldCustomer = Compress(cmbCst)
'         End If
'         ClearResultSet RdoCbt
'      End With
'   Else
'      txtBto = ""
'   End If
'   Set RdoCbt = Nothing
'   Exit Sub
'
'DiaErr1:
'   sProcName = "getbillto"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub


Private Sub GetBillTo()
   Dim RdoCbt As ADODB.Recordset
   Dim sCust As String
   On Error GoTo DiaErr1
   sCust = Compress(cmbCst)
   sSql = "SELECT CUBTNAME,CUBTADR,CUBCITY,CUBSTATE,CUBZIP," _
          & "CUSTNAME,CUSTADR,CUSTCITY,CUSTSTATE,CUSTZIP,CUTAXEXEMPT," _
          & "CUCCONTACT,CUCPHONE,CUCEXT,CUCINTPHONE,CUINTFAX,CUFAX," _
          & "CUDIVISION,CUREGION,CUCUTOFF FROM CustTable WHERE " _
          & "CUREF='" & sCust & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCbt, ES_FORWARD)
   If bSqlRows Then
      With RdoCbt
         If "" & !CUCUTOFF Then
            txtNme = "*** This Customer's Credit Is On Hold ***"
            txtNme.ForeColor = ES_RED
            bCutOff = 1
            cmdItm.Enabled = False
         Else
            txtNme.ForeColor = vbBlack
            bCutOff = 0
         End If
         txtBto = "" & Trim(!CUBTNAME) & vbCrLf _
                  & "" & Trim(!CUBTADR) & vbCrLf _
                  & "" & Trim(!CUBCITY) & ", " _
                  & " " & Trim(!CUBSTATE) & "  " & Trim(!CUBZIP)
         ClearResultSet RdoCbt
      End With
   Else
      txtBto = ""
   End If
   Set RdoCbt = Nothing
   Exit Sub

DiaErr1:
   sProcName = "getbillto"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub


'check for null dates

Private Sub CheckPriceBookDates()
   Dim vDate1 As Variant
   Dim vDate2 As Variant
   On Error Resume Next
   vDate1 = Format(ES_SYSDATE, "mm/dd/yyyy")
   vDate2 = Format(ES_SYSDATE + 365, "mm/dd/yyyy")
   sSql = "UPDATE PbhdTable SET PBHSTARTDATE='" & vDate1 & "'," _
          & "PBHENDDATE='" & vDate2 & "' WHERE (PBHSTARTDATE IS NULL " _
          & "OR PBHENDDATE IS NULL)"
      
   
   ' Aduit the column
'   SaveSOHistory "", "UPDATE", "PBHSTARTDATE", CStr(vDate1), ""
'   SaveSOHistory "", "UPDATE", "PBHENDDATE", CStr(vDate2), ""
   
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
End Sub

Private Sub UpdateSalesOrder()
   On Error Resume Next
   With RdoRso
      
      ' Aduit the column
'      SaveSOHistory "", "UPDATE", "SOTYPE", cmbPre, "" & !SOTYPE
'      SaveSOHistory "", "UPDATE", "SODATE", txtBok, "" & !SODATE
'      SaveSOHistory "", "UPDATE", "SODELDATE", txtDdt, "" & !SODELDATE
'      SaveSOHistory "", "UPDATE", "SOSHIPDATE", txtShp, "" & !SOSHIPDATE
'      SaveSOHistory "", "UPDATE", "SODIVISION", cmbDiv, "" & !SODIVISION
'      SaveSOHistory "", "UPDATE", "SOREGION", cmbReg, "" & !SOREGION
'      SaveSOHistory "", "UPDATE", "SOSALESMAN", cmbSlp, "" & !SOSALESMAN
'      SaveSOHistory "", "UPDATE", "SOTERMS", cmbTrm, "" & !SOSTERMS
'      SaveSOHistory "", "UPDATE", "SOREMARKS", RTrim(txtCmt), "" & !SOREMARKS
   
      '.Edit
      !SOTYPE = cmbPre
      !SODATE = txtBok
      !SODELDATE = txtDdt
      !SOSHIPDATE = txtShp
      !SODIVISION = cmbDiv
      !SOREGION = cmbReg
      !SOSALESMAN = "" & cmbSlp
      !SOSTERMS = cmbTrm
      !SOREMARKS = RTrim(txtCmt)
      .Update
      .Close
   
   End With
   bDataHasChanged = False
   
End Sub

'Refresh every 3rd time

Public Sub RefreshSoCursor(Optional ResetCounter As Byte)
   Static bByte As Byte
   On Error Resume Next
   If ResetCounter = 1 Then bByte = 0
   If bByte = 0 Then
      'rdoQry.RowsetSize = 1
      'rdoQry(0) = Val(cmbSon)
      cmdObj.parameters(0).Value = Val(cmbSon)
      bSqlRows = clsADOCon.GetQuerySet(RdoRso, cmdObj, ES_KEYSET, True)
      
      'bSqlRows = GetQuerySet(RdoRso, rdoQry, ES_KEYSET)
   End If
   bByte = bByte + 1
   If bByte > 2 Then bByte = 0
   
End Sub

'Private Function SaveSOHistory(strTransID As String, strType As String, strSOColName As String, _
'               strNewValue As String, strOldValue As String)
'   Dim strTableName As String
'   Dim strITSO As String
'   Dim strITRev As String
'   Dim strITSONum As String
'   Dim strUser As String
'
'
'   strTableName = "AuditSoTable"
'   strITSO = cmbSon
'   strITRev = ""
'   strITSONum = ""
'
'   If (TableExists(strTableName)) Then
'      If (strTransID = "") Then
'         strTransID = GetAuditTransID(strTableName)
'      End If
'
'      If (strOldValue <> strNewValue) Then
'         AuditSO strTableName, strTransID, strITSO, strITSONum, strITRev, _
'                  sInitials, strType, strSOColName, strOldValue, strNewValue
'      End If
'   End If
'End Function
'
