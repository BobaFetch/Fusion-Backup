VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PurcPRe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendors"
   ClientHeight    =   5685
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   7395
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.Frame tabFrame 
      Height          =   3972
      Index           =   2
      Left            =   8040
      TabIndex        =   71
      Top             =   1560
      Width           =   7092
      Begin VB.ComboBox txtPcntr 
         Height          =   288
         Left            =   1320
         TabIndex        =   32
         Tag             =   "3"
         ToolTipText     =   "Select Country From The List"
         Top             =   1776
         Width           =   2055
      End
      Begin VB.ComboBox cmbAct 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1320
         TabIndex        =   33
         Tag             =   "3"
         Top             =   2196
         Width           =   1935
      End
      Begin VB.ComboBox cmbPste 
         Height          =   288
         Left            =   4080
         Sorted          =   -1  'True
         TabIndex        =   30
         Tag             =   "3"
         ToolTipText     =   "State Code"
         Top             =   1440
         Width           =   660
      End
      Begin VB.TextBox txtPdue 
         Height          =   285
         Left            =   2880
         TabIndex        =   39
         Tag             =   "1"
         Top             =   3636
         Width           =   495
      End
      Begin VB.TextBox txtPdsc 
         Height          =   285
         Index           =   1
         Left            =   5040
         TabIndex        =   38
         Tag             =   "1"
         Top             =   3276
         Width           =   495
      End
      Begin VB.TextBox txtPdte 
         Height          =   285
         Left            =   3600
         TabIndex        =   37
         Tag             =   "1"
         Top             =   3276
         Width           =   495
      End
      Begin VB.TextBox txtPday 
         Height          =   285
         Left            =   5040
         TabIndex        =   36
         Tag             =   "1"
         Top             =   2916
         Width           =   495
      End
      Begin VB.TextBox txtPdsc 
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   35
         Tag             =   "1"
         Top             =   2916
         Width           =   495
      End
      Begin VB.TextBox txtPnet 
         Height          =   285
         Left            =   1080
         TabIndex        =   34
         Tag             =   "1"
         Top             =   2916
         Width           =   495
      End
      Begin VB.TextBox txtPnme 
         Height          =   285
         Left            =   1320
         TabIndex        =   27
         Tag             =   "2"
         Top             =   396
         Width           =   3475
      End
      Begin VB.TextBox txtPzip 
         Height          =   285
         Left            =   5280
         TabIndex        =   31
         Tag             =   "3"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtPcty 
         Height          =   285
         Left            =   1320
         TabIndex        =   29
         Tag             =   "2"
         Top             =   1440
         Width           =   2085
      End
      Begin VB.TextBox txtPadr 
         Height          =   615
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   28
         Tag             =   "9"
         Top             =   756
         Width           =   3475
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         Height          =   288
         Index           =   36
         Left            =   120
         TabIndex        =   89
         Top             =   1776
         Width           =   912
      End
      Begin VB.Label lblDsc 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   "
         Height          =   288
         Left            =   3600
         TabIndex        =   88
         Top             =   2196
         Width           =   3132
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Account"
         Height          =   288
         Index           =   34
         Left            =   120
         TabIndex        =   87
         Top             =   2196
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "% "
         Height          =   288
         Index           =   31
         Left            =   5640
         TabIndex        =   86
         Top             =   3276
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount If Paid By The"
         Height          =   288
         Index           =   30
         Left            =   1080
         TabIndex        =   85
         Top             =   3636
         Width           =   1872
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Receive a"
         Height          =   288
         Index           =   29
         Left            =   4200
         TabIndex        =   84
         Top             =   3276
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Prox Terms:  Invoices Dated On Or Before The "
         Height          =   288
         Index           =   28
         Left            =   120
         TabIndex        =   83
         Top             =   3276
         Width           =   3552
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Days Or Less"
         Height          =   288
         Index           =   27
         Left            =   5640
         TabIndex        =   82
         Top             =   2916
         Width           =   1152
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "% For"
         Height          =   288
         Index           =   26
         Left            =   4560
         TabIndex        =   81
         Top             =   2916
         Width           =   552
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Days.  Quick Payment Disc Of"
         Height          =   288
         Index           =   25
         Left            =   1680
         TabIndex        =   80
         Top             =   2916
         Width           =   2352
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Terms: Net "
         Height          =   288
         Index           =   24
         Left            =   120
         TabIndex        =   79
         Top             =   2916
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Terms (Either Terms or Prox Terms):"
         Height          =   288
         Index           =   23
         Left            =   120
         TabIndex        =   78
         Top             =   2556
         Width           =   2832
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   288
         Index           =   21
         Left            =   120
         TabIndex        =   77
         Top             =   396
         Width           =   1188
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         Height          =   288
         Index           =   20
         Left            =   4920
         TabIndex        =   76
         Top             =   1476
         Width           =   552
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         Height          =   288
         Index           =   19
         Left            =   3600
         TabIndex        =   75
         Top             =   1440
         Width           =   672
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   288
         Index           =   18
         Left            =   120
         TabIndex        =   74
         Top             =   756
         Width           =   1152
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   288
         Index           =   17
         Left            =   120
         TabIndex        =   73
         Top             =   1440
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Check Information:"
         Height          =   288
         Index           =   16
         Left            =   120
         TabIndex        =   72
         Top             =   120
         Width           =   1632
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   3852
      Index           =   1
      Left            =   7680
      TabIndex        =   61
      Top             =   1560
      Width           =   7092
      Begin VB.TextBox txtArEmail 
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         Tag             =   "2"
         ToolTipText     =   "Double Click Here To Send E-Mail (Requires An Entry)"
         Top             =   2040
         Width           =   5220
      End
      Begin VB.CheckBox optEom 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   5280
         TabIndex        =   24
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox optAtt 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   3360
         TabIndex        =   23
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtInts 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox opt1099 
         Caption         =   "___"
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   1320
         TabIndex        =   22
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtScmt 
         Height          =   1095
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Tag             =   "9"
         ToolTipText     =   "Comments (2048 Chars Max)"
         Top             =   2520
         Width           =   4170
      End
      Begin VB.TextBox txtStid 
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Tag             =   "3"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtSfob 
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Tag             =   "3"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtSext 
         Height          =   285
         Left            =   3480
         TabIndex        =   19
         Tag             =   "1"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtScnt 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Tag             =   "2"
         Top             =   240
         Width           =   2535
      End
      Begin MSMask.MaskEdBox txtSphn 
         Height          =   288
         Left            =   1704
         TabIndex        =   18
         Top             =   600
         Width           =   1308
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "AR EMail"
         Height          =   288
         Index           =   39
         Left            =   120
         TabIndex        =   91
         Top             =   2040
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "EOM"
         Height          =   288
         Index           =   38
         Left            =   4320
         TabIndex        =   70
         Top             =   1680
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Attorney"
         Height          =   288
         Index           =   37
         Left            =   2280
         TabIndex        =   69
         Top             =   1680
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Requires 1099"
         Height          =   288
         Index           =   35
         Left            =   120
         TabIndex        =   68
         Top             =   1680
         Width           =   1152
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Id No"
         Height          =   288
         Index           =   14
         Left            =   120
         TabIndex        =   67
         Top             =   1320
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "FOB"
         Height          =   288
         Index           =   13
         Left            =   120
         TabIndex        =   66
         Top             =   960
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   288
         Index           =   12
         Left            =   120
         TabIndex        =   65
         Top             =   600
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext"
         Height          =   288
         Index           =   11
         Left            =   3120
         TabIndex        =   64
         Top             =   600
         Width           =   432
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         Height          =   288
         Index           =   10
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         Height          =   288
         Index           =   15
         Left            =   120
         TabIndex        =   62
         Top             =   2520
         Width           =   912
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   3972
      Index           =   0
      Left            =   40
      TabIndex        =   48
      Top             =   1560
      Width           =   7176
      Begin VB.ComboBox txtCntr 
         Height          =   288
         Left            =   1320
         TabIndex        =   7
         Tag             =   "3"
         ToolTipText     =   "Select Country From The List"
         Top             =   1476
         Width           =   2055
      End
      Begin VB.TextBox txtIntf 
         Height          =   285
         Left            =   4680
         TabIndex        =   11
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   1920
         Width           =   475
      End
      Begin VB.TextBox txtIntp 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   1920
         Width           =   475
      End
      Begin VB.ComboBox cmbByr 
         Height          =   288
         Left            =   1320
         TabIndex        =   15
         Tag             =   "3"
         Top             =   3000
         Width           =   2775
      End
      Begin VB.ComboBox cmbSte 
         Height          =   288
         Left            =   4680
         Sorted          =   -1  'True
         TabIndex        =   5
         Tag             =   "3"
         ToolTipText     =   "State Code"
         Top             =   1164
         Width           =   660
      End
      Begin VB.TextBox txtCnt 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Tag             =   "2"
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtExt 
         Height          =   285
         Left            =   3480
         TabIndex        =   10
         Tag             =   "1"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtEml 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Tag             =   "2"
         ToolTipText     =   "Double Click Here To Send E-Mail (Requires An Entry)"
         Top             =   2640
         Width           =   5340
      End
      Begin VB.TextBox txtAdr 
         Height          =   855
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   3
         Tag             =   "9"
         Top             =   240
         Width           =   3475
      End
      Begin VB.TextBox txtCty 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Tag             =   "2"
         Top             =   1164
         Width           =   2085
      End
      Begin VB.TextBox txtZip 
         Height          =   285
         Left            =   5760
         TabIndex        =   6
         Tag             =   "3"
         Top             =   1164
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtPhn 
         Height          =   288
         Left            =   1800
         TabIndex        =   9
         Top             =   1920
         Width           =   1308
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFax 
         Height          =   288
         Left            =   5160
         TabIndex        =   12
         Top             =   1920
         Width           =   1308
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         Height          =   288
         Index           =   64
         Left            =   120
         TabIndex        =   60
         Top             =   1476
         Width           =   912
      End
      Begin VB.Label lblByr 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   1320
         TabIndex        =   59
         Top             =   3360
         Width           =   3612
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Buyer"
         Height          =   288
         Index           =   33
         Left            =   120
         TabIndex        =   58
         Top             =   3000
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         Height          =   288
         Index           =   9
         Left            =   120
         TabIndex        =   57
         Top             =   2280
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext"
         Height          =   288
         Index           =   8
         Left            =   3120
         TabIndex        =   56
         Top             =   1920
         Width           =   432
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
         Height          =   288
         Index           =   5
         Left            =   120
         TabIndex        =   55
         Top             =   2640
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   288
         Index           =   2
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   1032
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   288
         Index           =   6
         Left            =   120
         TabIndex        =   53
         Top             =   1920
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   288
         Index           =   7
         Left            =   4200
         TabIndex        =   52
         Top             =   1920
         Width           =   708
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   288
         Index           =   3
         Left            =   120
         TabIndex        =   51
         Top             =   1164
         Width           =   912
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         Height          =   288
         Index           =   4
         Left            =   4200
         TabIndex        =   50
         Top             =   1164
         Width           =   672
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         Height          =   288
         Index           =   22
         Left            =   5400
         TabIndex        =   49
         Top             =   1200
         Width           =   552
      End
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   4452
      Left            =   24
      TabIndex        =   47
      Top             =   1200
      Width           =   7308
      _ExtentX        =   12885
      _ExtentY        =   7858
      TabWidthStyle   =   2
      TabFixedWidth   =   2293
      TabFixedHeight  =   473
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Service/Other "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Payables"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame z2 
      Height          =   70
      Index           =   0
      Left            =   120
      TabIndex        =   46
      Top             =   1080
      Width           =   7035
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRe03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbTyp 
      Height          =   288
      Left            =   6480
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "(2) Character Vendor Type.  Searches Previous Entries"
      Top             =   720
      Width           =   732
   End
   Begin VB.Timer Timer1 
      Index           =   0
      Interval        =   6000
      Left            =   5280
      Top             =   120
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   288
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Vendor or Enter a New Vendor (10 Char Max)"
      Top             =   360
      Width           =   1555
   End
   Begin VB.TextBox txtNme 
      Height          =   285
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "2"
      Top             =   720
      Width           =   3475
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6360
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   5520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5685
      FormDesignWidth =   7395
   End
   Begin VB.Label Label1 
      Caption         =   "More >>>"
      Height          =   252
      Left            =   6360
      TabIndex        =   90
      Top             =   480
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Type"
      Height          =   285
      Index           =   32
      Left            =   5100
      TabIndex        =   44
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label lblNum 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3120
      TabIndex        =   43
      Top             =   360
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   42
      Top             =   360
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Name"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   41
      Top             =   720
      Width           =   1425
   End
End
Attribute VB_Name = "PurcPRe03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/8/03 International Calling Codes
'2/6/04  Email
'5/24/04 Countries
'5/24/04 EOM
'8/8/06 Replaced SSTab32
'1/8/07 VEAREMAIL 7.2.0
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter
Dim RdoVdr As ADODB.Recordset

Dim bGoodVendor As Byte
Dim bNewVendor As Byte
Dim bOnLoad As Byte

Dim iTimer As Integer

Dim dToday As Date
Dim sVendRef As String
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

Private Sub FillCountries()
   On Error GoTo DiaErr1
   sSql = "Qry_FillCountries"
   LoadComboBox txtCntr
   LoadComboBox txtPcntr
   If txtCntr.ListCount > 0 Then
      txtCntr = "USA"
      txtPcntr = "USA"
   Else
      txtCntr.AddItem "USA"
      txtCntr = "USA"
      txtPcntr.AddItem "USA"
      txtPcntr = "USA"
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcountries"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillAccounts()
   On Error GoTo DiaErr1
   sSql = "Qry_FillLowAccounts"
   LoadComboBox cmbAct
   If cmbAct.ListCount > 0 Then
      cmbAct = cmbAct.List(0)
      cmbAct.Enabled = True
      FindAccount Me
   Else
      cmbAct = "No Accounts."
   End If
   sSql = "SELECT DISTINCT VETYPE FROM VndrTable WHERE VETYPE<>''"
   LoadComboBox cmbTyp, -1
   Exit Sub
   
DiaErr1:
   sProcName = "fillaccou"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   b = AutoFormatChange(Me, txtChange())
   
End Sub


Private Sub cmbAct_Click()
   FindAccount Me
   
End Sub


Private Sub cmbAct_DropDown()
   bDataHasChanged = True
   
End Sub



Private Sub cmbAct_LostFocus()
   cmbAct = CheckLen(cmbAct, 12)
   If cmbAct = "" Then
      lblDsc = ""
   Else
      FindAccount Me
   End If
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEACCOUNT = Compress(cmbAct)
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub cmbByr_Click()
   If Len(Trim(cmbByr)) > 0 Then GetCurrentBuyer cmbByr
   
End Sub

Private Sub CmbByr_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub cmbByr_LostFocus()
   cmbByr = CheckLen(cmbByr, 20)
   If bGoodVendor Then
      GetCurrentBuyer cmbByr
      On Error Resume Next
      RdoVdr!VEBUYER = "" & Compress(cmbByr)
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub cmbPste_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub cmbPste_LostFocus()
   cmbPste = CheckLen(cmbPste, 2)
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VECSTATE = cmbPste
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub cmbSte_Change()
   If Len(cmbSte) > 2 Then cmbSte = Left(cmbSte, 2)
   
End Sub

Private Sub cmbSte_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub cmbSte_LostFocus()
   cmbSte = CheckLen(cmbSte, 2)
   If bGoodVendor Then
      On Error Resume Next
      If bNewVendor Then
         cmbPste = cmbSte
         RdoVdr!VEBSTATE = cmbSte
         RdoVdr!VECSTATE = cmbSte
         RdoVdr.Update
      Else
         RdoVdr!VEBSTATE = cmbSte
         RdoVdr.Update
      End If
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub cmbTyp_Change()
   If Len(cmbTyp) > 2 Then cmbTyp = Left(cmbTyp, 2)
   
End Sub

Private Sub cmbTyp_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub cmbTyp_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbTyp = CheckLen(cmbTyp, 2)
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VETYPE = "" & cmbTyp
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   For iList = 0 To cmbTyp.ListCount - 1
      If cmbTyp = cmbTyp.List(iList) Then bByte = 1
   Next
   If bByte = 0 And cmbTyp <> "" Then cmbTyp.AddItem cmbTyp
   
End Sub


Private Sub cmbVnd_Click()
   bGoodVendor = GetVendor(0)
   
End Sub


Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   bGoodVendor = GetVendor(1)
   If Len(cmbVnd) > 0 Then If bGoodVendor = 0 Then AddVendor
   If bGoodVendor Then ManageBoxes True
   
End Sub


Private Sub cmdCan_Click()
   Timer1(0).Enabled = False
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4303
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bNewVendor = 0
      bOnLoad = 0
      FillAccounts
      FillBuyers
      FillCountries
      FillStates Me
      sState = cmbSte
      FillVendors
      If cUR.CurrentVendor <> "" Then cmbVnd = cUR.CurrentVendor
      bDataHasChanged = False
      bGoodVendor = GetVendor(0)
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim cNewSize As Currency
   MouseCursor 13
   FormLoad Me, ES_LIST, ES_RESIZE
   If bResize = 1 Then
      If lScreenWidth > 9999 Then '640X480
         cNewSize = 1.05
         Height = Height * cNewSize
         Width = Width * cNewSize
      End If
   End If
   tabFrame(0).BorderStyle = 0
   tabFrame(1).BorderStyle = 0
   tabFrame(2).BorderStyle = 0
   tabFrame(0).Visible = True
   tabFrame(1).Visible = False
   tabFrame(2).Visible = False
   tabFrame(0).Left = 40
   tabFrame(1).Left = 40
   tabFrame(2).Left = 40
   
   FormatControls
   
   sSql = "SELECT TOP 1 * FROM VndrTable " _
          & "WHERE (VEREF = ? AND VEREF<>'NONE') "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Size = 10
   AdoParameter.Type = adChar
   AdoQry.Parameters.Append AdoParameter
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   'tab1.Tab = 0
   'RdoQry.MaxRows = 1
   dToday = Format(ES_SYSDATE, "mm/dd/yy")
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bDataHasChanged And bGoodVendor = 1 Then UpdateVendorSet True
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If bGoodVendor Then cUR.CurrentVendor = cmbVnd
   SaveCurrentSelections
   FormUnload
   Set AdoParameter = Nothing
   Set RdoVdr = Nothing
   Set AdoQry = Nothing
   Set PurcPRe03a = Nothing
   
End Sub






Private Function GetVendor(Optional bUpdate As Byte) As Byte
   If bDataHasChanged And bGoodVendor = 1 Then UpdateVendorSet
   sVendRef = Compress(cmbVnd)
   On Error GoTo DiaErr1
   ClearBoxes
   bGoodVendor = 0
   AdoQry.Parameters(0).value = sVendRef
   bSqlRows = clsADOCon.GetQuerySet(RdoVdr, AdoQry, ES_KEYSET, True, 1)
   If bSqlRows Then
      With RdoVdr
         On Error Resume Next
         cmbVnd = "" & Trim(!VENICKNAME)
         lblNum = 0 + Format(!VENUMBER, "#0")
         txtNme = "" & Trim(!VEBNAME)
         cmbTyp = "" & Trim(!VETYPE)
         txtAdr = "" & Trim(!VEBADR)
         txtCty = "" & Trim(!VEBCITY)
         If Len(Trim(!VEBADR)) = 0 And Len(Trim(!VEBCITY)) = 0 Then _
                bNewVendor = 1 Else bNewVendor = 0
         If Len(Trim(!VEBSTATE)) Then
            cmbSte = "" & Trim(!VEBSTATE)
         Else
            cmbSte = sState
         End If
         txtZip = "" & Trim(!VEBZIP)
         '5/14/04
         txtCntr = "" & Trim(!VEBCOUNTRY)
         txtPcntr = "" & Trim(!VECCOUNTRY)
         optEom = !VEBEOM
         
         If Len("" & Trim(!VEBPHONE)) Then txtPhn = "" & Trim(!VEBPHONE)
         txtCnt = "" & Trim(!VEBCONTACT)
         If Len(Trim(!VEFAX)) Then txtFax = "" & Trim(!VEFAX)
         If !VEBEXT > 0 Then txtExt = !VEBEXT
         txtEml = "" & Trim(!VEEMAIL)
         cmbByr = "" & Trim(!VEBUYER)
         txtScnt = "" & Trim(!VESCONTACT)
         If Len(Trim("" & !VESPHONE)) Then txtSphn = "" & Trim(!VESPHONE)
         txtSfob = "" & Trim(!VEFOB)
         txtStid = "" & Trim(!VETAXIDNO)
         opt1099 = 0 + !VE1099
         txtScmt = "" & Trim(!VECOMT)
         If !VESEXT > 0 Then txtSext = !VESEXT
         txtPnme = "" & Trim(!VECNAME)
         txtPadr = "" & Trim(!VECADR)
         txtPcty = "" & Trim(!VECCITY)
         If Len(Trim(!VECSTATE)) = 0 Then
            cmbPste = cmbSte
         Else
            cmbPste = "" & Trim(!VECSTATE)
         End If
         txtPzip = "" & Trim(!VECZIP)
         txtPdsc(0) = 0 + Format(!VEDISCOUNT, "#0.0")
         txtPdsc(1) = 0 + Format(!VEDISCOUNT, "#0.0")
         txtPnet = 0 + Format(!VENETDAYS, "#0")
         txtPday = 0 + Format(!VEDDAYS, "#0")
         txtPdte = 0 + Format(!VEPROXDT, "#0")
         txtPdue = 0 + Format(!VEPROXDUE, "#0")
         '10/8/03 International phones
         txtIntp = "" & Trim(!VEINTPHONE)
         txtIntf = "" & Trim(!VEINTFAX)
         txtInts = "" & Trim(!VESINTPHONE)
         '11/20/03
         cmbAct = "" & Trim(!VEACCOUNT)
         If Trim(cmbAct) > "" Then
            FindAccount Me
         Else
            lblDsc = ""
         End If
         '1/29/04
         optAtt.value = !VEATTORNEY
         '1/8/07
         txtArEmail = "" & Trim(!VEAREMAIL)
      End With
      If Len(Trim(cmbByr)) > 0 Then GetCurrentBuyer cmbByr
      ManageBoxes True
      GetVendor = 1
   Else
      ManageBoxes False
      GetVendor = 0
   End If
   bDataHasChanged = False
   iTimer = 0
   Exit Function
   
DiaErr1:
   sProcName = "getvendor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub ClearBoxes()
   Dim iList As Integer
   On Error Resume Next
   For iList = 0 To Controls.Count - 1
      If TypeOf Controls(iList) Is TextBox Then
         Controls(iList).Text = " "
      End If
   Next
   opt1099.value = vbUnchecked
   txtPhn.Mask = ""
   txtPhn = ""
   txtPhn.Mask = "###-###-####"
   txtFax.Mask = ""
   txtFax = ""
   txtFax.Mask = "###-###-####"
   txtSphn.Mask = ""
   txtSphn = ""
   txtSphn.Mask = "###-###-####"
   
End Sub

Private Sub lblByr_Change()
   If Left(lblByr, 8) = "*** Buye" Then
      lblByr.ForeColor = ES_RED
   Else
      lblByr.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub opt1099_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub opt1099_LostFocus()
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VE1099 = opt1099.value
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub optAtt_Click()
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEATTORNEY = optAtt.value
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub optAtt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optEom_Click()
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEBEOM = optEom.value
      RdoVdr.Update
   End If
   
End Sub

Private Sub optEom_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub tab1_Click()
   On Error Resume Next
   ClearCombos
   If tab1.SelectedItem.Index = 1 Then
      tabFrame(0).Visible = True
      tabFrame(1).Visible = False
      tabFrame(2).Visible = False
      txtAdr.SetFocus
   ElseIf tab1.SelectedItem.Index = 2 Then
      tabFrame(1).Visible = True
      tabFrame(0).Visible = False
      tabFrame(2).Visible = False
      txtScnt.SetFocus
      
   Else
      tabFrame(2).Visible = True
      tabFrame(0).Visible = False
      tabFrame(1).Visible = False
      txtPnme.SetFocus
   End If
   
End Sub

Private Sub Tab1_GotFocus()
   Dim bReponse As Byte
   Dim sMsg As String
   iTimer = 0
   If iTimer > 300 Then
      '6000/300 sec 30 min
      On Error Resume Next
      'RdoVdr.Close
      Set RdoVdr = Nothing
      sMsg = Format(Now, "mm/dd/yy hh:mm AM/PM") _
             & " This Normal Time Out." & vbCr _
             & "All Current Data Has Been Saved."
      MsgBox sMsg, vbInformation, Caption
      Timer1(0).Enabled = False
      Unload Me
   End If
   
End Sub



Private Sub Timer1_Timer(Index As Integer)
   Dim bReponse As Byte
   Dim sMsg As String
   iTimer = iTimer + 1
   If iTimer > 300 Then
      '6000/300 sec 30 min
      Set RdoVdr = Nothing
      sMsg = Format(Now, "mm/dd/yy hh:mm AM/PM") _
             & " This Normal Time Out." & vbCr _
             & "All Current Data Has Been Saved."
      MsgBox sMsg, vbInformation, Caption
      Timer1(0).Enabled = False
      Unload Me
   End If
   
End Sub

Private Sub txtAdr_LostFocus()
   txtAdr = CheckLen(txtAdr, 160)
   If bGoodVendor Then
      On Error Resume Next
      If bNewVendor Then
         txtPadr = txtAdr
         RdoVdr!VEBADR = txtAdr
         RdoVdr!VECADR = txtAdr
         RdoVdr!VEDATEREV = dToday
         RdoVdr.Update
      Else
         RdoVdr!VEBADR = txtAdr
         RdoVdr!VEDATEREV = dToday
         RdoVdr.Update
      End If
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtArEmail_DblClick()
   If Trim(txtArEmail) <> "" Then SendEMail Trim(txtArEmail)
   
End Sub

Private Sub txtArEmail_LostFocus()
   txtArEmail = CheckLen(txtArEmail, 60)
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEAREMAIL = txtArEmail
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtCnt_LostFocus()
   txtCnt = CheckLen(txtCnt, 20)
   txtCnt = StrCase(txtCnt)
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEBCONTACT = txtCnt
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtCntr_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub txtCntr_LostFocus()
   txtCntr = CheckLen(txtCntr, 20)
   If bGoodVendor Then
      On Error Resume Next
      If bNewVendor Then
         txtPcntr = txtCntr
         RdoVdr!VEBCOUNTRY = txtCntr
         RdoVdr!VECCOUNTRY = txtCntr
         RdoVdr.Update
      Else
         RdoVdr!VEBCOUNTRY = txtCntr
         RdoVdr.Update
      End If
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtCty_LostFocus()
   txtCty = CheckLen(txtCty, 18)
   txtCty = StrCase(txtCty)
   iTimer = 0
   If bGoodVendor Then
      On Error Resume Next
      If bNewVendor Then
         txtPcty = txtCty
         RdoVdr!VEBCITY = txtCty
         RdoVdr!VECCITY = txtCty
         RdoVdr.Update
      Else
         RdoVdr!VEBCITY = txtCty
         RdoVdr.Update
      End If
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtEml_DblClick()
   If Trim(txtEml) <> "" Then SendEMail Trim(txtEml)
   
End Sub


Private Sub txtEml_LostFocus()
   txtEml = CheckLen(txtEml, 60)
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEEMAIL = txtEml
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtExt_LostFocus()
   txtExt = CheckLen(txtExt, 4)
   If Len(txtExt) > 0 Then txtExt = Format(Abs(Val(txtExt)), "###0")
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEBEXT = Val(txtExt)
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtFax_LostFocus()
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEFAX = "" & txtFax
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtIntf_Change()
   If Len(txtIntf) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtIntf_LostFocus()
   txtIntf = CheckLen(txtIntf, 4)
   txtIntf = Format(Abs(Val(txtIntf)), "###")
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEINTFAX = txtIntf
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtIntp_Change()
   If Len(txtIntp) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtIntp_LostFocus()
   txtIntp = CheckLen(txtIntp, 4)
   txtIntp = Format(Abs(Val(txtIntp)), "###")
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEINTPHONE = txtIntp
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtInts_Change()
   If Len(txtInts) > 3 Then SendKeys "{tab}"
   
End Sub

Private Sub txtInts_LostFocus()
   txtInts = CheckLen(txtInts, 4)
   txtInts = Format(Abs(Val(txtInts)), "###")
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VESINTPHONE = txtInts
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtNme_LostFocus()
   txtNme = CheckLen(txtNme, 40)
   txtNme = StrCase(txtNme)
   If bGoodVendor Then
      On Error Resume Next
      If bNewVendor Then
         txtPnme = txtNme
         RdoVdr!VEBNAME = txtNme
         RdoVdr!VECNAME = txtNme
         RdoVdr.Update
      Else
         RdoVdr!VEBNAME = txtNme
         RdoVdr.Update
      End If
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtPadr_LostFocus()
   txtPadr = CheckLen(txtPadr, 160)
   iTimer = 0
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VECADR = txtPadr
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtPcntr_DropDown()
   bDataHasChanged = True
   
End Sub

Private Sub txtPcntr_LostFocus()
   txtPcntr = CheckLen(txtPcntr, 20)
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VECCOUNTRY = txtPcntr
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   bNewVendor = 0
   
End Sub


Private Sub txtPcty_LostFocus()
   txtPcty = CheckLen(txtPcty, 18)
   txtPcty = StrCase(txtPcty)
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VECCITY = txtPcty
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtPday_LostFocus()
   txtPday = CheckLen(txtPday, 3)
   If Len(txtPday) > 0 Then txtPday = Format(Abs(Val(txtPday)), "##0")
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEDDAYS = Val(txtPday)
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtPdsc_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtPdsc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtPdsc_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtPdsc_LostFocus(Index As Integer)
   txtPdsc(Index) = CheckLen(txtPdsc(Index), 4)
   If Len(txtPdsc(Index)) > 0 Then txtPdsc(Index) = Format(Abs(Val(txtPdsc(Index))), "#0.0")
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEDISCOUNT = Val(txtPdsc(Index))
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtPdte_LostFocus()
   txtPdte = CheckLen(txtPdte, 2)
   If Len(txtPdte) > 0 Then txtPdte = Format(Abs(Val(txtPdte)), "#0")
   If Val(txtPdte) > 30 Then txtPdte = "30"
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEPROXDT = Val(txtPdte)
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtPdue_LostFocus()
   txtPdue = CheckLen(txtPdue, 2)
   If Len(txtPdue) > 0 Then txtPdue = Format(Abs(Val(txtPdue)), "#0")
   If Val(txtPdue) > 30 Then txtPdue = "30"
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEPROXDUE = Val(txtPdue)
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtPhn_LostFocus()
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEBPHONE = txtPhn
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtPnet_LostFocus()
   txtPnet = CheckLen(txtPnet, 3)
   If Len(txtPnet) > 0 Then txtPnet = Format(Abs(Val(txtPnet)), "##0")
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VENETDAYS = Val(txtPnet)
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtPnme_LostFocus()
   txtPnme = CheckLen(txtPnme, 40)
   txtPnme = StrCase(txtPnme)
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VECNAME = txtPnme
      RdoVdr!VEDATEREV = dToday
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtPzip_LostFocus()
   txtPzip = CheckLen(txtPzip, 10)
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VECZIP = txtPzip
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtScmt_LostFocus()
   txtScmt = CheckLen(txtScmt, 2048)
   txtScmt = StrCase(txtScmt, ES_FIRSTWORD)
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VECOMT = txtScmt
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtScnt_LostFocus()
   txtScnt = CheckLen(txtScnt, 20)
   txtScnt = StrCase(txtScnt)
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VESCONTACT = txtScnt
      RdoVdr!VEDATEREV = dToday
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtSext_LostFocus()
   txtSext = CheckLen(txtSext, 4)
   If Len(txtSext) > 0 Then txtSext = Format(Abs(Val(txtSext)), "###0")
   txtScmt = CheckLen(txtScmt, 255)
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VESEXT = Val(txtSext)
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtSfob_LostFocus()
   txtSfob = CheckLen(txtSfob, 12)
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VEFOB = txtSfob
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtSphn_LostFocus()
   iTimer = 0
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VESPHONE = txtSphn
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtStid_LostFocus()
   txtStid = CheckLen(txtStid, 20)
   If bGoodVendor Then
      On Error Resume Next
      RdoVdr!VETAXIDNO = txtStid
      RdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtZip_LostFocus()
   txtZip = CheckLen(txtZip, 10)
   If bGoodVendor Then
      On Error Resume Next
      If bNewVendor Then
         txtPzip = txtZip
         RdoVdr!VEBZIP = txtZip
         RdoVdr!VECZIP = txtZip
         RdoVdr.Update
      Else
         RdoVdr!VEBZIP = txtZip
         RdoVdr.Update
      End If
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub




Private Sub AddVendor()
   Dim bEntry As Boolean
   Dim bResponse As Byte
   Dim sMsg As String
   Dim iNextNumber As Integer
   
   sVendRef = Compress(cmbVnd)
   If sVendRef = "ALL" Then
      Beep
      MsgBox "ALL Is An Illegal Vendor Nickname.", vbExclamation, Caption
      Exit Sub
   End If
   
   If sVendRef = "NONE" Then
      Beep
      MsgBox "NONE Is An Illegal Vendor Nickname.", vbExclamation, Caption
      Exit Sub
   End If
   
   bResponse = IllegalCharacters(cmbVnd)
   If bResponse > 0 Then
      MsgBox "The Nickname Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   
   sMsg = "That Vendor Wasn't Found.  " & vbCr _
          & "Add New " & cmbVnd & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      bEntry = CheckValidColumn(cmbVnd)
      If Not bEntry Then Exit Sub
      
      MouseCursor 13
      On Error GoTo PvndrAv1
      'tab1.Tab = 0
      iNextNumber = GetNextNumber()
      sSql = "INSERT INTO VndrTable (VEREF,VENICKNAME,VENUMBER) " _
             & "VALUES('" & sVendRef & "','" & cmbVnd & "'," & str(iNextNumber) & ") "
      MouseCursor 0
      clsADOCon.ExecuteSQL sSql
      MouseCursor 0
      AddComboStr cmbVnd.hWnd, cmbVnd
      SysMsg "Vendor Added.", True
      On Error Resume Next
      bNewVendor = 1
      bDataHasChanged = False
      bGoodVendor = GetVendor()
      txtNme.SetFocus
   Else
      CancelTrans
      cmbVnd = ""
   End If
   Exit Sub
   
PvndrAv1:
   sProcName = "addvendor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume PvndrAv2
PvndrAv2:
   MouseCursor 0
   ManageBoxes False
   MsgBox CurrError.Description & vbCr _
      & "Couldn't Add Vendor.", vbExclamation, Caption
   DoModuleErrors Me
   
End Sub

Private Function GetNextNumber() As Integer
   Dim RdoVnd As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT MAX(VENUMBER) FROM VndrTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd)
   If bSqlRows Then
      With RdoVnd
         If Not IsNull(.Fields(0)) Then
            GetNextNumber = (.Fields(0)) + 1
         Else
            GetNextNumber = 1
         End If
         ClearResultSet RdoVnd
      End With
   Else
      GetNextNumber = 1
   End If
   Set RdoVnd = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getnextnu"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



Private Sub ManageBoxes(bEnabled As Byte)
   Dim iList As Integer
   For iList = 0 To Controls.Count - 1
      If Err > 0 And (TypeOf Controls(iList) Is TextBox Or _
                      TypeOf Controls(iList) Is ComboBox Or TypeOf Controls(iList) Is MaskEdBox) Then
         If Controls(iList).TabIndex > 1 Then Controls(iList).Enabled = bEnabled
      End If
   Next
   'tab1.Enabled = bEnabled
   
End Sub

Private Sub UpdateVendorSet(Optional ShowMouse As Boolean)
   On Error GoTo DiaErr1
   If ShowMouse Then MouseCursor 13
   With RdoVdr
      !VEACCOUNT = Compress(cmbAct)
      !VEBUYER = "" & Compress(cmbByr)
      !VECSTATE = cmbPste
      !VEBSTATE = cmbSte
      !VE1099 = opt1099.value
      !VEATTORNEY = optAtt.value
      !VEBEOM = optEom.value
      !VEBADR = RTrim(txtAdr)
      !VEBCONTACT = txtCnt
      !VEBCOUNTRY = txtCntr
      !VEBCITY = txtCty
      !VEEMAIL = txtEml
      !VEBEXT = Val(txtExt)
      !VEFAX = "" & txtFax
      !VEINTFAX = txtIntf
      !VEINTPHONE = txtIntp
      !VESINTPHONE = txtInts
      !VEBNAME = Trim(txtNme)
      !VECADR = txtPadr
      !VECCOUNTRY = Trim(txtPcntr)
      !VECCITY = Trim(txtPcty)
      !VEDDAYS = Val(txtPday)
      !VEPROXDT = Val(txtPdte)
      !VEPROXDUE = Val(txtPdue)
      !VEBPHONE = txtPhn
      !VENETDAYS = Val(txtPnet)
      !VECNAME = txtPnme
      !VECZIP = txtPzip
      !VECOMT = RTrim(txtScmt)
      !VESCONTACT = txtScnt
      !VESEXT = Val(txtSext)
      !VEFOB = txtSfob
      !VESPHONE = txtSphn
      !VETAXIDNO = txtStid
      !VETYPE = "" & Trim(cmbTyp)
      .Update
   End With
   bDataHasChanged = False
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   MouseCursor 0
   On Error GoTo 0
   
End Sub
