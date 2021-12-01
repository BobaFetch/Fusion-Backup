VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form VendorEdit01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendors"
   ClientHeight    =   5970
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   7395
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAccountNumber 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Tag             =   "2"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Frame tabFrame 
      Enabled         =   0   'False
      Height          =   3852
      Index           =   3
      Left            =   5000
      TabIndex        =   61
      Tag             =   "3"
      Top             =   1850
      Width           =   7095
      Begin VB.CommandButton cmdAddStat 
         Height          =   375
         Left            =   6360
         Picture         =   "VendorEdit01.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Add Internal Status Code"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox dteSurveyRec 
         Height          =   315
         Left            =   4800
         TabIndex        =   103
         Tag             =   "3"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox dteSurveySent 
         Height          =   315
         Left            =   1680
         TabIndex        =   101
         Tag             =   "3"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox dteReview 
         Height          =   315
         Left            =   4800
         TabIndex        =   98
         Tag             =   "3"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox cbSurvReq 
         Caption         =   "Survey Required?"
         Height          =   255
         Left            =   240
         TabIndex        =   99
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox dteApproved 
         Height          =   315
         Left            =   1680
         TabIndex        =   96
         Tag             =   "3"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox cbApprovalReq 
         Caption         =   "Approval Required?"
         Height          =   255
         Left            =   240
         TabIndex        =   94
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblStatCd 
         Height          =   255
         Left            =   2640
         TabIndex        =   105
         Top             =   2280
         Width           =   3495
      End
      Begin VB.Label lblQuality 
         Caption         =   "Next Review Date"
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   97
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblQuality 
         Caption         =   "Survey Received"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   102
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblQuality 
         Caption         =   "Survey Sent"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   100
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblQuality 
         Caption         =   "Approval Date"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   95
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   3972
      Index           =   2
      Left            =   1
      TabIndex        =   71
      Tag             =   "2"
      Top             =   1850
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
      Left            =   5000
      TabIndex        =   92
      Tag             =   "1"
      Top             =   1850
      Width           =   7092
      Begin VB.TextBox txtArEmail 
         Height          =   285
         Left            =   1320
         TabIndex        =   93
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
         TabIndex        =   25
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox optAtt 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   3360
         TabIndex        =   24
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtInts 
         Height          =   285
         Left            =   1320
         TabIndex        =   18
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
         TabIndex        =   23
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
         TabIndex        =   22
         Tag             =   "3"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtSfob 
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Tag             =   "3"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtSext 
         Height          =   285
         Left            =   3480
         TabIndex        =   20
         Tag             =   "1"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtScnt 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Tag             =   "2"
         Top             =   240
         Width           =   2535
      End
      Begin MSMask.MaskEdBox txtSphn 
         Height          =   288
         Left            =   1704
         TabIndex        =   19
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
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   1035
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
      Left            =   5000
      TabIndex        =   48
      Tag             =   "0"
      Top             =   1850
      Width           =   7176
      Begin VB.ComboBox txtCntr 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Tag             =   "3"
         ToolTipText     =   "Select Country From The List"
         Top             =   1476
         Width           =   2055
      End
      Begin VB.TextBox txtIntf 
         Height          =   285
         Left            =   4680
         TabIndex        =   12
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   1920
         Width           =   475
      End
      Begin VB.TextBox txtIntp 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Tag             =   "1"
         ToolTipText     =   "International Prefix (Country Code)"
         Top             =   1920
         Width           =   475
      End
      Begin VB.ComboBox cmbByr 
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         Tag             =   "3"
         Top             =   3000
         Width           =   2775
      End
      Begin VB.ComboBox cmbSte 
         Height          =   315
         Left            =   4680
         Sorted          =   -1  'True
         TabIndex        =   6
         Tag             =   "3"
         ToolTipText     =   "State Code"
         Top             =   1164
         Width           =   660
      End
      Begin VB.TextBox txtCnt 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Tag             =   "2"
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtExt 
         Height          =   285
         Left            =   3480
         TabIndex        =   11
         Tag             =   "1"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtEml 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Tag             =   "2"
         ToolTipText     =   "Double Click Here To Send E-Mail (Requires An Entry)"
         Top             =   2640
         Width           =   5340
      End
      Begin VB.TextBox txtAdr 
         Height          =   855
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   4
         Tag             =   "9"
         Top             =   240
         Width           =   3475
      End
      Begin VB.TextBox txtCty 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Tag             =   "2"
         Top             =   1164
         Width           =   2085
      End
      Begin VB.TextBox txtZip 
         Height          =   285
         Left            =   5760
         TabIndex        =   7
         Tag             =   "3"
         Top             =   1164
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtPhn 
         Height          =   288
         Left            =   1800
         TabIndex        =   10
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
         TabIndex        =   13
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
      Height          =   4455
      Left            =   0
      TabIndex        =   47
      Top             =   1560
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   7858
      TabWidthStyle   =   2
      TabFixedWidth   =   2293
      TabFixedHeight  =   473
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
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
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Quality"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame z2 
      Height          =   75
      Index           =   0
      Left            =   120
      TabIndex        =   46
      Top             =   1440
      Width           =   7035
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "VendorEdit01.frx":048F
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
      FormDesignHeight=   5970
      FormDesignWidth =   7395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Number"
      Height          =   285
      Index           =   40
      Left            =   120
      TabIndex        =   106
      Top             =   1080
      Width           =   1425
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
Attribute VB_Name = "VendorEdit01"
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
'1/5/2010 BBS Adding Approved Vendor Logic

Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter
Dim AdoVdr As ADODB.Recordset

Dim bGoodVendor As Byte
Dim bNewVendor As Byte
Dim bOnLoad As Byte
Dim bQualTabDisabled As Byte


Dim iTimer As Integer

Dim dToday As Date
Dim sVendRef As String
Dim sState As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtChange() As New EsiKeyBd

'BBS Added logic on 12/27/2009 (Ticket #25640)
'Using this ordinal type to identify the calling
'procedure. DO NOT alter the order of these
'You CAN add to the end of them
Private Enum CallingModule
    NotDefined
    Production
    Finance
    Quality
End Enum

Private Sub DisableModuleBasedTab()
  'Now we can turn off the ones we don't want based on the calling module
   Select Case Val(VendorEdit01.Tag)
    Case Production:    SetEnabledStatusOfFrame 2, False
                        SetEnabledStatusOfFrame 3, False
                        
    Case Finance: SetEnabledStatusOfFrame 3, False
    Case Quality: SetEnabledStatusOfFrame 2, False
   End Select
   SetEnabledFields
   'End of BBS Changes (Ticket #25640)
   
End Sub


'BBS I took the next two routines (FillBuyers and GetCurrentBuyer) from the ESIPROD.BAS Files and made them private.
'It appears that they have been duplicated through a variety of modules
'If these are in fact the correct routines, they probably should be put in a
'common modules somewhere and made public
Private Sub FillBuyers()
   Dim AdoByr As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_GetBuyerList"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoByr, ES_FORWARD)
   If bSqlRows Then
      With AdoByr
         Do Until .EOF
            AddComboStr MDISect.ActiveForm.cmbByr.hwnd, "" & Trim(!BYNUMBER)
            .MoveNext
         Loop
         ClearResultSet AdoByr
      End With
   End If
   Set AdoByr = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillbuyers"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Sub

Private Sub GetCurrentBuyer(sBuyer As String, Optional HideLabel As Byte)
   Dim AdoByr As ADODB.Recordset
   On Error GoTo modErr1
   sBuyer = UCase$(Compress(sBuyer))
   sSql = "SELECT BYNUMBER,BYLSTNAME,BYFSTNAME,BYMIDINIT FROM " _
          & "BuyrTable WHERE BYREF='" & sBuyer & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoByr, ES_FORWARD)
   If bSqlRows Then
      With AdoByr
         If MDISect.ActiveForm.Caption = "Revise A Purchase Order" Then
         
            MDISect.ActiveForm.cmbByr = "" & Trim(!BYNUMBER)
            If HideLabel = 0 Then
               MDISect.ActiveForm.lblByr = "" & Trim(!BYFSTNAME) _
                                           & " " & Trim(!BYMIDINIT) & " " & Trim(!BYLSTNAME)
            End If
         End If
         ClearResultSet AdoByr
      End With
   Else
      If Len(Trim(sBuyer)) > 0 Then
         MDISect.ActiveForm.lblByr = "*** Buyer Wasn't Found ***"
      Else
         MDISect.ActiveForm.lblByr = ""
      End If
   End If
   Set AdoByr = Nothing
   Exit Sub
   
modErr1:
   sProcName = "getcurrentbuyer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Forms(1)
   
End Sub


'This function will enable/disable all the controls in a frame
'where the tag of the frame control array matches the
'tab item index being disabled (starting with 0 as the first tab)
Private Sub SetEnabledStatusOfFrame(ByVal nTag As Integer, ByVal State As Boolean)
Dim C As Control
Dim sTag As String
    If nTag = 3 And Not (State) Then bQualTabDisabled = 1
    On Error Resume Next
    For Each C In Me.Controls
        sTag = ""
        sTag = C.Container.Tag
        If IsNumeric(sTag) Then
'        If nTag = CInt(sTag) And Not (TypeOf C Is Label) Then 'I initially had it not do labels, but changed my mind
        If nTag = CInt(sTag) Then
            C.Enabled = State
        End If
        End If
    Next C
    
    If Err.Number <> 0 Then Err.Clear  ' Error Clean up
    On Error GoTo 0
End Sub

Private Sub SetEnabledFields()
    If (bQualTabDisabled = 0) Then
    
        If cbApprovalReq.Value = 1 Then
            dteApproved.Enabled = True
            dteReview.Enabled = True
        Else
            dteApproved.Enabled = False
            dteReview.Enabled = False
        End If
        If cbSurvReq.Value = 1 Then
            dteSurveySent.Enabled = True
            dteSurveyRec.Enabled = True
        Else
            dteSurveySent.Enabled = False
            dteSurveyRec.Enabled = False
        End If
    End If
    
    
End Sub
'End of BBS Changes (Ticket #25640)



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

'bbs added (new field) 1/6/2010 (Ticket #26450)
Private Sub cbApprovalReq_Click()
    dteApproved.Text = ""
    dteReview.Text = ""
    SetEnabledFields
End Sub

'bbs added (new field) 1/6/2010 (Ticket #26450)
Private Sub cbApprovalReq_KeyPress(KeyAscii As Integer)
    KeyLock KeyAscii
End Sub

'bbs added (new field) 1/6/2010 (Ticket #26450)
Private Sub cbApprovalReq_LostFocus()
   If bGoodVendor Then
        bDataHasChanged = True  'BBS Added on 03/12/2010 for Ticket #29147 to force record to resave
      On Error Resume Next
     
      AdoVdr!VEAPPROVREQ = cbApprovalReq.Value
      If cbApprovalReq.Value <> 1 Then
        AdoVdr!VEAPPDATE = Null
        AdoVdr!VEREVIEWDT = Null
        dteApproved.Text = ""
        dteReview.Text = ""
      End If
      
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
    SetEnabledFields
End Sub



Private Sub cbSurvReq_Click()
    dteSurveySent.Text = ""
    dteSurveyRec.Text = ""
    SetEnabledFields
End Sub

Private Sub cbSurvReq_KeyPress(KeyAscii As Integer)
    KeyLock KeyAscii
End Sub

Private Sub cbSurvReq_LostFocus()
   If bGoodVendor Then
        bDataHasChanged = True  'BBS Added on 03/12/2010 for Ticket #29147 to force record to resave
      On Error Resume Next
      AdoVdr!VESURVEY = cbSurvReq.Value
      If cbSurvReq.Value <> 1 Then
        AdoVdr!VESURVSENT = Null
        AdoVdr!VESURVREC = Null
        dteSurveySent.Text = ""
        dteSurveyRec.Text = ""
      End If
      
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
    SetEnabledFields

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
      AdoVdr!VEACCOUNT = Compress(cmbAct)
      AdoVdr.Update
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
      AdoVdr!VEBUYER = "" & Compress(cmbByr)
      AdoVdr.Update
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
      AdoVdr!VECSTATE = cmbPste
      AdoVdr.Update
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
         AdoVdr!VEBSTATE = cmbSte
         AdoVdr!VECSTATE = cmbSte
         AdoVdr.Update
      Else
         AdoVdr!VEBSTATE = cmbSte
         AdoVdr.Update
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
      AdoVdr!VETYPE = "" & cmbTyp
      AdoVdr.Update
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


Private Sub cmdAddStat_Click()
 StatusCode.lblSCTypeRef = "Vendor ID"
 StatusCode.txtSCTRef = lblNum.Caption
 
 StatusCode.lblSCTRef1 = "0"
 StatusCode.lblSCTRef2 = ""
 StatusCode.lblStatType = "VE"
 StatusCode.lblSysCommIndex = 2 ' The index in the Sys Comment "SO"
 StatusCode.txtCurUser = cUR.CurrentUser

 StatusCode.Show

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


'BBS (1/5/2010) added the following events

Private Sub dteApproved_DropDown()
    ShowCalendar Me
End Sub

Private Sub dteApproved_LostFocus()
    dteApproved = CheckDate(dteApproved)
    If bGoodVendor Then
        bDataHasChanged = True  'BBS Added on 03/12/2010 for Ticket #29147 to force record to resave
        On Error Resume Next
        AdoVdr!VEAPPDATE = Format(dteApproved.Text, "mm/dd/yy")
        AdoVdr.Update
        If Err > 0 Then ValidateEdit
    End If
End Sub


Private Sub dteApproved_Validate(Cancel As Boolean)
    If dteApproved.Text = "" Or dteReview.Text = "" Then Exit Sub
    If (Format(dteReview.Text, "yyyymmdd") < Format(dteApproved.Text, "yyyymmdd")) Then
        MsgBox "Date reviewed cannot be before the date approved", vbInformation + vbOKOnly, "Invalid Date"
        Cancel = True
    End If
    
End Sub

Private Sub dteReview_DropDown()
    ShowCalendar Me
End Sub

Private Sub dteReview_LostFocus()
   dteReview = CheckDate(dteReview)
   If bGoodVendor Then
      bDataHasChanged = True  'BBS Added on 03/12/2010 for Ticket #29147 to force record to resave
      On Error Resume Next
      AdoVdr!VEREVIEWDT = Format(dteReview, "mm/dd/yy")
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
End Sub


Private Sub dteReview_Validate(Cancel As Boolean)
    If dteApproved.Text = "" Or dteReview.Text = "" Then Exit Sub
    If (Format(dteReview.Text, "yyyymmdd") < Format(dteApproved.Text, "yyyymmdd")) Then
        MsgBox "Date reviewed cannot be before the date approved", vbInformation + vbOKOnly, "Invalid Date"
        Cancel = True
        
    End If
End Sub

Private Sub dteSurveyRec_DropDown()
    ShowCalendar Me
End Sub

Private Sub dteSurveyRec_LostFocus()
 dteSurveyRec = CheckDate(dteSurveyRec)
   If bGoodVendor Then
      bDataHasChanged = True  'BBS Added on 03/12/2010 for Ticket #29147 to force record to resave
      On Error Resume Next
      AdoVdr!VESURVREC = Format(dteSurveyRec.Text, "mm/dd/yy")
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
End Sub

Private Sub dteSurveyRec_Validate(Cancel As Boolean)
    If dteSurveyRec.Text = "" Or dteSurveySent.Text = "" Then Exit Sub
    If (Format(dteSurveyRec.Text, "yyyymmdd") < Format(dteSurveySent.Text, "yyyymmdd")) Then
        MsgBox "Date survey received cannot be before the date sent", vbInformation + vbOKOnly, "Invalid Date"
        Cancel = True
        
    End If
    
End Sub

Private Sub dteSurveySent_DropDown()
    ShowCalendar Me
End Sub


Private Sub dteSurveySent_LostFocus()
 dteSurveySent = CheckDate(dteSurveySent)
   If bGoodVendor Then
      bDataHasChanged = True  'BBS Added on 03/12/2010 for Ticket #29147 to force record to resave
      On Error Resume Next
      AdoVdr!VESURVSENT = Format(dteSurveySent.Text, "mm/dd/yy")
      AdoVdr.Update
      If Err > 0 Then ValidateEdit
   End If
End Sub
'End of BBS Changes (1/05/2010)


Private Sub dteSurveySent_Validate(Cancel As Boolean)
    If dteSurveyRec.Text = "" Or dteSurveySent.Text = "" Then Exit Sub
    If (Format(dteSurveyRec.Text, "yyyymmdd") < Format(dteSurveySent.Text, "yyyymmdd")) Then
        MsgBox "Date survey received cannot be before the date sent", vbInformation + vbOKOnly, "Invalid Date"
        Cancel = True
        
    End If
    
End Sub

Private Sub Form_Activate()
Dim I As Integer

   MDISect.lblBotPanel = Caption
   'BBS Added logic on 12/27/2009 (Ticket #25640)
   'Using the tag of the form to help identify the calling procedure
   'using an ordinal type defined above
    For I = 0 To tabFrame.UBound    'First Turn them all on just so we don't have any leftovers from instances
        SetEnabledStatusOfFrame I, True
    Next I
    
   DisableModuleBasedTab    'BBS Procedure to disable controls of tabs based on the calling procedure
   
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
      'FillStatusCode cmbStatCd, Me  'bbs added to fill status codes
      
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim I As Integer
    
   Dim cNewSize As Currency
   MouseCursor 13
   FormLoad Me, ES_LIST, ES_RESIZE
   bQualTabDisabled = 0 'BBS Added this on 1/10/2010
   If bResize = 1 Then
      If lScreenWidth > 9999 Then '640X480
         cNewSize = 1.05
         Height = Height * cNewSize
         Width = Width * cNewSize
      End If
   End If
   'bbs changed the logic below on 1/5/2010 to make it more efficient and portable
    For I = 0 To tabFrame().UBound
        tabFrame(I).BorderStyle = 0
        tabFrame(I).Visible = False
        tabFrame(I).Left = 40
    Next I
    tabFrame(0).Visible = True
    
    
   'tabFrame(0).BorderStyle = 0
   'tabFrame(1).BorderStyle = 0
   'tabFrame(2).BorderStyle = 0
   'tabFrame(0).Visible = True
   'tabFrame(1).Visible = False
   'tabFrame(2).Visible = False
   'tabFrame(0).Left = 40
   'tabFrame(1).Left = 40
   'tabFrame(2).Left = 40
   'tabFrame(3).Left = 40
   
   
   FormatControls
   
   sSql = "SELECT * FROM VndrTable " _
          & "WHERE (VEREF = ? AND VEREF<>'NONE') "
          
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 10
   
   AdoQry.Parameters.Append AdoParameter
   
   'tab1.Tab = 0
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
   Set AdoVdr = Nothing
   Set AdoQry = Nothing
   Set VendorEdit01 = Nothing
   
End Sub



Private Function GetVendor(Optional bUpdate As Byte) As Byte
   If bDataHasChanged And bGoodVendor = 1 Then UpdateVendorSet
   sVendRef = Compress(cmbVnd)
   On Error GoTo DiaErr1
   ClearBoxes
   bGoodVendor = 0
   AdoQry.Parameters(0).Value = sVendRef
   bSqlRows = clsADOCon.GetQuerySet(AdoVdr, AdoQry, ES_KEYSET, True, 1)
   If bSqlRows Then
      With AdoVdr
         On Error Resume Next
         cmbVnd = "" & Trim(!VENICKNAME)
         lblNum = 0 + Format(!VENUMBER, "#0")
         txtNme = "" & Trim(!VEBNAME)
         cmbTyp = "" & Trim(!VETYPE)
         txtAdr = "" & TrimAddress(!VEBADR)
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
         If !VESEXT > 0 Then txtSExt = !VESEXT
         txtPnme = "" & Trim(!VECNAME)
         txtPadr = "" & TrimAddress(!VECADR)
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
         txtPDue = 0 + Format(!VEPROXDUE, "#0")
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
         optAtt.Value = !VEATTORNEY
         '1/8/07
         txtArEmail = "" & Trim(!VEAREMAIL)
         'BBS Added the below fields on 1/6/2010 for Ticket #25640
         cbApprovalReq = 0 + !VEAPPROVREQ
         cbSurvReq = 0 + !VESURVEY
         dteApproved = Format(!VEAPPDATE, "mm/dd/yy")
         dteSurveySent = Format(!VESURVSENT, "mm/dd/yy")
         dteSurveyRec = Format(!VESURVREC, "mm/dd/yy")
         dteReview = Format(!VEREVIEWDT, "mm/dd/yy")
         txtAccountNumber = "" & Trim(!VEACCTNO)
      End With
      If Len(Trim(cmbByr)) > 0 Then GetCurrentBuyer cmbByr 'bbs
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
   opt1099.Value = vbUnchecked
   txtPhn.Mask = ""
   txtPhn = ""
   txtPhn.Mask = "###-###-####"
   txtFax.Mask = ""
   txtFax = ""
   txtFax.Mask = "###-###-####"
   txtSphn.Mask = ""
   txtSphn = ""
   txtSphn.Mask = "###-###-####"
   'BBS Added on 1/6/2010 for Ticket #25640
   'cbApprovalReq.value = vbChecked
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
      AdoVdr!VE1099 = opt1099.Value
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub optAtt_Click()
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEATTORNEY = optAtt.Value
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub optAtt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optEom_Click()
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEBEOM = optEom.Value
      AdoVdr.Update
   End If
   
End Sub

Private Sub optEom_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub tab1_Click()
'bbs changed the logic below to make it more efficient and portable on 1/5/2010

    Dim I As Integer

    On Error Resume Next
    ClearCombos
   
    For I = 0 To tabFrame.UBound
        tabFrame(I).Visible = False
    Next I
    tabFrame(tab1.SelectedItem.Index - 1).Visible = True
    Select Case tab1.SelectedItem.Index
    Case 1: txtAdr.SetFocus
    Case 2: txtScnt.SetFocus
    Case 3: txtPnme.SetFocus
    Case 4:
    End Select
    
    
    
    

   
'   If tab1.SelectedItem.Index = 1 Then
'      tabFrame(0).Visible = True
'      tabFrame(1).Visible = False
'      tabFrame(2).Visible = False
'      txtAdr.SetFocus
'   ElseIf tab1.SelectedItem.Index = 2 Then
'      tabFrame(1).Visible = True
'      tabFrame(0).Visible = False
'      tabFrame(2).Visible = False
'      txtScnt.SetFocus
'
'   Else
'      tabFrame(2).Visible = True
'      tabFrame(0).Visible = False
'      tabFrame(1).Visible = False
'      txtPnme.SetFocus
   'End If
   
End Sub

Private Sub Tab1_GotFocus()
   Dim bReponse As Byte
   Dim sMsg As String
   iTimer = 0
   If iTimer > 300 Then
      '6000/300 sec 30 min
      On Error Resume Next
      AdoVdr.Close
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
      AdoVdr.Close
      sMsg = Format(Now, "mm/dd/yy hh:mm AM/PM") _
             & " This Normal Time Out." & vbCr _
             & "All Current Data Has Been Saved."
      MsgBox sMsg, vbInformation, Caption
      Timer1(0).Enabled = False
      Unload Me
   End If
   
End Sub



Private Sub txtAccountNumber_LostFocus()
   txtAccountNumber = CheckLen(txtAccountNumber, 30)
   txtAccountNumber = StrCase(txtAccountNumber)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEACCTNO = txtAccountNumber
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtAdr_LostFocus()
   txtAdr = CheckLen(txtAdr, 160)
   If bGoodVendor Then
      On Error Resume Next
      If bNewVendor Then
         txtPadr = txtAdr
         AdoVdr!VEBADR = TrimAddress(txtAdr)
         AdoVdr!VECADR = TrimAddress(txtAdr)
         AdoVdr!VEDATEREV = dToday
         AdoVdr.Update
      Else
         AdoVdr!VEBADR = TrimAddress(txtAdr)
         AdoVdr!VEDATEREV = dToday
         AdoVdr.Update
      End If
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

'Private Function TrimAddress(adr As String)
'    'remove trailing blanks and trailing CRLFs
'    Dim rtn As String
'
'    rtn = adr
'    Do While Len(rtn) > 0
'        rtn = Trim(rtn)
'        If Len(rtn) >= 2 Then
'            If Right(Trim(rtn), 2) = vbCrLf Then
'                rtn = Trim(Left(rtn, Len(rtn) - 2))
'            Else
'                Exit Do
'            End If
'        Else
'            Exit Do
'        End If
'    Loop
'    TrimAddress = rtn
'End Function
'
Private Sub txtArEmail_DblClick()
   If Trim(txtArEmail) <> "" Then SendEMail Trim(txtArEmail)
   
End Sub

Private Sub txtArEmail_LostFocus()
   txtArEmail = CheckLen(txtArEmail, 60)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEAREMAIL = txtArEmail
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtCnt_LostFocus()
   txtCnt = CheckLen(txtCnt, 20)
   txtCnt = StrCase(txtCnt)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEBCONTACT = txtCnt
      AdoVdr.Update
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
         AdoVdr!VEBCOUNTRY = txtCntr
         AdoVdr!VECCOUNTRY = txtCntr
         AdoVdr.Update
      Else
         AdoVdr!VEBCOUNTRY = txtCntr
         AdoVdr.Update
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
         AdoVdr!VEBCITY = txtCty
         AdoVdr!VECCITY = txtCty
         AdoVdr.Update
      Else
         AdoVdr!VEBCITY = txtCty
         AdoVdr.Update
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
      AdoVdr!VEEMAIL = txtEml
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtExt_LostFocus()
   txtExt = CheckLen(txtExt, 4)
   If Len(txtExt) > 0 Then txtExt = Format(Abs(Val(txtExt)), "###0")
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEBEXT = Val(txtExt)
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtFax_LostFocus()
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEFAX = "" & txtFax
      AdoVdr.Update
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
      AdoVdr!VEINTFAX = txtIntf
      AdoVdr.Update
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
      AdoVdr!VEINTPHONE = txtIntp
      AdoVdr.Update
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
      AdoVdr!VESINTPHONE = txtInts
      AdoVdr.Update
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
         AdoVdr!VEBNAME = txtNme
         AdoVdr!VECNAME = txtNme
         AdoVdr.Update
      Else
         AdoVdr!VEBNAME = txtNme
         AdoVdr.Update
      End If
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtPadr_LostFocus()
   txtPadr = CheckLen(txtPadr, 160)
   iTimer = 0
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VECADR = TrimAddress(txtPadr)
      AdoVdr.Update
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
      AdoVdr!VECCOUNTRY = txtPcntr
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   bNewVendor = 0
   
End Sub


Private Sub txtPcty_LostFocus()
   txtPcty = CheckLen(txtPcty, 18)
   txtPcty = StrCase(txtPcty)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VECCITY = txtPcty
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtPday_LostFocus()
   txtPday = CheckLen(txtPday, 3)
   If Len(txtPday) > 0 Then txtPday = Format(Abs(Val(txtPday)), "##0")
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEDDAYS = Val(txtPday)
      AdoVdr.Update
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
      AdoVdr!VEDISCOUNT = Val(txtPdsc(Index))
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtPdte_LostFocus()
   txtPdte = CheckLen(txtPdte, 2)
   If Len(txtPdte) > 0 Then txtPdte = Format(Abs(Val(txtPdte)), "#0")
   If Val(txtPdte) > 30 Then txtPdte = "30"
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEPROXDT = Val(txtPdte)
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtPdue_LostFocus()
   txtPDue = CheckLen(txtPDue, 2)
   If Len(txtPDue) > 0 Then txtPDue = Format(Abs(Val(txtPDue)), "#0")
   If Val(txtPDue) > 30 Then txtPDue = "30"
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEPROXDUE = Val(txtPDue)
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtPhn_LostFocus()
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEBPHONE = txtPhn
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtPnet_LostFocus()
   txtPnet = CheckLen(txtPnet, 3)
   If Len(txtPnet) > 0 Then txtPnet = Format(Abs(Val(txtPnet)), "##0")
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VENETDAYS = Val(txtPnet)
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtPnme_LostFocus()
   txtPnme = CheckLen(txtPnme, 40)
   txtPnme = StrCase(txtPnme)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VECNAME = txtPnme
      AdoVdr!VEDATEREV = dToday
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtPzip_LostFocus()
   txtPzip = CheckLen(txtPzip, 10)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VECZIP = txtPzip
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtScmt_LostFocus()
   txtScmt = CheckLen(txtScmt, 2048)
   txtScmt = StrCase(txtScmt, ES_FIRSTWORD)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VECOMT = txtScmt
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtScnt_LostFocus()
   txtScnt = CheckLen(txtScnt, 20)
   txtScnt = StrCase(txtScnt)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VESCONTACT = txtScnt
      AdoVdr!VEDATEREV = dToday
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtSext_LostFocus()
   txtSExt = CheckLen(txtSExt, 4)
   If Len(txtSExt) > 0 Then txtSExt = Format(Abs(Val(txtSExt)), "###0")
   txtScmt = CheckLen(txtScmt, 255)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VESEXT = Val(txtSExt)
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub

Private Sub txtSfob_LostFocus()
   txtSfob = CheckLen(txtSfob, 12)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VEFOB = txtSfob
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtSphn_LostFocus()
   iTimer = 0
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VESPHONE = txtSphn
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtStid_LostFocus()
   txtStid = CheckLen(txtStid, 20)
   If bGoodVendor Then
      On Error Resume Next
      AdoVdr!VETAXIDNO = txtStid
      AdoVdr.Update
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub


Private Sub txtZip_LostFocus()
   txtZip = CheckLen(txtZip, 10)
   If bGoodVendor Then
      On Error Resume Next
      If bNewVendor Then
         txtPzip = txtZip
         AdoVdr!VEBZIP = txtZip
         AdoVdr!VECZIP = txtZip
         AdoVdr.Update
      Else
         AdoVdr!VEBZIP = txtZip
         AdoVdr.Update
      End If
      If Err > 0 Then bGoodVendor = GetVendor()
   End If
   
End Sub




Private Sub AddVendor()
   Dim bEntry As Boolean
   Dim bResponse As Byte
   Dim sMsg As String
   Dim iNextNumber As Integer
   Dim strVendorApproval As String
   
   
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
      strVendorApproval = GetPreferenceValue("ReqVendorApproval")
      If strVendorApproval = "" Then strVendorApproval = "0"
      
      'BBS Changed query below to default values to true for a new record (Ticket #25640)
      sSql = "INSERT INTO VndrTable (VEREF,VENICKNAME,VENUMBER,VEAPPROVREQ,VESURVEY) " _
             & "VALUES('" & sVendRef & "','" & cmbVnd & "'," & str(iNextNumber) & ", " & strVendorApproval & "," & strVendorApproval & ") "
      MouseCursor 0
      clsADOCon.ExecuteSql sSql
      MouseCursor 0
      AddComboStr cmbVnd.hwnd, cmbVnd
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
   Dim AdoVnd As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT MAX(VENUMBER) FROM VndrTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoVnd)
   If bSqlRows Then
      With AdoVnd
         If Not IsNull(.Fields(0)) Then
            GetNextNumber = (.Fields(0)) + 1
         Else
            GetNextNumber = 1
         End If
         ClearResultSet AdoVnd
      End With
   Else
      GetNextNumber = 1
   End If
   Set AdoVnd = Nothing
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
   DisableModuleBasedTab ' Make sure we haven't enabled anything that shouldn't based on calling procedure
End Sub

Private Sub UpdateVendorSet(Optional ShowMouse As Boolean)
   On Error GoTo DiaErr1
   If ShowMouse Then MouseCursor 13
   With AdoVdr
      !VEACCOUNT = Compress(cmbAct)
      !VEBUYER = "" & Compress(cmbByr)
      !VECSTATE = cmbPste
      !VEBSTATE = cmbSte
      !VE1099 = opt1099.Value
      !VEATTORNEY = optAtt.Value
      !VEBEOM = optEom.Value
      !VEBADR = TrimAddress(txtAdr)
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
      !VECADR = TrimAddress(txtPadr)
      !VECCOUNTRY = Trim(txtPcntr)
      !VECCITY = Trim(txtPcty)
      !VEDDAYS = Val(txtPday)
      !VEPROXDT = Val(txtPdte)
      !VEPROXDUE = Val(txtPDue)
      !VEBPHONE = txtPhn
      !VENETDAYS = Val(txtPnet)
      !VECNAME = txtPnme
      !VECZIP = txtPzip
      !VECOMT = RTrim(txtScmt)
      !VESCONTACT = txtScnt
      !VESEXT = Val(txtSExt)
      !VEFOB = txtSfob
      !VESPHONE = txtSphn
      !VETAXIDNO = txtStid
      !VETYPE = "" & Trim(cmbTyp)
      !VEACCTNO = "" & Trim(txtAccountNumber)
      
      'BBS added this on 03/12/2010 to make sure the data matches the screen for Ticket #29147
      If (dteSurveySent.Text <> "") Or (dteSurveyRec.Text <> "") Then cbSurvReq.Value = 1
      If (dteApproved.Text <> "") Or (dteReview.Text <> "") Then cbSurvReq.Value = 1

      
      'BBS Added the below fields for Ticket #25640
      !VEAPPROVREQ = cbApprovalReq.Value
      !VESURVEY = cbSurvReq.Value
      !VESURVSENT = Format(dteSurveySent.Text, "mm/dd/yy")
      !VESURVREC = Format(dteSurveyRec.Text, "mm/dd/yy")
      !VEAPPDATE = Format(dteApproved.Text, "mm/dd/yy")
      !VEREVIEWDT = Format(dteReview.Text, "mm/dd/yy")
      .Update
   End With
   bDataHasChanged = False
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   MouseCursor 0
   On Error GoTo 0
   
End Sub
