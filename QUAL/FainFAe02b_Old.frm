VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FainFAe02b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "First Article Report"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VB.Frame tabFrame 
      Height          =   4332
      Index           =   1
      Left            =   8100
      TabIndex        =   76
      Top             =   1380
      Visible         =   0   'False
      Width           =   7275
      Begin VB.TextBox txtPrefix 
         Height          =   315
         Index           =   10
         Left            =   600
         TabIndex        =   36
         Top             =   3720
         Width           =   600
      End
      Begin VB.TextBox txtPrefix 
         Height          =   315
         Index           =   9
         Left            =   600
         TabIndex        =   32
         Top             =   3360
         Width           =   600
      End
      Begin VB.TextBox txtPrefix 
         Height          =   315
         Index           =   8
         Left            =   600
         TabIndex        =   28
         Top             =   3000
         Width           =   600
      End
      Begin VB.TextBox txtPrefix 
         Height          =   315
         Index           =   7
         Left            =   600
         TabIndex        =   24
         Top             =   2640
         Width           =   600
      End
      Begin VB.TextBox txtPrefix 
         Height          =   315
         Index           =   6
         Left            =   600
         TabIndex        =   20
         Top             =   2280
         Width           =   600
      End
      Begin VB.TextBox txtPrefix 
         Height          =   315
         Index           =   5
         Left            =   600
         TabIndex        =   16
         Top             =   1920
         Width           =   600
      End
      Begin VB.TextBox txtPrefix 
         Height          =   315
         Index           =   4
         Left            =   600
         TabIndex        =   12
         Top             =   1560
         Width           =   600
      End
      Begin VB.TextBox txtPrefix 
         Height          =   315
         Index           =   3
         Left            =   600
         TabIndex        =   8
         Top             =   1200
         Width           =   600
      End
      Begin VB.TextBox txtPrefix 
         Height          =   315
         Index           =   2
         Left            =   600
         TabIndex        =   4
         Top             =   840
         Width           =   600
      End
      Begin VB.TextBox txtPrefix 
         Height          =   315
         Index           =   1
         Left            =   600
         TabIndex        =   0
         Top             =   480
         Width           =   600
      End
      Begin VB.ComboBox txtDch 
         Height          =   315
         Index           =   10
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Tag             =   "3"
         ToolTipText     =   "Document Revision Or Change"
         Top             =   3720
         Width           =   1572
      End
      Begin VB.ComboBox txtDch 
         Height          =   315
         Index           =   9
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Tag             =   "3"
         ToolTipText     =   "Document Revision Or Change"
         Top             =   3360
         Width           =   1572
      End
      Begin VB.ComboBox txtDch 
         Height          =   315
         Index           =   8
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Tag             =   "3"
         ToolTipText     =   "Document Revision Or Change"
         Top             =   3000
         Width           =   1572
      End
      Begin VB.ComboBox txtDch 
         Height          =   315
         Index           =   7
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Tag             =   "3"
         ToolTipText     =   "Document Revision Or Change"
         Top             =   2640
         Width           =   1572
      End
      Begin VB.ComboBox txtDch 
         Height          =   315
         Index           =   6
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Tag             =   "3"
         ToolTipText     =   "Document Revision Or Change"
         Top             =   2280
         Width           =   1572
      End
      Begin VB.ComboBox txtDch 
         Height          =   315
         Index           =   5
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Tag             =   "3"
         ToolTipText     =   "Document Revision Or Change"
         Top             =   1920
         Width           =   1572
      End
      Begin VB.ComboBox txtDch 
         Height          =   315
         Index           =   4
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Tag             =   "3"
         ToolTipText     =   "Document Revision Or Change"
         Top             =   1560
         Width           =   1572
      End
      Begin VB.ComboBox txtDch 
         Height          =   315
         Index           =   3
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "3"
         ToolTipText     =   "Document Revision Or Change"
         Top             =   1200
         Width           =   1572
      End
      Begin VB.ComboBox txtDch 
         Height          =   315
         Index           =   2
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "3"
         ToolTipText     =   "Document Revision Or Change"
         Top             =   840
         Width           =   1572
      End
      Begin VB.ComboBox txtDch 
         Height          =   315
         Index           =   1
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "3"
         ToolTipText     =   "Document Revision Or Change"
         Top             =   480
         Width           =   1572
      End
      Begin VB.ComboBox txtDsh 
         Height          =   315
         Index           =   10
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Tag             =   "3"
         ToolTipText     =   "Document Sheet Or Page"
         Top             =   3720
         Width           =   972
      End
      Begin VB.ComboBox txtDsh 
         Height          =   315
         Index           =   9
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Tag             =   "3"
         ToolTipText     =   "Document Sheet Or Page"
         Top             =   3360
         Width           =   972
      End
      Begin VB.ComboBox txtDsh 
         Height          =   315
         Index           =   8
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Tag             =   "3"
         ToolTipText     =   "Document Sheet Or Page"
         Top             =   3000
         Width           =   972
      End
      Begin VB.ComboBox txtDsh 
         Height          =   315
         Index           =   7
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Tag             =   "3"
         ToolTipText     =   "Document Sheet Or Page"
         Top             =   2640
         Width           =   972
      End
      Begin VB.ComboBox txtDsh 
         Height          =   315
         Index           =   6
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Tag             =   "3"
         ToolTipText     =   "Document Sheet Or Page"
         Top             =   2280
         Width           =   972
      End
      Begin VB.ComboBox txtDsh 
         Height          =   315
         Index           =   5
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Tag             =   "3"
         ToolTipText     =   "Document Sheet Or Page"
         Top             =   1920
         Width           =   972
      End
      Begin VB.ComboBox txtDsh 
         Height          =   315
         Index           =   4
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Tag             =   "3"
         ToolTipText     =   "Document Sheet Or Page"
         Top             =   1560
         Width           =   972
      End
      Begin VB.ComboBox txtDsh 
         Height          =   315
         Index           =   3
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Tag             =   "3"
         ToolTipText     =   "Document Sheet Or Page"
         Top             =   1200
         Width           =   972
      End
      Begin VB.ComboBox txtDsh 
         Height          =   315
         Index           =   2
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "3"
         ToolTipText     =   "Document Sheet Or Page"
         Top             =   840
         Width           =   972
      End
      Begin VB.ComboBox txtDsh 
         Height          =   315
         Index           =   1
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "3"
         ToolTipText     =   "Document Sheet Or Page"
         Top             =   480
         Width           =   972
      End
      Begin VB.ComboBox txtDno 
         Height          =   315
         Index           =   10
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Tag             =   "3"
         ToolTipText     =   "Document Number/Name"
         Top             =   3720
         Width           =   3012
      End
      Begin VB.ComboBox txtDno 
         Height          =   315
         Index           =   9
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Tag             =   "3"
         ToolTipText     =   "Document Number/Name"
         Top             =   3360
         Width           =   3012
      End
      Begin VB.ComboBox txtDno 
         Height          =   315
         Index           =   8
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Tag             =   "3"
         ToolTipText     =   "Document Number/Name"
         Top             =   3000
         Width           =   3012
      End
      Begin VB.ComboBox txtDno 
         Height          =   315
         Index           =   7
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Tag             =   "3"
         ToolTipText     =   "Document Number/Name"
         Top             =   2640
         Width           =   3012
      End
      Begin VB.ComboBox txtDno 
         Height          =   315
         Index           =   6
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Tag             =   "3"
         ToolTipText     =   "Document Number/Name"
         Top             =   2280
         Width           =   3012
      End
      Begin VB.ComboBox txtDno 
         Height          =   315
         Index           =   5
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "3"
         ToolTipText     =   "Document Number/Name"
         Top             =   1920
         Width           =   3012
      End
      Begin VB.ComboBox txtDno 
         Height          =   315
         Index           =   4
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "3"
         ToolTipText     =   "Document Number/Name"
         Top             =   1560
         Width           =   3012
      End
      Begin VB.ComboBox txtDno 
         Height          =   315
         Index           =   3
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Tag             =   "3"
         ToolTipText     =   "Document Number/Name"
         Top             =   1200
         Width           =   3012
      End
      Begin VB.ComboBox txtDno 
         Height          =   315
         Index           =   2
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "3"
         ToolTipText     =   "Document Number/Name"
         Top             =   840
         Width           =   3012
      End
      Begin VB.ComboBox txtDno 
         Height          =   315
         Index           =   1
         ItemData        =   "FainFAe02b.frx":0000
         Left            =   1320
         List            =   "FainFAe02b.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "3"
         ToolTipText     =   "Document Number/Name"
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Prefix               "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   600
         TabIndex        =   93
         Top             =   240
         Width           =   600
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Revision/Change       "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   5520
         TabIndex        =   89
         Top             =   240
         Width           =   2985
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sheet         "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   4470
         TabIndex        =   88
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Drawing/Document                                    "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   1320
         TabIndex        =   87
         Top             =   240
         Width           =   2985
      End
      Begin VB.Label lblDoc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Index           =   10
         Left            =   120
         TabIndex        =   86
         Top             =   3720
         Width           =   372
      End
      Begin VB.Label lblDoc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Index           =   9
         Left            =   120
         TabIndex        =   85
         Top             =   3360
         Width           =   372
      End
      Begin VB.Label lblDoc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Index           =   8
         Left            =   120
         TabIndex        =   84
         Top             =   3000
         Width           =   372
      End
      Begin VB.Label lblDoc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Index           =   7
         Left            =   120
         TabIndex        =   83
         Top             =   2640
         Width           =   372
      End
      Begin VB.Label lblDoc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Index           =   6
         Left            =   120
         TabIndex        =   82
         Top             =   2280
         Width           =   372
      End
      Begin VB.Label lblDoc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Index           =   5
         Left            =   120
         TabIndex        =   81
         Top             =   1920
         Width           =   372
      End
      Begin VB.Label lblDoc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Index           =   4
         Left            =   120
         TabIndex        =   80
         Top             =   1560
         Width           =   372
      End
      Begin VB.Label lblDoc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Index           =   3
         Left            =   120
         TabIndex        =   79
         Top             =   1200
         Width           =   372
      End
      Begin VB.Label lblDoc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Index           =   2
         Left            =   120
         TabIndex        =   78
         Top             =   840
         Width           =   372
      End
      Begin VB.Label lblDoc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Index           =   1
         Left            =   120
         TabIndex        =   77
         Top             =   480
         Width           =   372
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   4332
      Index           =   0
      Left            =   240
      TabIndex        =   63
      Top             =   1380
      Width           =   7212
      Begin VB.TextBox txtRun 
         Height          =   285
         Left            =   6000
         TabIndex        =   50
         Tag             =   "1"
         ToolTipText     =   "Manufacturing Order And Run (Optional)"
         Top             =   2280
         Width           =   852
      End
      Begin VB.TextBox txtMon 
         Height          =   285
         Left            =   2160
         TabIndex        =   49
         Tag             =   "3"
         ToolTipText     =   "Manufacturing Order And Run (Optional)"
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox txtCmt 
         Height          =   1335
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Tag             =   "9"
         ToolTipText     =   "Comments"
         Top             =   2760
         Width           =   5295
      End
      Begin VB.TextBox txtAnp 
         Height          =   285
         Left            =   2160
         TabIndex        =   47
         Tag             =   "2"
         ToolTipText     =   "Degrees"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtAnm 
         Height          =   285
         Left            =   3720
         TabIndex        =   48
         Tag             =   "2"
         ToolTipText     =   "Degrees"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtDwm 
         Height          =   285
         Left            =   3720
         TabIndex        =   46
         Tag             =   "2"
         ToolTipText     =   "Decimal"
         Top             =   1560
         Width           =   615
      End
      Begin VB.CheckBox optCom 
         Caption         =   "____"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   6000
         TabIndex        =   44
         ToolTipText     =   "The Report Is Complete"
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox txtIns 
         Height          =   288
         Left            =   6000
         TabIndex        =   43
         Tag             =   "4"
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cmbIns 
         Height          =   288
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   42
         Tag             =   "3"
         ToolTipText     =   "Select Inspector ID From List"
         Top             =   600
         Width           =   1665
      End
      Begin VB.ComboBox txtBeg 
         Height          =   288
         Left            =   6000
         TabIndex        =   41
         Tag             =   "4"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtDwp 
         Height          =   285
         Left            =   2160
         TabIndex        =   45
         Tag             =   "2"
         ToolTipText     =   "Decimal"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtDsc 
         Height          =   285
         Left            =   1440
         TabIndex        =   40
         Tag             =   "2"
         ToolTipText     =   "Up To 30 Chars"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Run"
         Height          =   252
         Index           =   17
         Left            =   5400
         TabIndex        =   92
         Top             =   2280
         Width           =   732
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacturing Order"
         Height          =   252
         Index           =   16
         Left            =   240
         TabIndex        =   91
         Top             =   2280
         Width           =   1900
      End
      Begin VB.Label Label1 
         Caption         =   "More >>>>"
         Height          =   255
         Left            =   5820
         TabIndex        =   90
         Top             =   1740
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments:"
         Height          =   252
         Index           =   12
         Left            =   240
         TabIndex        =   75
         Top             =   2760
         Width           =   2292
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Angle-"
         Height          =   252
         Index           =   11
         Left            =   3000
         TabIndex        =   74
         Top             =   1920
         Width           =   972
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ang +"
         Height          =   252
         Index           =   10
         Left            =   1440
         TabIndex        =   73
         Top             =   1920
         Width           =   732
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Drwg -"
         Height          =   252
         Index           =   9
         Left            =   3000
         TabIndex        =   72
         Top             =   1560
         Width           =   1212
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Drwg +"
         Height          =   252
         Index           =   8
         Left            =   1440
         TabIndex        =   71
         Top             =   1560
         Width           =   1212
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Drawing Tolerances:"
         Height          =   252
         Index           =   7
         Left            =   120
         TabIndex        =   70
         Top             =   1320
         Width           =   2292
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Complete"
         Height          =   288
         Index           =   6
         Left            =   4560
         TabIndex        =   69
         Top             =   960
         Width           =   1548
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Inspected"
         Height          =   288
         Index           =   5
         Left            =   4560
         TabIndex        =   68
         Top             =   600
         Width           =   1548
      End
      Begin VB.Label lblNme 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   1440
         TabIndex        =   67
         Top             =   960
         Width           =   2892
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inspector Id"
         Height          =   252
         Index           =   4
         Left            =   120
         TabIndex        =   66
         Top             =   600
         Width           =   1212
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   288
         Index           =   3
         Left            =   4560
         TabIndex        =   65
         Top             =   240
         Width           =   1548
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   288
         Index           =   2
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   1548
      End
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   4935
      Left            =   60
      TabIndex        =   62
      Top             =   960
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   8705
      TabWidthStyle   =   2
      TabFixedWidth   =   1764
      TabFixedHeight  =   473
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Documents"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "FainFAe02b.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   61
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optPrn 
      Caption         =   "Check1"
      Height          =   195
      Left            =   5640
      TabIndex        =   60
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdPrn 
      DownPicture     =   "FainFAe02b.frx":07B2
      Height          =   320
      Left            =   5880
      Picture         =   "FainFAe02b.frx":093C
      Style           =   1  'Graphical
      TabIndex        =   59
      TabStop         =   0   'False
      ToolTipText     =   "Print/Display Report"
      Top             =   480
      Width           =   350
   End
   Begin VB.CommandButton cmdItm 
      Caption         =   "&Items"
      Height          =   315
      Left            =   6480
      TabIndex        =   58
      ToolTipText     =   "Add/Revise Tag Items"
      Top             =   480
      Width           =   875
   End
   Begin VB.CheckBox optFrom 
      Caption         =   "From New"
      Height          =   255
      Left            =   1560
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6480
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8400
      Top             =   360
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5970
      FormDesignWidth =   7770
   End
   Begin VB.Label txtRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   57
      Top             =   600
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Revision"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   56
      Top             =   600
      Width           =   1545
   End
   Begin VB.Label txtPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   55
      Top             =   270
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Number"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   54
      Top             =   270
      Width           =   1545
   End
End
Attribute VB_Name = "FainFAe02b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'8/9/06 Replaced Tab with TabStrip
'10/4/06 Added Manufacturing Order and Run
'10/6/06 Added ComboBoxes for Documents, Shees,Revs (no relationship req'd)
Option Explicit
Dim RdoRpt As ADODB.Recordset
Dim bGoodIns As Byte
Dim bGoodReport As Byte
Dim bOnLoad As Byte

'Dim sDocuments(11, 4) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private prefixValueAtGotFocus As String

Private Sub GetDocSheets(iIndex As Integer, Optional useSheet As String = "use old value")
   ' if sheet specified, select it
   Dim sOldSheet As String
   sOldSheet = txtDsh(iIndex)
   txtDsh(iIndex).Clear
   If txtDno(iIndex) = "" Then Exit Sub
   
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT RTRIM(DOSHEET) as DOSHEET FROM DdocTable " _
          & "WHERE DOREF='" & Compress(txtDno(iIndex)) & "' " _
          & "ORDER BY DOSHEET"
   LoadComboBox txtDsh(iIndex), -1, False      ' don't select first item
'   If useSheet = "use old value" Then
'      txtDsh(iIndex) = sOldSheet
'   Else
'      txtDsh(iIndex) = useSheet
'   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getsheets"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillDocuments()
   Dim RdoDoc As ADODB.Recordset
   Dim iRow As Integer
   Dim sDoc(11, 4) As String
   
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM FadcTable WHERE FA_DOCNUMBER='" _
          & Compress(txtPrt) & "' AND FA_DOCREVISION='" _
          & txtRev & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_FORWARD)
   If bSqlRows Then
      With RdoDoc
         Do Until .EOF
            iRow = iRow + 1
            sDoc(iRow, 1) = "" & Trim(!FA_DOCDESCRIPTION)
            sDoc(iRow, 2) = "" & Trim(!FA_DOCSHEET)
            sDoc(iRow, 3) = "" & Trim(!FA_DOCCHANGE)
            .MoveNext
         Loop
         ClearResultSet RdoDoc
      End With
   End If
   sSql = "SELECT * FROM EsReportFarp01 WHERE FA_RECORD=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_KEYSET)
   If bSqlRows Then
      With RdoDoc
         !FA_DOCNUMBER = Compress(txtPrt)
         !FA_DOCDESCRIPTION1 = sDoc(1, 1)
         !FA_DOCSHEET1 = sDoc(1, 2)
         !FA_DOCCHANGE1 = sDoc(1, 3)
         
         !FA_DOCDESCRIPTION2 = sDoc(2, 1)
         !FA_DOCSHEET2 = sDoc(2, 2)
         !FA_DOCCHANGE2 = sDoc(2, 3)
         
         !FA_DOCDESCRIPTION3 = sDoc(3, 1)
         !FA_DOCSHEET3 = sDoc(3, 2)
         !FA_DOCCHANGE3 = sDoc(3, 3)
         
         !FA_DOCDESCRIPTION4 = sDoc(4, 1)
         !FA_DOCSHEET4 = sDoc(4, 2)
         !FA_DOCCHANGE4 = sDoc(4, 3)
         
         !FA_DOCDESCRIPTION5 = sDoc(5, 1)
         !FA_DOCSHEET5 = sDoc(5, 2)
         !FA_DOCCHANGE5 = sDoc(5, 3)
         
         !FA_DOCDESCRIPTION6 = sDoc(6, 1)
         !FA_DOCSHEET6 = sDoc(6, 2)
         !FA_DOCCHANGE6 = sDoc(6, 3)
         
         !FA_DOCDESCRIPTION7 = sDoc(7, 1)
         !FA_DOCSHEET7 = sDoc(7, 2)
         !FA_DOCCHANGE7 = sDoc(7, 3)
         
         !FA_DOCDESCRIPTION8 = sDoc(8, 1)
         !FA_DOCSHEET8 = sDoc(8, 2)
         !FA_DOCCHANGE8 = sDoc(8, 3)
         
         !FA_DOCDESCRIPTION9 = sDoc(9, 1)
         !FA_DOCSHEET9 = sDoc(9, 2)
         !FA_DOCCHANGE9 = sDoc(9, 3)
         
         !FA_DOCDESCRIPTION10 = sDoc(10, 1)
         !FA_DOCSHEET10 = sDoc(10, 2)
         !FA_DOCCHANGE10 = sDoc(10, 3)
         .update
         ClearResultSet RdoDoc
      End With
   End If
   Set RdoDoc = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "filldocu"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Sub

Private Sub PrintReport()
   Dim sBook As String
   MouseCursor 13
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   FillDocuments
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("quafa01")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaName.Add "ShowDocs"
   aFormulaValue.Add CStr("'1'")
   
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
   sSql = "{FahdTable.FA_REF}='" & Compress(txtPrt) & "' " _
          & "AND {FahdTable.FA_REVISION}='" & txtRev & "'"
   cCRViewer.SetReportSelectionFormula (sSql)
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub PrintReport1()
   Dim sBook As String
   MouseCursor 13
   
   FillDocuments
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   sCustomReport = GetCustomReport("quafa01")
   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.Crw.Formulas(1) = "ShowDocs='1'"
   sSql = "{FahdTable.FA_REF}='" & Compress(txtPrt) & "' " _
          & "AND {FahdTable.FA_REVISION}='" & txtRev & "'"
   MdiSect.Crw.SelectionFormula = sSql
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillInspectors()
   On Error GoTo DiaErr1
   sSql = "Qry_FillInspectorsActive"
   LoadComboBox cmbIns, -1
   If cmbIns.ListCount > 0 Then cmbIns = cmbIns.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillinspe"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub cmbIns_Click()
   bGoodIns = GetInspector()
   
End Sub


Private Sub cmbIns_LostFocus()
   cmbIns = CheckLen(cmbIns, 12)
   cmbIns = Compress(cmbIns)
   bGoodIns = GetInspector()
   If bGoodReport Then
      On Error Resume Next
      With RdoRpt
         !FA_INSPECTOR = cmbIns
         .update
      End With
   End If
End Sub


Private Sub cmdCan_Click()
   UpdateDrawings
   
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6202
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdItm_Click()
   If bGoodReport Then
      On Error Resume Next
      With RdoRpt
         !FA_REVISED = Format(ES_SYSDATE, "mm/dd/yy")
         .update
      End With
   End If
   FainFAe01b.Show
   
End Sub

Private Sub cmdPrn_Click()
   If cmdPrn Then
      cmdPrn = False
      PrintReport
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillInspectors
      'FillFaDocuments
      bGoodReport = GetReport()
      bOnLoad = 0
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   tabFrame(0).BorderStyle = 1
   tabFrame(1).BorderStyle = 1
   tabFrame(1).Left = 180
   
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If optFrom.Value = vbChecked Then
      FainFAe01a.Show
   Else
      FainFAe02a.Show
   End If
   optFrom.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set RdoRpt = Nothing
   Set FainFAe02b = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtIns.ToolTipText = "Date Format. Leave Blank If Not Inspected."
   
End Sub


Private Function GetReport() As Byte
   Dim RdoDoc As ADODB.Recordset
   Dim b As Integer
   Dim sReptRef As String
   If optFrom.Value = vbChecked Then
      txtPrt = FainFAe01a.txtPrt
      txtRev = FainFAe01a.txtRev
      Unload FainFAe01a
   Else
      txtPrt = FainFAe02a.cmbPrt
      txtRev = FainFAe02a.cmbRev
      Unload FainFAe02a
   End If
   On Error GoTo DiaErr1
   sReptRef = Compress(txtPrt)
   sSql = "SELECT * FROM FahdTable WHERE FA_REF='" & sReptRef _
          & "' AND FA_REVISION='" & txtRev & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_KEYSET)
   If bSqlRows Then
      With RdoRpt
         txtDsc = "" & Trim(!FA_DESCRIPTION)
         txtBeg = Format(!FA_CREATED, "mm/dd/yy")
         cmbIns = "" & Trim(!FA_INSPECTOR)
         If Not IsNull(!FA_INSPECTED) Then
            txtIns = Format(!FA_INSPECTED, "mm/dd/yy")
         Else
            txtIns = ""
         End If
         optCom.Value = !FA_COMPLETE
         txtDwp = Format(!FA_DRAWINGTOLPLUS, ES_QuantityDataFormat)
         txtDwm = Format(!FA_DRAWINGTOLMINUS, ES_QuantityDataFormat)
         txtAnp = Format(!FA_ANGLETOLPLUS, "#0.0")
         txtAnm = Format(!FA_ANGLETOLMINUS, "#0.0")
         '10/4/06
         txtMon = "" & Trim(!FA_MORUNPART)
         txtRun = "" & Trim(!FA_MORUNNO)
         txtCmt = "" & Trim(!FA_COMMENTS)
         bGoodIns = GetInspector()
         GetReport = 1
      End With
   End If
   If GetReport = 1 Then
      Dim doc As String, sheet As String, rev As String, item As String
      sSql = "SELECT FA_DOCITEM, FA_DOCDESCRIPTION," _
             & "FA_DOCSHEET,FA_DOCCHANGE FROM FadcTable " _
             & "WHERE FA_DOCNUMBER='" & sReptRef & "' AND " _
             & "FA_DOCREVISION='" & txtRev & "' " _
             & "ORDER BY FA_DOCITEM"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_FORWARD)
      If bSqlRows Then
         With RdoDoc
            Do Until .EOF
               b = b + 1
               If b > 10 Then Exit Do
               item = "" & str(!FA_DOCITEM)
               doc = "" & Trim(!FA_DOCDESCRIPTION)
               sheet = "" & Trim(!FA_DOCSHEET)
               rev = "" & Trim(!FA_DOCCHANGE)
'               sDocuments(b, 0) = item
'               sDocuments(b, 1) = doc
'               sDocuments(b, 2) = sheet
'               sDocuments(b, 3) = rev
               
               PopulateDocCombo b, doc, doc
               
               ' set values in controls
               lblDoc(b) = item
               Dim S As String
               On Error Resume Next    ' error if mo matching document/sht/rev
                  'txtDno(b) = doc      ' @@@ ERROR 383 HAPPENS HERE
                  SetComboBox txtDno(b), Compress(doc)
                  If txtDsh(b) <> sheet Then    ' get error 383 if attempt to set blank
                     txtDsh(b) = sheet
                  Else
                     GetDocRevisions b       ' if sheet blank, have to fill rev list
                  End If
                  If txtDch(b) <> rev Then    ' get error 383 if attempt to set blank
                     txtDch(b) = rev
                  End If
                  If Err Then
                     If doc <> "" Then
                        S = "doc: " & doc & " sheet: " & sheet & " rev: " & rev & " does not exist."
                        MsgBox S
                     End If
                     Err.Clear
                  End If
               On Error GoTo DiaErr1

              ' add prefix and select all docs with that prefix
'               If sDocuments(b, 1) <> "" Then
'                  GetDocSheets CInt(b), sDocuments(b, 2)
'               End If
               
               .MoveNext
            Loop
         End With
         ClearResultSet RdoDoc
      End If
   End If
   Set RdoDoc = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub PopulateDocCombo(Index As Integer, pfx As String, Optional doc As String = "")
   ' populate combo with all items beginning with first three characters of prefix
   ' if doc is nonblank, choose that value
   Dim prefix As String
   prefix = Compress(pfx)
   If Len(prefix) > 3 Then
      prefix = Mid(prefix, 1, 3)
   End If
   txtPrefix(Index) = prefix
   txtDno(Index).Clear
   
   If prefix <> "" Then
'      sSql = "SELECT DISTINCT DOREF,DONUM,DOCLASS FROM DdocTable" & vbCrLf _
'         & "where DOREF like '" & prefix & "%' order by DOREF"
      sSql = "SELECT DONUM FROM DdocTable" & vbCrLf _
         & "where DOREF like '" & prefix & "%'" & vbCrLf _
         & "UNION" & vbCrLf _
         & "SELECT '' as DONUM" & vbCrLf _
         & "order by DONUM"
      Dim rdodcs As ADODB.Recordset
      bSqlRows = clsADOCon.GetDataSet(sSql, rdodcs, ES_FORWARD)
      If bSqlRows Then
         With rdodcs
            Do Until .EOF
               AddComboStr txtDno(Index).hwnd, "" & Trim(!DONUM)
               .MoveNext
            Loop
            ClearResultSet rdodcs
         End With
      End If
      
      ' set doc selection
      If doc <> "" Then
         'txtDno(Index) = doc
         
         ' populate sheet dropdown
         GetDocSheets Index
      End If
   End If
   
End Sub

Private Function GetInspector() As Byte
   Dim RdoIns As ADODB.Recordset
   On Error GoTo DiaErr1
   If Trim(cmbIns) = "" Then
      lblNme = ""
      Exit Function
   End If
   sSql = "SELECT INSID,INSFIRST,INSMIDD,INSLAST,INSSTAMP,INSDIVISION " _
          & "FROM RinsTable WHERE INSID='" & Compress(cmbIns) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoIns, ES_FORWARD)
   If bSqlRows Then
      With RdoIns
         lblNme = "" & Trim(!INSFIRST) _
                  & " " & Trim(!INSMIDD) _
                  & " " & Trim(!INSLAST)
      End With
      ClearResultSet RdoIns
      GetInspector = 1
   Else
      lblNme = "*** Inspector Wasn't Found ***"
      GetInspector = False
   End If
   Set RdoIns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getinspect"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub optCom_Click()
   If bGoodReport Then
      On Error Resume Next
      With RdoRpt
         !FA_COMPLETE = optCom.Value
         !FA_REVISED = Format(ES_SYSDATE, "mm/dd/yy")
         .update
      End With
   End If
   
End Sub

Private Sub optFrom_Click()
   'never visible.  Checked is from new else from revise
   
End Sub



Private Sub tab1_Click()
   On Error Resume Next
   If tab1.SelectedItem.Index = 1 Then
      tabFrame(0).Visible = True
      tabFrame(1).Visible = False
      txtDsc.SetFocus
   Else
      tabFrame(1).Visible = True
      tabFrame(0).Visible = False
      txtDno(1).SetFocus
   End If
   
End Sub


Private Sub txtAnm_LostFocus()
   txtAnm = CheckLen(txtAnm, 4)
   txtAnm = Format(Abs(Val(txtAnm)), "#0.0")
   If bGoodReport Then
      On Error Resume Next
      With RdoRpt
         !FA_ANGLETOLMINUS = Val(txtAnm)
         .update
      End With
   End If
   
End Sub


Private Sub txtAnp_LostFocus()
   txtAnp = CheckLen(txtAnp, 4)
   txtAnp = Format(Abs(Val(txtAnp)), "#0.0")
   If bGoodReport Then
      On Error Resume Next
      With RdoRpt
         !FA_ANGLETOLPLUS = Val(txtAnp)
         !FA_REVISED = Format(ES_SYSDATE, "mm/dd/yy")
         .update
      End With
   End If
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
   If bGoodReport Then
      On Error Resume Next
      With RdoRpt
         !FA_REVISED = Format(ES_SYSDATE, "mm/dd/yy")
         !FA_CREATED = Format(txtBeg, "mm/dd/yy")
         .update
      End With
   End If
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 1020)
   If bGoodReport Then
      On Error Resume Next
      With RdoRpt
         !FA_COMMENTS = txtCmt
         .update
      End With
   End If
   
End Sub


Private Sub txtDch_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtDch_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub

Private Sub txtDch_LostFocus(Index As Integer)
   'txtDch(Index) = CheckLen(txtDch(Index), 12)
   'sDocuments(Index, 3) = Trim(txtDch(Index))
   
End Sub

Private Sub txtDno_Change(Index As Integer)
   GetDocSheets Index
End Sub

Private Sub txtDno_Click(Index As Integer)
   GetDocSheets Index
End Sub

Private Sub txtDno_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtDno_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub txtDno_LostFocus(Index As Integer)
   'txtDno(Index) = CheckLen(txtDno(Index), 30)
   'sDocuments(Index, 1) = Trim(txtDno(Index))
   GetDocSheets Index
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   If bGoodReport Then
      On Error Resume Next
      With RdoRpt
         !FA_DESCRIPTION = txtDsc
         .update
      End With
   End If
   
End Sub


Private Sub txtDsh_Change(Index As Integer)
   GetDocRevisions Index
End Sub

Private Sub txtDsh_Click(Index As Integer)
   GetDocRevisions Index
End Sub

Private Sub txtDsh_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtDsh_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub txtDsh_LostFocus(Index As Integer)
   'txtDsh(Index) = CheckLen(txtDsh(Index), 6)
   'sDocuments(Index, 2) = Trim(txtDsh(Index))
   GetDocRevisions Index
End Sub


Private Sub txtDwm_LostFocus()
   txtDwm = CheckLen(txtDwm, 6)
   txtDwm = Format(Abs(Val(txtDwm)), ES_QuantityDataFormat)
   If bGoodReport Then
      On Error Resume Next
      With RdoRpt
         !FA_DRAWINGTOLMINUS = Val(txtDwm)
         .update
      End With
   End If
   
End Sub


Private Sub txtDwp_LostFocus()
   txtDwp = CheckLen(txtDwp, 6)
   txtDwp = Format(Abs(Val(txtDwp)), ES_QuantityDataFormat)
   If bGoodReport Then
      On Error Resume Next
      With RdoRpt
         !FA_DRAWINGTOLPLUS = Val(txtDwp)
         !FA_REVISED = Format(ES_SYSDATE, "mm/dd/yy")
         .update
      End With
   End If
   
End Sub


Private Sub txtIns_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtIns_LostFocus()
   If Len(Trim(txtIns)) Then
      txtIns = CheckDate(txtIns)
      On Error Resume Next
      If bGoodReport Then
         On Error Resume Next
         With RdoRpt
            !FA_INSPECTED = txtIns
            !FA_REVISED = Format(ES_SYSDATE, "mm/dd/yy")
            .update
         End With
      End If
   Else
      If bGoodReport Then
         On Error Resume Next
         With RdoRpt
            !FA_INSPECTED = Null
            !FA_REVISED = Format(ES_SYSDATE, "mm/dd/yy")
            .update
         End With
      End If
   End If
End Sub



Private Sub UpdateDrawings()
   Dim b As Byte
   Dim doc As String, sheet As String, rev As String, item As String
   Dim update As Boolean
   Dim SReportRef As String
   Dim OKtoUpdate
   SReportRef = Compress(txtPrt)
   MouseCursor 13
   On Error GoTo DiaErr1
   For b = 1 To 10
      update = True
      item = lblDoc(b)
      doc = Compress(txtDno(b))
      sheet = Compress(txtDsh(b))
      rev = Compress(txtDch(b))
   
'      If Trim(sDocuments(b, 1)) = "" Then
'         sDocuments(b, 2) = ""
'         sDocuments(b, 3) = ""
'      End If
      If doc = "" Then
         sheet = ""
         rev = ""
         update = True
      'do not write the record if it does not match a document
      Else
         Dim rdo As ADODB.Recordset
         sSql = "select DOREF, DOSHEET, DOREV FROM DdocTable" & vbCrLf _
            & "WHERE DOREF='" & doc & "' and DOSHEET='" & sheet & "' and  DOREV='" & rev & "'"
         bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
         Set rdo = Nothing
         If bSqlRows = 0 Then
            update = False
            MsgBox "doc: " & doc & " sheet: " & sheet & " rev: " & rev & " does not exist.  Document will not be updated. "
         End If
      End If
      If update Then
'      sSql = "UPDATE FadcTable SET " _
'             & "FA_DOCDESCRIPTION='" & sDocuments(b, 1) & "'," _
'             & "FA_DOCSHEET='" & sDocuments(b, 2) & "'," _
'             & "FA_DOCCHANGE='" & sDocuments(b, 3) & "' " _
'             & "WHERE (FA_DOCNUMBER='" & SReportRef & "' AND " _
'             & "FA_DOCREVISION='" & txtRev & "' AND " _
'             & "FA_DOCITEM=" & str(b) & ")"
         sSql = "UPDATE FadcTable SET " _
                & "FA_DOCDESCRIPTION='" & doc & "'," _
                & "FA_DOCSHEET='" & sheet & "'," _
                & "FA_DOCCHANGE='" & rev & "' " & vbCrLf _
                & "WHERE (FA_DOCNUMBER='" & SReportRef & "' AND " _
                & "FA_DOCREVISION='" & txtRev & "' AND " _
                & "FA_DOCITEM=" & str(b) & ")"
         clsADOCon.ExecuteSql sSql
      End If
   Next
   Unload Me
   Exit Sub
   
DiaErr1:
   sProcName = "updatedrawings"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Sub

Private Sub txtMon_LostFocus()
   txtMon = CheckLen(txtMon, 30)
   If bGoodReport Then
      On Error Resume Next
      With RdoRpt
         !FA_MORUNPART = txtMon
         !FA_REVISED = Format(ES_SYSDATE, "mm/dd/yy")
         .update
      End With
   End If
   
End Sub



Private Sub txtPrefix_GotFocus(Index As Integer)
   prefixValueAtGotFocus = txtPrefix(Index).Text
End Sub

Private Sub txtPrefix_LostFocus(Index As Integer)
   Dim newPrefixValue As String
   newPrefixValue = txtPrefix(Index).Text
   If newPrefixValue <> prefixValueAtGotFocus Then
      PopulateDocCombo Index, txtPrefix(Index).Text
   End If
End Sub

Private Sub txtRun_LostFocus()
   txtRun = CheckLen(txtRun, 6)
   txtRun = Format$(Abs(Val(txtRun)), "000000")
   If bGoodReport Then
      On Error Resume Next
      With RdoRpt
         !FA_MORUNNO = txtRun
         .update
      End With
   End If
   
End Sub



Private Sub FillFaDocuments()
   Dim rdodcs As ADODB.Recordset
   Dim iList As Integer
   On Error GoTo DiaErr1
'StopwatchStart
'   sSql = "SELECT DISTINCT DOREF,DONUM,DOCLASS FROM DdocTable  "
'   bSqlRows = clsADOCon.GetDataSet(sSql, rdodcs, ES_FORWARD)
'   If bSqlRows Then
'      With rdodcs
'         Do Until .EOF
'            For iList = 1 To 10
'               AddComboStr txtDno(iList).hwnd, "" & Trim(!DONUM)
'            Next
'            .MoveNext
'         Loop
'         ClearResultSet rdodcs
'      End With
'   End If
'StopwatchStop "FillFADocuments populate combo"

   If txtDno(1).ListCount > 0 Then
      For iList = 1 To 10
         GetDocRevisions iList
      Next
   End If
   Set rdodcs = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "FillFaDocu"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetDocRevisions(iIndex As Integer)
   Dim sOldRev As String
   sOldRev = txtDch(iIndex)
   
   txtDch(iIndex).Clear
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT RTRIM(DOREV) AS DOREV FROM DdocTable " & vbCrLf _
          & "WHERE DOREF='" & Compress(txtDno(iIndex)) & "' and DOSHEET='" & Compress(txtDsh(iIndex)) & "' " & vbCrLf _
          & "ORDER BY DOREV"
   LoadComboBox txtDch(iIndex), -1, False    ' select none
   On Error Resume Next
   If sOldRev <> "" Then
      txtDch(iIndex) = sOldRev
      If Err Then
   '      If doc <> "" Then
'            Dim s As String
'            s = "doc(rev): " & txtDno(iIndex) & " sheet: " & txtDsh(iIndex) & " rev: " & txtDch(iIndex) & " does not exist."
'            MsgBox s
   '      End If
         Err.Clear
      End If
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getrevisi"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
