VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ESI Point Of Manufacturing"
   ClientHeight    =   11595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13650
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11595
   ScaleWidth      =   13650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWCStat 
      Caption         =   "WC Status"
      Height          =   1000
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   9480
      Width           =   1500
   End
   Begin VB.CommandButton cmdMelterLog 
      Caption         =   "Melter Log"
      Height          =   1000
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   6960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdSetupTime 
      BackColor       =   &H000000FF&
      Caption         =   "SetupTime"
      Default         =   -1  'True
      Height          =   1000
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   6960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdCurrentOps 
      Caption         =   "Show Only Current Operations"
      Height          =   1000
      Left            =   2160
      Picture         =   "frmMain.frx":F172
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   6960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdOpenOps 
      Caption         =   "Show All Open Operations"
      Height          =   1000
      Left            =   3720
      Picture         =   "frmMain.frx":F73F
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   6960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "Minimize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
   End
   Begin VB.PictureBox picComplete 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   180
      ScaleHeight     =   5655
      ScaleWidth      =   6315
      TabIndex        =   66
      Top             =   5100
      Visible         =   0   'False
      Width           =   6315
      Begin VB.TextBox txtNotes 
         Height          =   1335
         Left            =   3000
         MultiLine       =   -1  'True
         TabIndex        =   90
         Top             =   4320
         Width           =   4575
      End
      Begin VB.CheckBox chkComplete 
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CommandButton cmdScrap 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         TabIndex        =   76
         Top             =   2400
         Width           =   2175
      End
      Begin VB.CommandButton cmdRej 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         TabIndex        =   75
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CommandButton cmdCom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         TabIndex        =   74
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Comments:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   91
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label z1 
         Caption         =   "Operation Completed:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1080
         TabIndex        =   87
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label lblScrap 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   73
         Tag             =   "8"
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label lblRej 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   72
         Tag             =   "8"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lblCom 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   71
         Tag             =   "8"
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblJob 
         Caption         =   "lblJob"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   70
         Top             =   0
         Width           =   6735
      End
      Begin VB.Label z1 
         Caption         =   "Quantity Scrap:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   69
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label z1 
         Caption         =   "Quantity Rejected:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   68
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label z1 
         Caption         =   "Quantity Completed:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   67
         Top             =   720
         Width           =   2895
      End
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   840
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "OK"
      Height          =   1000
      Left            =   8400
      Picture         =   "frmMain.frx":FD0B
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   6960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   1000
      Left            =   600
      Picture         =   "frmMain.frx":1034B
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   6960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picDummy 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1035
      TabIndex        =   61
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
      Begin VB.Label lblDummy 
         Caption         =   "lblDummy"
         Height          =   255
         Left            =   0
         TabIndex        =   63
         Tag             =   "0"
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdPgDn 
      Caption         =   "Scroll Down"
      Height          =   1000
      Left            =   6840
      Picture         =   "frmMain.frx":10920
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   6960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdPgUP 
      Caption         =   "Scroll Up"
      Height          =   1000
      Left            =   5280
      Picture         =   "frmMain.frx":10F05
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   6960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picJobs 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   180
      ScaleHeight     =   2145
      ScaleWidth      =   2385
      TabIndex        =   8
      Top             =   2760
      Width           =   2415
      Begin MSFlexGridLib.MSFlexGrid grdMO 
         Height          =   6375
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   11245
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   800
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblMOMsg 
         Caption         =   "lblMOMsg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   0
         Width           =   11055
      End
   End
   Begin VB.PictureBox picShops 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   10980
      ScaleHeight     =   1545
      ScaleWidth      =   2100
      TabIndex        =   7
      Top             =   240
      Width           =   2130
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   17
         Left            =   9360
         TabIndex        =   27
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   16
         Left            =   7560
         TabIndex        =   26
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   15
         Left            =   5760
         TabIndex        =   25
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   14
         Left            =   3960
         TabIndex        =   24
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   13
         Left            =   2160
         TabIndex        =   23
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   12
         Left            =   360
         TabIndex        =   22
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   11
         Left            =   9360
         TabIndex        =   21
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   10
         Left            =   7560
         TabIndex        =   20
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   9
         Left            =   5760
         TabIndex        =   19
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   8
         Left            =   3960
         TabIndex        =   18
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   7
         Left            =   2160
         TabIndex        =   17
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   6
         Left            =   360
         TabIndex        =   16
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   5
         Left            =   9360
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   4
         Left            =   7560
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   3
         Left            =   5760
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   2
         Left            =   3960
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   1
         Left            =   2160
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblShopMsg 
         Caption         =   "lblShopMsg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   0
         Width           =   11055
      End
   End
   Begin VB.PictureBox picWCS 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   3900
      ScaleHeight     =   3345
      ScaleWidth      =   6060
      TabIndex        =   6
      Top             =   4200
      Width           =   6090
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   23
         Left            =   9120
         TabIndex        =   82
         Top             =   4920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   22
         Left            =   7320
         TabIndex        =   81
         Top             =   4920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   21
         Left            =   5520
         TabIndex        =   80
         Top             =   4920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   20
         Left            =   3720
         TabIndex        =   79
         Top             =   4920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   19
         Left            =   1920
         TabIndex        =   78
         Top             =   4920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   18
         Left            =   120
         TabIndex        =   77
         Top             =   4920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   17
         Left            =   9120
         TabIndex        =   45
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   16
         Left            =   7320
         TabIndex        =   44
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   15
         Left            =   5520
         TabIndex        =   43
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   14
         Left            =   3720
         TabIndex        =   42
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   13
         Left            =   1920
         TabIndex        =   41
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   12
         Left            =   120
         TabIndex        =   40
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   11
         Left            =   9120
         TabIndex        =   39
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   10
         Left            =   7320
         TabIndex        =   38
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   9
         Left            =   5520
         TabIndex        =   37
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   8
         Left            =   3720
         TabIndex        =   36
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   7
         Left            =   1920
         TabIndex        =   35
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   6
         Left            =   120
         TabIndex        =   34
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   5
         Left            =   9120
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   4
         Left            =   7320
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   3
         Left            =   5520
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   2
         Left            =   3720
         TabIndex        =   30
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   1
         Left            =   1920
         TabIndex        =   29
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdWC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblWCMsg 
         Caption         =   "lblWCMsg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   0
         Width           =   10575
      End
   End
   Begin VB.PictureBox picLogin 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   2040
      ScaleHeight     =   4665
      ScaleWidth      =   9705
      TabIndex        =   0
      Top             =   2040
      Width           =   9735
      Begin VB.CommandButton cmdTimesheet 
         Caption         =   "Timesheet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   7800
         TabIndex        =   55
         Top             =   1920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chkKeyboard 
         Caption         =   "Use Keyboard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5340
         TabIndex        =   57
         Top             =   4260
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdChangeEmployee 
         Caption         =   "Change Employee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   7800
         TabIndex        =   54
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdPIN 
         Caption         =   "PIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   5580
         TabIndex        =   51
         Top             =   0
         Width           =   1935
      End
      Begin VB.CommandButton cmdVersion 
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   7800
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdPunchOut 
         Caption         =   "Punch Out"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   5580
         TabIndex        =   53
         Top             =   1920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "Log Off Job"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   5580
         TabIndex        =   52
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Log On Job"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   5580
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblBulBrd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Bulletin Board"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   88
         Top             =   4680
         Width           =   10935
      End
      Begin VB.Label lblTme 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblTme"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3060
         TabIndex        =   49
         Top             =   120
         Width           =   2070
      End
      Begin VB.Label lblMsg 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2250
         Left            =   0
         TabIndex        =   4
         Top             =   4920
         Width           =   10935
      End
      Begin VB.Label lblEmp 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   0
         TabIndex        =   3
         Top             =   600
         Width           =   5235
      End
      Begin VB.Label lblDte 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   60
         Width           =   5235
      End
      Begin VB.Label lblPIN 
         Caption         =   "PIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5580
         TabIndex        =   1
         Tag             =   "6"
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Timer tmr1 
      Interval        =   15000
      Left            =   120
      Top             =   360
   End
   Begin VB.Label lblLotQty 
      Height          =   495
      Left            =   3960
      TabIndex        =   89
      Tag             =   "8"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image imgBuf 
      Height          =   375
      Left            =   3120
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblPickQty 
      Height          =   495
      Left            =   1440
      TabIndex        =   85
      Tag             =   "8"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblSystemMsg 
      Caption         =   "lblSystemMsg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   6975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) ESI 2003
'
'**********************************************************************************
'
' frmMain - Handles all the UI and controls
'
' Notes:
'
' Created: (nth) 03/12/03
' Revisions:
'   05/15/03 (nth) Add picklist for pick operations.
'   06/06/03 (nth) Added pictures to grids per SJW.
'   06/16/03 (nth) Fixed runtime error invalid use of NULL in OPSUDATE and OPRUNDATE
'   07/09/03 (nth) Make frm12key minimize
'   07/25/03 (nth) Disabled multiple selection when logging off job.
'   12/19/03 (JCW) Add Time Charges Journal Logic... Must record in TcitTable when logged off!
'   12/22/03 (JCW) Add Overhead Rate Calculation Function; Also: Misc. Bug Fixes
'   03/04/03 (JCW) Fixed UpdateTables \ fixed completeOp(no current row bug) \ timing problem
'   06/15/04 (nth) Changed to a left(opcomt,255) to prevent invalid cursor position
'                  error in fillmos because of increase in opcomt to 5120.
'
'**********************************************************************************

Option Explicit

Private bSetupTime As Boolean
Private bMOComt As Boolean

Dim mjobComplete As Job

Dim mblnGoodEmpl As Boolean
Dim mblnLocked As Boolean
Dim mblnOnload As Boolean
Dim mblnRevise As Boolean
Dim mblnPick As Boolean
Dim mblnShow12Key As Boolean

Dim mlngGridNormal As Long

Dim mintScrollMode As Integer
Dim mintGridY As Integer ' Current grid row
Dim denyLoginIfPriorOpOpen As Boolean

Dim mstrLotPart As String

'Dim mempCurrentEmployee  As Employee


'**********************************************************************************


Private Function PadNumber(sIn As String, padlen As Integer)
    PadNumber = Right(String(padlen, "0") & Trim(sIn), padlen)
End Function

Private Sub chkComplete_Click()
   With chkComplete
      If .Value = vbChecked Then
         .Caption = "Yes"
      Else
         .Caption = "No"
      End If
   End With
End Sub

Private Sub chkComplete_MouseDown(Button As Integer, _
                                  Shift As Integer, x As Single, y As Single)
   frm12Key.Hide
End Sub

Private Sub chkKeyboard_Click()
   SaveSetting "Esi2000", "EsiPOM", "UseKeyboard", CStr(chkKeyboard.Value)
End Sub

Private Sub cmdBack_Click()
   Select Case gbytScreen
      Case LOGIN
         Beep
      Case SHOPS
         ShowLogin True
      Case WCS
         ShowShops
      Case jobs
         If mblnRevise Then
            ShowLogin True
         Else
            ShowWcs
         End If
      Case complete
         '*
         '* This will need to change (nth)
         '*
         mblnRevise = False
         frm12Key.Hide
         ShowLogin True
      Case PKLIST
         '*
         '* This will need to change (nth)
         '*
         mblnPick = False
         frm12Key.Hide
         ShowLogin True
      Case Lots
         frm12Key.Hide
         ShowPkList
   End Select
End Sub

Private Sub cmdCom_Click()
   cmdScrap.Visible = True
   cmdScrap.Caption = lblScrap
   
   cmdRej.Visible = True
   cmdRej.Caption = lblRej
   
   cmdCom.Visible = False
   Activate_Label lblCom, True, True
End Sub

Private Sub cmdCurrentOps_Click()
   cmdCurrentOps.Enabled = cmdOpenOps.Enabled
   cmdOpenOps.Enabled = Not cmdCurrentOps.Enabled
   SaveSetting "Esi2000", "EsiPOM", "ShowOpenOps", CStr(cmdCurrentOps.Enabled)
   ShowJobs
End Sub

Private Sub cmdExit_Click()
   Unload Me
   End
End Sub

Private Sub cmdChangeEmployee_Click()
   ' Log employee off system
   mblnLocked = False
   glblActive = lblDummy
   ShowLogin
End Sub

Private Sub cmdMin_Click()
   With frmMain
      WindowState = 1
   End With
   
   If frm12Key.Visible = True Then
      mblnShow12Key = True
      frm12Key.Visible = False
   End If
End Sub

Private Sub cmdNew_Click()
   Dim CurrentLogins As Integer
   CurrentLogins = UBound(mempCurrentEmployee.jobCurMO) + 1
   If mempCurrentEmployee.jobCurMO(0).lngRun = 0 Then
      CurrentLogins = 0
   Else
      CurrentLogins = UBound(mempCurrentEmployee.jobCurMO) + 1
   End If
   
   
   If CurrentLogins >= MAX_CONCURRENT_LOGINS Then
      MsgBox "You are currently logged on to " & CurrentLogins & " jobs." & vbCrLf _
         & "The maximum simultaneous logins is " & MAX_CONCURRENT_LOGINS & vbCrLf _
         & "You must log off of " & CurrentLogins - MAX_CONCURRENT_LOGINS + 1 _
         & " jobs before logging on to another job.", vbInformation
      Exit Sub
   End If
   mblnRevise = False
   If chkKeyboard.Value = 0 Then
      ShowShops
   Else
      frmKeyInJob.Show vbModal
      ShowLogin True
   End If
End Sub

Private Sub cmdMelterLog_Click()
   MelterLog.txtMOPartNum = mjobComplete.strPart
   MelterLog.txtMORun = mjobComplete.lngRun
   MelterLog.txtMelNum = mempCurrentEmployee.intNumber
   MelterLog.txtGCast = lblCom
   MelterLog.txtRejQty = lblRej
   MelterLog.txtNotes = txtNotes
   MelterLog.Show vbModal
End Sub

Private Sub cmdOff_Click()
   ' Complete / log off a current job.
   cmdOff.Enabled = False
   chkComplete.Value = vbUnchecked
   mblnRevise = True
   ShowJobs
   cmdOff.Enabled = True
End Sub

Private Sub cmdOpenOps_Click()
   cmdCurrentOps_Click
End Sub

Private Sub cmdPgDn_MouseDown(Button As Integer, _
                              Shift As Integer, x As Single, y As Single)
   If grdMO.TopRow + 1 < grdMO.Rows Then
      tmrScroll.Enabled = True
      mintScrollMode = 1
   End If
End Sub

Private Sub cmdPgDn_MouseUp(Button As Integer, _
                            Shift As Integer, x As Single, y As Single)
   tmrScroll.Enabled = False
   mintScrollMode = 0
End Sub

Private Sub cmdPgUP_MouseDown(Button As Integer, _
                              Shift As Integer, x As Single, y As Single)
   tmrScroll.Enabled = True
   mintScrollMode = -1
End Sub

Private Sub cmdPgUP_MouseUp(Button As Integer, _
                            Shift As Integer, x As Single, y As Single)
   tmrScroll.Enabled = False
   mintScrollMode = 0
End Sub

Private Sub cmdPIN_Click()
   If Not mblnLocked Then
      Activate_Label lblPIN, True, True
      cmdPIN.Visible = False
   End If
End Sub

Private Sub cmdProceed_Click()
   cmdProceed.Enabled = False
   Select Case gbytScreen
      Case jobs
         If mblnRevise Then
            If IsPickOp(mjobComplete) Then
               ShowComplete
            Else
               ShowComplete
            End If
         Else
            LogOnToJobs grdMO
            ShowLogin True
         End If
      Case PKLIST
         ShowComplete
         
      Case complete
         frm12Key.Hide
         
         If lblCom.Caption = "" Then lblCom.Caption = "0"
         If lblRej.Caption = "" Then lblRej.Caption = "0"
         If lblScrap.Caption = "" Then lblScrap.Caption = "0"
         
         If mblnPick Then
            PickItems grdMO, mjobComplete
            mblnPick = False
         End If
         
         'gstrJournal = GetOpenJournal(Format(GetServerDateTime, "mm/dd/yy"))
         gstrJournal = GetOpenJournal("TJ", GetServerDateTime)
         
'        If GetOpenTimeJournalForThisDate(GetServerDateTime, gstrJournal) Then
         
         If Trim(gstrJournal) <> "" Then
            gsngOverHead = GetOverheadRate(mempCurrentEmployee.intNumber, mjobComplete.strPart, mjobComplete.lngRun, mjobComplete.intOp, mempCurrentEmployee.sngRate)
            LogOffJob _
               mempCurrentEmployee, _
               mjobComplete, _
               CSng(lblCom.Caption), _
               CSng(lblRej.Caption), _
               CSng(lblScrap.Caption), _
               chkComplete.Value, _
               True, txtNotes.Text, bMOComt
               
            txtNotes.Text = ""
            ShowLogin True
         Else
            SystemAlert SYSMSG13
         End If
   End Select
   cmdProceed.Enabled = True
End Sub

Private Sub cmdProceed_MouseDown(Button As Integer, _
                                 Shift As Integer, x As Single, y As Single)
   frm12Key.Hide
End Sub

Private Sub cmdPunchOut_Click()
   cmdPunchOut.Enabled = False      ' don't allow it to be clicked more than once
   If PunchOut(mempCurrentEmployee) Then
      SystemAlert "Punching Out." & vbCrLf & vbCrLf _
         & mempCurrentEmployee.strFirstName & " " _
         & mempCurrentEmployee.strLastName, _
         Me.Caption, True
      cmdChangeEmployee_Click
   End If
   cmdPunchOut.Enabled = True
End Sub

Private Sub cmdRej_Click()
   cmdCom.Visible = True
   cmdCom.Caption = lblCom
   
   cmdScrap.Visible = True
   cmdScrap.Caption = lblScrap
   
   cmdRej.Visible = False
   Activate_Label lblRej, True, True
End Sub

Private Sub cmdScrap_Click()
   cmdCom.Visible = True
   cmdCom.Caption = lblCom
   
   cmdRej.Visible = True
   cmdRej.Caption = lblRej
   
   cmdScrap.Visible = False
   Activate_Label lblScrap, True, True
End Sub

Private Sub cmdSetupTime_Click()
   If (bSetupTime = True) Then
      bSetupTime = False
      cmdSetupTime.BackColor = &HFF&
   Else
      bSetupTime = True
      cmdSetupTime.BackColor = &HFF00&
   End If
   
End Sub



Private Sub cmdTimesheet_Click()
   frmTimecard.lblEmployeeNo = mempCurrentEmployee.intNumber
   frmTimecard.lblEmployeeName = mempCurrentEmployee.strFirstName & " " & mempCurrentEmployee.strLastName
   frmTimecard.Show vbModal
End Sub

Private Sub cmdVersion_Click()
   
   'get creation date for program
   Dim sExeName As String
   Dim sFullPath As String
   Dim vDate As Variant
   Dim sDate As String
   
   sExeName = App.EXEName
   Me.Caption = "About " & sExeName
   sFullPath = App.Path & "\" & sExeName & ".exe"
   vDate = FileDateTime(sFullPath)
   If vDate <> "" Then
      sDate = Format(vDate, "mm/dd/yyyy")
   End If
   
   MsgBox "Point Of Manufacturing Version " & App.Major & "." & App.Minor _
      & "." & App.Revision & " - " & sDate & Chr(13) & Chr(10) _
      & "Copyright Enterprise Systems Inc, 2007"
End Sub

Private Sub cmdWCStat_Click()
   
   Dim strwc As String
   Dim strshp As String
   
   GetWCShop mjobComplete.strPart, mjobComplete.lngRun, mjobComplete.intOp, strshp, strwc
   
   frmWCState.strPart = mjobComplete.strPart
   frmWCState.strRun = mjobComplete.lngRun
   frmWCState.strOp = mjobComplete.intOp
   
   frmWCState.cboShop = strshp
   frmWCState.cboWorkCenter = strwc
   
   
   frmWCState.Show vbModal

End Sub

Private Sub Form_Activate()
   If mblnOnload Then
      
      'Me.Caption = Me.Caption & " - " & sDataBase
      Caption = GetSystemCaption
      CenterPictureBox picLogin, Me
      CenterPictureBox picShops, Me
      CenterPictureBox picWCS, Me
      CenterPictureBox picJobs, Me
      CenterPictureBox picComplete, Me
      
      lblTme = Format(GetServerDateTime, "h:nn AM/PM")
      
      lblSystemMsg.Left = picLogin.Left
      lblSystemMsg.Top = picLogin.Top + picLogin.Height
      lblSystemMsg.Width = picLogin.Width
      'lblSystemMsg.Visible = True
      
      ES_SYSDATE = GetServerDateTime()
      cmdMelterLog.Left = (picLogin.Left + picLogin.Width) - (cmdMelterLog.Width)
      cmdMelterLog.Top = picLogin.Top + picLogin.Height

      cmdWCStat.Left = (picLogin.Left + picLogin.Width) - (cmdMelterLog.Width)
      cmdWCStat.Top = picLogin.Top + picLogin.Height
      
      cmdSetupTime.Left = (picLogin.Left + picLogin.Width) - (cmdSetupTime.Width * 2)
      cmdSetupTime.Top = picLogin.Top + picLogin.Height
      
      cmdProceed.Left = (picLogin.Left + picLogin.Width) - (cmdPgUP.Width * 3)
      cmdProceed.Top = picLogin.Top + picLogin.Height
      
      cmdPgUP.Left = (picLogin.Left + picLogin.Width) - (cmdPgUP.Width * 4)
      cmdPgUP.Top = picLogin.Top + picLogin.Height
      
      cmdPgDn.Left = (picLogin.Left + picLogin.Width) - (cmdPgUP.Width * 5)
      cmdPgDn.Top = picLogin.Top + picLogin.Height
      
      cmdOpenOps.Left = (picLogin.Left + picLogin.Width) - (cmdPgUP.Width * 6)
      cmdOpenOps.Top = picLogin.Top + picLogin.Height
      
      cmdCurrentOps.Left = (picLogin.Left + picLogin.Width) - (cmdPgUP.Width * 7)
      cmdCurrentOps.Top = picLogin.Top + picLogin.Height
      
      'cmdBack.Left = picLogin.Left
      cmdBack.Left = (picLogin.Left + picLogin.Width) - (cmdPgUP.Width * 8)
      cmdBack.Top = picLogin.Top + picLogin.Height
      
      bSetupTime = False
      bMOComt = MOComtEnabled
      
      With cmdExit
         .Left = Me.Width - .Width - 200
         .Top = Me.Top + 200
      End With
      
      With cmdMin
         .Left = (Me.Width - (.Width * 2)) - 200
         .Top = Me.Top + 200
      End With
      
      gstrCaption = Me.Caption
      
      ShowLogin
      
      'Check for Journals
      'gstrJournal = GetOpenJournal(Format(GetServerDateTime, "mm/dd/yy"))
      
'      gstrJournal = GetOpenJournal("TJ", GetServerDateTime)
'      If Trim(gstrJournal) = "" Then
'         SystemAlert SYSMSG13
'         Unload Me
'      End If
      Dim tc As New ClassTimeCharge
      If Not tc.GetOpenTimeJournalForThisDate(GetServerDateTime, gstrJournal) Then
         Unload Me
      End If
      
      lblMsg = GetSystemMessage
      
      mblnOnload = False
   End If
   
   'get deny login flag
    Dim rs As ADODB.Recordset
    Dim deny As Integer
    
    sSql = "SELECT isnull(DenyLoginIfPriorOpOpen,0) as DenyLoginIfPriorOpOpen FROM ComnTable WHERE COREF=1"
    bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
    If bSqlRows Then
        denyLoginIfPriorOpOpen = IIf(rs!denyLoginIfPriorOpOpen = 1, True, False)
    End If
    Set rs = Nothing
   
End Sub



Private Sub Form_Load()
   Dim pintHeight As Integer
   Dim pintWidth As Integer
  
   
   pintHeight = 7100
   pintWidth = 11000
   
   With frmMain
      .WindowState = 2
   End With
   chkKeyboard = GetSetting("Esi2000", "EsiPOM", "UseKeyboard", "0")
   
   With picLogin
      .Height = pintHeight
      .Width = pintWidth
      .Visible = False
      .BorderStyle = 0
   End With
   
   With picShops
      .Height = pintHeight
      .Width = pintWidth
      .Visible = False
      .BorderStyle = 0
   End With
   
   With picWCS
      .Height = pintHeight
      .Width = pintWidth
      .Visible = False
      .BorderStyle = 0
   End With
   
   With picJobs
      .Height = pintHeight
      .Width = pintWidth
      .Visible = False
      .BorderStyle = 0
   End With
   
   With picComplete
      .Height = pintHeight
      .Width = pintWidth
      .Visible = False
      .BorderStyle = 0
   End With
   
   ReDim mempCurrentEmployee.jobCurMO(0)
   mblnOnload = True
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
   Unload frm12Key
   
   Set clsADOCon = Nothing
   'Set RdoCon = Nothing
   Unload Me
   End
End Sub

Private Sub Form_Resize()
   If mblnShow12Key = True Then
      mblnShow12Key = False
      frm12Key.Visible = True
   End If
End Sub

Private Sub grdMO_Click()
   Dim psngQty As Single
   Dim MouseRow As Integer, MouseCol As Integer
   Dim FakeSort As Integer
   
   With grdMO
      MouseCol = .MouseCol
      MouseRow = .MouseRow
      
      .Row = .RowSel
      mintGridY = .RowSel
      .Col = .ColSel
      Select Case gbytScreen
         
         Case PKLIST
            Highlight_Grid_Item True, grdMO
            Sleep 500
            .Col = 2
            psngQty = CSng(.Text)
            .Col = 0
            mstrLotPart = Trim(.Text)
            ShowLots Trim(.Text), psngQty
            
         Case Lots
            Highlight_Grid_Item True, grdMO
            lblLotQty.Left = grdMO.Left + 2500
            lblLotQty.Top = grdMO.Top + 1600
            lblLotQty.Width = 1500
            Activate_Label lblLotQty, True, True, True
            
         Case jobs
            'if clicking on heading, sort by that row
            If MouseRow = 0 Then
                Select Case MouseCol
                Case 4: FakeSort = 10
                Case 1, 2, 3: FakeSort = MouseCol + 10
                Case 7: FakeSort = 14
                Case Else
                    FakeSort = 0
                End Select
                
'                If MouseCol = 4 Then FakeSort = 10 Else FakeSort = 0
               SortFlex grdMO, MouseCol, 3, FakeSort
               Exit Sub
            End If
            
            If .Rows > 1 Then
               
               'if operation is red, cannot proceed
               .Col = 3
               If .CellForeColor = ES_RED Then
                  MsgBox "You cannot log into this operation until the prior operation is complete"
                  Exit Sub
               End If
               .Col = .ColSel
               
               If .CellBackColor = YELLOW Then
                  Highlight_Grid_Item False, grdMO
               Else
                  .Col = 0
                  mjobComplete.strPart = Compress(.Text)
                  .Col = 1
                  mjobComplete.lngRun = CLng(.Text)
                  .Col = 2
                  mjobComplete.sngQty = CSng(.Text)
                  .Col = 3
                  mjobComplete.intOp = CInt(.Text)
                  Highlight_Grid_Item True, grdMO
                  
                  If (bMOComt = True) Then
                    .Col = 5
                    Dim strCom As String
                    
                    strCom = GetMOComment(mjobComplete.strPart, mjobComplete.lngRun)
                    txtNotes = strCom
                  End If
                  
                  
                  Sleep 500
                  cmdProceed_Click ' we'll just press the ok button for the user
               End If

'not required - OK button has already been pressed
'               If Something_Is_Selected(grdMO) Then
'                  cmdProceed.Visible = True
'               Else
'                  cmdProceed.Visible = False
'               End If
            End If
      End Select
      
   End With
End Sub

Private Function GetMOComment(PartNumber As String, RunNumber As Long) As String
   
   Dim Comments As String
   Dim rdo As ADODB.Recordset
   
   sSql = "SELECT RUNCOMMENTS FROM RunsTable WHERE RUNREF = '" & Compress(PartNumber) & "' " & vbCrLf _
         & " AND RUNNO = " & RunNumber
         
   If clsADOCon.GetDataSet(sSql, rdo) Then
      GetMOComment = rdo!RUNCOMMENTS
   Else
      GetMOComment = ""
   End If
   Set rdo = Nothing

            
End Function


Private Sub lblLotQty_Change()
   Dim pintI As Byte
   Dim psngQty As Single
   With grdMO
      .Row = mintGridY '.RowSel
      .Col = 4
      .Text = lblLotQty
      .Row = 1
      For pintI = 1 To .Rows - 1
         .Row = pintI
         If .CellBackColor = YELLOW Then
            If IsNumeric(.Text) Then
               psngQty = psngQty + CSng(.Text)
            End If
         End If
      Next
      
      lblMOMsg = SYSMSG12 & mstrLotPart & "  Qty. Required " & _
                 Format(lblPickQty, "0.000") & " \ Qty. Select " & _
                 Format(psngQty, "0.000")
      
      If psngQty = CSng(lblPickQty) Then
         cmdProceed.Visible = True
      Else
         cmdProceed.Visible = False
      End If
   End With
End Sub

Private Sub lblPickQty_Change()
   With grdMO
      .Row = mintGridY '.RowSel
      .Col = 4
      .Text = lblPickQty
      If gintResponse = vbYes Then
         .Col = 5
         .Text = "Yes"
         gintResponse = 0
      ElseIf gintResponse = vbNo Then
         .Col = 5
         .Text = "No"
         gintResponse = 0
      End If
   End With
End Sub


Private Sub tmr1_Timer()
'   Dim pstrMsg As String
   Dim pdteNow As Date
   
   pdteNow = GetServerDateTime()
   
   lblDte = Format(pdteNow, "ddd mmm dd yyyy")
   lblTme = Format(pdteNow, "h:nn AM/PM")
   
'   pstrMsg = GetSystemMessage
'   If pstrMsg <> "" Then
'      lblMsg = pstrMsg
'   End If
   lblMsg = GetSystemMessage
End Sub

Public Sub ProcessEnter()

   cmdProceed.Visible = False
   
   Select Case glblActive.Name
      Case "lblClk"
         Activate_Label lblPIN, True, True
      Case "lblPIN"
         Activate_Label lblDummy, True
         mblnGoodEmpl = GetEmployee(Val(lblPIN), mempCurrentEmployee)
         
         If mblnGoodEmpl = True Then
            frm12Key.Hide
            ShowLogin True
         Else
             If bInvEmpFlag Then
                SystemAlert SYSMSG14
                bInvEmpFlag = False
             Else
                SystemAlert SYSMSG1
             End If
           Activate_Label lblPIN, True, True
         End If
      
      Case "lblCom"
         cmdProceed.Visible = True
         cmdRej_Click
      Case "lblRej"
         cmdProceed.Visible = True
         cmdScrap_Click
      Case "lblScrap"
         frm12Key.Hide
         cmdScrap.Caption = lblScrap
         cmdScrap.Visible = True
         cmdProceed.Visible = True
         Activate_Label lblDummy, True, False
         'cmdProceed.SetFocus
      Case "lblPickQty"
         frm12Key.Hide
      Case "lblLotQty"
         With grdMO
            .Row = mintGridY
            .Col = 4
            .Text = Format(.Text, "0.000")
            If CSng(.Text) = 0 Then
               Highlight_Grid_Item False, grdMO
            End If
         End With
         frm12Key.Hide
      Case Else
         frm12Key.EnterOn False
         mblnLocked = False
         MsgBox SYSMSG2
   End Select
End Sub

Public Sub Activate_Label( _
                          plblToActivate As Label, _
                          pblnActive As Boolean, _
                          Optional pblnShow12Key As Boolean, _
                          Optional pblnDelay As Boolean)
   
   Dim plngTopWindow As Long
   
   mblnLocked = True
   glblActive.BorderStyle = 0
   glblActive.BackColor = lblEmp.BackColor
   
   If pblnActive Then
      plblToActivate.Caption = ""
      plblToActivate.BorderStyle = 1
      plblToActivate.BackColor = lblMsg.BackColor
      plblToActivate.Visible = True
      Set glblActive = plblToActivate
   End If
   
   If pblnShow12Key Then
      frm12Key.Left = picLogin.Left + glblActive.Left + glblActive.Width
      frm12Key.Top = picLogin.Top + glblActive.Top
      
      If pblnDelay Then
         frm12Key.Show vbModal, Me
      Else
         frm12Key.Show
      End If
      plngTopWindow = SetTopMostWindow(frm12Key.hwnd, True)
   End If
End Sub

Private Sub SetSelectionPageControlsVisibility(bVisible As Boolean)
   cmdPgUP.Visible = bVisible
   cmdPgDn.Visible = bVisible
   
   ' get the flag from system setting
   cmdSetupTime.Visible = IIf((GetSetupTimeEnabled = 1), bVisible, False)
      
   
   If mblnRevise Or Not bVisible Then
      cmdCurrentOps.Visible = False
      cmdOpenOps.Visible = False
   Else
      cmdCurrentOps.Visible = True
      cmdOpenOps.Visible = True
      cmdCurrentOps.Enabled = CBool(GetSetting("Esi2000", "EsiPOM", "ShowOpenOps", "True"))
      cmdOpenOps.Enabled = Not cmdCurrentOps.Enabled
   End If
End Sub


Public Sub ShowLogin(Optional pblnLoggedIn As Boolean)
   Dim plngTopWindow As Long
   Dim pintI As Integer
   Dim pstrTemp As String
   
   gbytScreen = LOGIN
   picLogin.Visible = True
   picShops.Visible = False
   picWCS.Visible = False
   picJobs.Visible = False
   picComplete.Visible = False
   
   '    cmdPgUP.Visible = False
   '    cmdPgDn.Visible = False
   SetSelectionPageControlsVisibility False
   cmdBack.Visible = False
   cmdProceed.Visible = False
   cmdMelterLog.Visible = False
   cmdWCStat.Visible = False
   
   
   ' Check for jobs
   If pblnLoggedIn Then
      pstrTemp = "Good " & Time_Of_Day & " " _
         & mempCurrentEmployee.strFirstName & " " & mempCurrentEmployee.strLastName _
         & " (" & mempCurrentEmployee.intNumber & ")" & vbCrLf & vbCrLf
      
      GetCurrentJobs _
         mempCurrentEmployee.intNumber, _
         mempCurrentEmployee.jobCurMO()
      
      ' Check if no jobs found (lngRun = 0)
      If mempCurrentEmployee.jobCurMO(0).lngRun <> 0 Then
         
         pstrTemp = pstrTemp _
            & "You Are Currently Logged On To:" & vbCrLf _
            & ShowCurrentLogins(mempCurrentEmployee.intNumber)
         
         cmdOff.Enabled = True
      
      'indirect charge
      Else
'         pstrTemp = pstrTemp _
'                    & "You Are Not Currently Logged On To Any Jobs."
         
         pstrTemp = pstrTemp _
            & "You Are Currently Logged On To:" & vbCrLf _
            & ShowCurrentLogins(mempCurrentEmployee.intNumber)
         
         cmdOff.Enabled = False
         OpenIndirectTC mempCurrentEmployee
      End If
      
      lblEmp = pstrTemp
      
      lblPIN.Visible = False
      cmdPIN.Visible = False
      
      cmdNew.Visible = True
      chkKeyboard.Visible = True
      cmdChangeEmployee.Visible = True
      cmdTimesheet.Visible = True
      cmdVersion.Visible = True
      cmdPunchOut.Visible = True
      cmdOff.Visible = True
   Else
      ReDim mempCurrentEmployee.jobCurMO(0)
      mblnGoodEmpl = False
      
      lblEmp = SYSMSG0
      cmdPIN.Visible = True
      cmdNew.Visible = False
      chkKeyboard.Visible = False
      cmdChangeEmployee.Visible = False
      cmdTimesheet.Visible = False
      cmdPunchOut.Visible = False
      cmdOff.Visible = False
   End If
   
   lblDte = Format(GetServerDateTime(), "ddd mmm dd yyyy")
   
   frm12Key.Left = picLogin.Left + lblPIN.Left + lblPIN.Width + 100
   frm12Key.Top = picLogin.Top + 100
   
   'frm12Key.Show
   'plngTopWindow = SetTopMostWindow(frm12Key.hwnd, True)
   Set glblActive = lblDummy
End Sub

Public Sub ShowShops()
   Dim prdoShops As ADODB.Recordset
   Dim pintI As Integer
   
   gbytScreen = SHOPS
   Unload frm12Key
   
   picLogin.Visible = False
   picShops.Visible = True
   picWCS.Visible = False
   picJobs.Visible = False
   picComplete.Visible = False
   
   'cmdPgUP.Visible = False
   'cmdPgDn.Visible = False
   SetSelectionPageControlsVisibility False
   cmdBack.Visible = True
   cmdProceed.Visible = False
   
   lblShopMsg = "Good " & Time_Of_Day & " " _
                & mempCurrentEmployee.strFirstName & " " _
                & mempCurrentEmployee.strLastName _
                & ".  Please Select A Shop."
   sSql = "SELECT SHPNUM FROM ShopTable"
   gblnSqlRows = clsADOCon.GetDataSet(sSql, prdoShops)
   If gblnSqlRows Then
      With prdoShops
         Do While Not .EOF
            If pintI < cmdShop.UBound Then
               cmdShop(pintI).Caption = "" & Trim(!SHPNUM)
               If cmdShop(pintI).Caption <> "" Then
                  cmdShop(pintI).Visible = True
               End If
            Else
               MsgBox "Only " & cmdShop.UBound & " shops may be shown at once.", vbInformation
               Exit Do
            End If
            .MoveNext
            pintI = pintI + 1
         Loop
      End With
   End If
   Set prdoShops = Nothing
   
   'add "ALL"
   cmdShop(pintI).Caption = "ALL"
   cmdShop(pintI).Visible = True
   
End Sub

Public Sub ShowWcs()
   Dim prdoWCS As ADODB.Recordset
   Dim pintI As Integer
   
   gbytScreen = WCS
   picLogin.Visible = False
   picShops.Visible = False
   picWCS.Visible = True
   picJobs.Visible = False
   picComplete.Visible = False
   
   'cmdPgUP.Visible = False
   'cmdPgDn.Visible = False
   SetSelectionPageControlsVisibility False
   cmdBack.Visible = True
   cmdProceed.Visible = False
   
   lblWCMsg = SYSMSG3 & mempCurrentEmployee.strCurShop
   
   sSql = "SELECT DISTINCT WCNNUM FROM WcntTable "
   If mempCurrentEmployee.strCurShop <> "ALL" Then
      sSql = sSql & "WHERE WCNSHOP = '" & Compress(mempCurrentEmployee.strCurShop) & "'"
   End If
   gblnSqlRows = clsADOCon.GetDataSet(sSql, prdoWCS)
   
   ' 4/30/07: Make all WC's invisible and make invividual WC's visible below as
   ' they are initialized
   Dim I As Integer
   For I = 0 To cmdWC.UBound - 1
      cmdWC(I).Visible = False
   Next
   
   If gblnSqlRows Then
      With prdoWCS
         Do While Not .EOF
            'If pintI < 24 Then
            If pintI < cmdWC.UBound Then
               cmdWC(pintI).Caption = "" & Trim(!WCNNUM)
               If cmdWC(pintI).Caption <> "" Then
                  cmdWC(pintI).Visible = True
               End If
            Else
               MsgBox "Only " & cmdWC.UBound & " workcenters may be shown at once.  " _
                  & "If you are trying to view ALL shops, you must select a single shop.", vbInformation
               Exit Do
            End If
            .MoveNext
            pintI = pintI + 1
         Loop
      End With
   End If
   Set prdoWCS = Nothing
   
   'add "ALL"
   cmdWC(pintI).Caption = "ALL"
   cmdWC(pintI).Visible = True
End Sub

Public Sub ShowJobs()
   'Dim prdoMOs     As rdoResultset
   Dim pintI As Integer
   Dim pstrItem As String
   
   gbytScreen = jobs
   mlngGridNormal = grdMO.CellBackColor
   picLogin.Visible = False
   picShops.Visible = False
   picWCS.Visible = False
   picJobs.Visible = True
   picComplete.Visible = False
   
   'cmdPgUP.Visible = True
   'cmdPgDn.Visible = True
   SetSelectionPageControlsVisibility True
   cmdBack.Visible = True
   cmdProceed.Visible = False
   
   '    If mblnRevise Then
   '        cmdCurrentOps.Visible = False
   '        cmdOpenOps.Visible = False
   '    Else
   '        cmdCurrentOps.Visible = True
   '        cmdOpenOps.Visible = True
   '        cmdCurrentOps.Enabled = CBool(GetSetting("Esi2000", "EsiPOM", "ShowOpenOps", "True"))
   '        cmdOpenOps.Enabled = Not cmdCurrentOps.Enabled
   '    End If
   
   Unload frm12Key
   
   If mblnRevise Then
      lblMOMsg = SYSMSG6
   Else
      lblMOMsg = SYSMSG5 & mempCurrentEmployee.strCurWC & " shop " & mempCurrentEmployee.strCurShop
   End If
   
   With grdMO
      .Cols = 10
      .Clear
      .Rows = 1
      .Row = 0
      
      .Col = 0
      .Text = "Part Number"
      .ColWidth(0) = 3240
      
      .Col = 1
      .Text = "Run"
      .ColWidth(1) = 550
      
      .Col = 2
      .Text = "Qty"
      .ColWidth(2) = 610
      
      .Col = 3
      .Text = "OP"
      .ColWidth(3) = 500
      
      .Col = 4
      .Text = "Start"
      .ColWidth(4) = 1000
      
      .Col = 5
      .Text = "OP Comments"
      .ColWidth(5) = 3000
      
      .Col = 6
      .Text = "Status"
      .ColWidth(6) = 750
      
      .Col = 7
      .Text = "Pri"
      .ColWidth(7) = 500
      
      .Col = 8
      .Text = "WC"
      .ColWidth(8) = 0
      
      .Col = 9
      .Text = "Shop"
      .ColWidth(9) = 0
   End With
   
   FillMos
   
   'Set prdoMOs = Nothing
   
   'Highlight_Jobs_Logged_On grdMO
End Sub

Private Function GetWCShop(strPartRef As String, Run As Long, _
                                Opnum As Integer, ByRef wcshop As String, _
                                ByRef wccenter As String) As Boolean
   Dim rdoOPNum As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT OPSHOP, OPCENTER FROM rnopTable where OPREF = '" & strPartRef & "' " _
          & " AND OPRUN =" & CStr(Run) & " AND OPNO = '" & CStr(Opnum) & "'"
   
   gblnSqlRows = clsADOCon.GetDataSet(sSql, rdoOPNum)
   If gblnSqlRows Then
      With rdoOPNum
         wcshop = Trim(!OPSHOP)
         wccenter = Trim(!OPCENTER)
      End With
   Else
     wcshop = ""
     wccenter = ""
   End If
   
   Set rdoOPNum = Nothing
   
   Exit Function
   
DiaErr1:
   Set rdoOPNum = Nothing
   If Err <> 0 Then
      DisplayError
   End If

End Function

Private Sub ShowComplete()
   gbytScreen = complete
   
   picLogin.Visible = False
   picShops.Visible = False
   picWCS.Visible = False
   picJobs.Visible = False
   picComplete.Visible = True
   
   '    cmdPgUP.Visible = False
   '    cmdPgDn.Visible = False
   SetSelectionPageControlsVisibility False
   cmdBack.Visible = True
   cmdProceed.Visible = True
   cmdWCStat.Visible = True
   ' TODO : Enable Later
   'bret = CheckOpTen(mjobComplete.strPart, mjobComplete.lngRun, mjobComplete.intOp)
   
   'If (CInt(mjobComplete.intOp) = 10) Then
   '   cmdMelterLog.Visible = True
   'End If
      
   lblCom.Caption = ""
   lblRej.Caption = ""
   lblScrap.Caption = ""
   chkComplete.Value = vbUnchecked
   
   cmdCom.Caption = ""
   cmdRej.Caption = ""
   cmdScrap.Caption = ""
   
   lblCom.Visible = False
   lblRej.Visible = False
   lblScrap.Visible = False
   
   cmdCom.Visible = True
   cmdRej.Visible = True
   cmdScrap.Visible = True
   
   
   Dim bret As Boolean
   bret = MelterLogEnabled
   If (bret = True) Then
      If (CInt(mjobComplete.intOp) = 10) Then
         bret = CheckPermMold(mjobComplete.strPart, mjobComplete.lngRun, mjobComplete.intOp)
         If (bret = True) Then
            cmdMelterLog_Click
         End If
         cmdMelterLog.Visible = True
      Else
         cmdMelterLog.Visible = False
      End If
   End If
      
      
   
   lblJob = Trim(mjobComplete.strPart) _
            & "  Run " & mjobComplete.lngRun _
            & "  Op " & mjobComplete.intOp
End Sub

Private Function CheckPermMold(strPartRef As String, Run As Long, Opnum As Integer) As Boolean
   Dim rdoOPNum As ADODB.Recordset
   Dim wcshop As String
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT OPSHOP FROM rnopTable where OPREF = '" & strPartRef & "' " _
          & " AND OPRUN =" & CStr(Run) & " AND OPNO = '" & CStr(Opnum) & "'"
   
   gblnSqlRows = clsADOCon.GetDataSet(sSql, rdoOPNum)
   If gblnSqlRows Then
      With rdoOPNum
         wcshop = Trim(!OPSHOP)
         If (wcshop = "PM") Then
            CheckPermMold = True
         Else
            CheckPermMold = False
         End If
      End With
   Else
      CheckPermMold = False
   End If
   
   Set rdoOPNum = Nothing
   
   Exit Function
   
DiaErr1:
   Set rdoOPNum = Nothing
   If Err <> 0 Then
      DisplayError
   End If

End Function

Private Function CheckOpTen(strPartRef As String, Run As Long, Opnum As String) As Boolean
   Dim rdoOPNum As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT COUNT(*) WHERE PartTable WHERE PARTREF = '" & strPartRef & "' " _
          & "AND PKMOPART='" & mjobComplete.strPart & "' AND PKMORUN=" _
          & mjobComplete.lngRun & " AND PKTYPE=9 ORDER BY PARTREF"
   
   gblnSqlRows = clsADOCon.GetDataSet(sSql, rdoOPNum)
   If gblnSqlRows Then
   Else
   End If
   
DiaErr1:
   Set rdoOPNum = Nothing
   If Err <> 0 Then
      DisplayError
   End If

End Function

Public Sub ShowLots(pstrPart As String, psngQty As Single)
   gbytScreen = Lots
   picLogin.Visible = False
   picShops.Visible = False
   picWCS.Visible = False
   picJobs.Visible = True
   picComplete.Visible = False
   
   lblPickQty = psngQty
   
   lblMOMsg = SYSMSG12 & mstrLotPart & "  Qty. Required " & _
              Format(psngQty, "0.000") & " \ Qty. Select 0.000"
   
   With grdMO
      .Cols = 5
      .Clear
      .Rows = 1
      .Row = 0
      
      .Col = 0
      .Text = "System ID"
      .ColWidth(0) = 3000
      
      .Col = 1
      .Text = "User ID"
      .ColWidth(1) = 3000
      
      .Col = 2
      .Text = "Date"
      .ColWidth(2) = 1500
      
      .Col = 3
      .Text = "Remaining Qty"
      .ColWidth(3) = 1500
      
      .Col = 4
      .Text = "Select Qty"
      .ColWidth(4) = 1500
   End With
   
   FillLots Compress(pstrPart)
   
End Sub

Public Sub ShowPkList()
   gbytScreen = PKLIST
   
   mblnRevise = False
   mblnPick = True
   
   mlngGridNormal = 0 'grdMO.CellBackColor
   picLogin.Visible = False
   picShops.Visible = False
   picWCS.Visible = False
   picJobs.Visible = True
   picComplete.Visible = False
   
   '    cmdPgUP.Visible = True
   '    cmdPgDn.Visible = True
   SetSelectionPageControlsVisibility True
   cmdBack.Visible = True
   cmdProceed.Visible = False
   
   lblMOMsg = SYSMSG10 & " " & mjobComplete.intOp & " (Pick Op)"
   
   With grdMO
      .Cols = 7
      .Clear
      .Rows = 1
      .Row = 0
      
      .Col = 0
      .Text = "Part Number"
      .ColWidth(0) = 3000
      
      .Col = 1
      .Text = "Description"
      .ColWidth(1) = 3000
      
      .Col = 2
      .Text = "Qty"
      .ColWidth(2) = 1200
      
      .Col = 3
      .Text = "Units"
      .ColWidth(3) = 900
      
      .Col = 4
      .Text = "Picked"
      .ColWidth(4) = 1200
      
      .Col = 5
      .Text = "Complete"
      .ColWidth(5) = 1500
      
   End With
   
   FillPkList
End Sub

Private Sub cmdShop_Click(Index As Integer)
   mempCurrentEmployee.strCurShop = cmdShop(Index).Caption
   ShowWcs
End Sub

Private Sub cmdWC_Click(Index As Integer)
   mempCurrentEmployee.strCurWC = cmdWC(Index).Caption
   ShowJobs
End Sub

Private Sub Highlight_Grid_Item( _
                                pblnSelect As Boolean, _
                                pgrdMygrid As MSFlexGrid)
   
   Dim pintI As Integer
   Dim plngColorToUse As Long
   
   With pgrdMygrid
      .Row = .RowSel
      
      If pblnSelect Then
         ' highlight row
         plngColorToUse = YELLOW
      Else
         ' return to normal
         plngColorToUse = mlngGridNormal
      End If
      
      For pintI = 0 To (.Cols - 1)
         .Col = pintI
         .CellBackColor = plngColorToUse
      Next
   End With
End Sub

Private Function Something_Is_Selected( _
                                       pgrdMygrid As MSFlexGrid) As Boolean
   Dim pintI As Integer
   With pgrdMygrid
      .Col = 0
      For pintI = 0 To (.Rows - 1)
         .Row = pintI
         If .CellBackColor = YELLOW Then
            Something_Is_Selected = True
            Exit For
         End If
      Next
   End With
End Function

Private Sub Highlight_Jobs_Logged_On( _
                                     pgrdMygrid As MSFlexGrid)
   
   Dim pintI As Integer
   Dim pintK As Integer
   Dim pintJobs As Integer
   
   pintJobs = UBound(mempCurrentEmployee.jobCurMO())
   
   If Trim(mempCurrentEmployee.jobCurMO(0).strPart) <> "" Then
      With pgrdMygrid
         For pintI = 0 To .Rows - 1
            .Row = pintI
            For pintK = 0 To pintJobs
               .Col = 0
               If Trim(mempCurrentEmployee.jobCurMO(pintK).strPart) _
                       = Trim(.Text) Then
                  .Col = 1
                  If mempCurrentEmployee.jobCurMO(pintK).lngRun _
                                                  = Val(.Text) Then
                     Highlight_Grid_Item True, pgrdMygrid
                  End If
               End If
            Next
         Next
      End With
   End If
End Sub

Private Sub LogOnToJobs( _
                        pgrdMygrid As MSFlexGrid)
   
   Dim pintI As Integer
   Dim pintMoCount As Integer
   Dim pstrMO As String
   Dim pintRun As Integer
   Dim pintOp As Integer
   Dim pjobSelect As Job
   
   gstrCurRoutine = "LogOnToJobs"
   On Error GoTo DiaErr1
   
   SystemAlert "Logging " & mempCurrentEmployee.strFirstName & " " _
      & mempCurrentEmployee.strLastName & " On To Jobs", , True
   
   ' Check in no jobs perviously logged onto.
   ' If so close indirect time charge before assinging
   ' the job.
   If mempCurrentEmployee.jobCurMO(0).lngRun = 0 Then
      CloseIndirectTC mempCurrentEmployee
   End If
   
   On Error Resume Next
   'RdoCon.BeginTrans
   
   With pgrdMygrid
      For pintI = 0 To (.Rows - 1)
         .Row = pintI
         .Col = 0
         If .CellBackColor = YELLOW Then
            pstrMO = Compress(.Text)
            
            
            .Col = 1
            pintRun = Val(.Text)
            
            
            .Col = 3
            pintOp = Val(.Text)
            
            .Col = 8
            mempCurrentEmployee.strCurWC = .Text
            
            .Col = 9
            mempCurrentEmployee.strCurShop = .Text
            
            
            '                ' Add partial time charges to database...
            '                sSql = "INSERT INTO IstcTable (ISEMPLOYEE,ISMO," _
            '                    & "ISRUN,ISOP,ISMOSTART,ISSHOP,ISWCNT) " _
            '                    & "VALUES (" & mempCurrentEmployee.intNumber & ",'" _
            '                    & Trim(pstrMO) & "'," _
            '                    & pintRun & "," _
            '                    & pintOp & ",'" _
            '                    & GetServerDateTime() & "','" & mempCurrentEmployee.strCurShop _
            '                    & "','" & mempCurrentEmployee.strCurWC & " ')"
            '                RdoCon.Execute sSql, rdExecDirect
            LogInToJob pstrMO, pintRun, pintOp, bSetupTime
            
         End If
      Next
   End With
   
   '    If Err = 0 Then
   '        RdoCon.CommitTrans
   '    Else
   '        RdoCon.RollbackTrans
   '        MsgBox Err
   '    End If
   Exit Sub
DiaErr1:
   gstrCurRoutine = "LogOnToJobs"
   DisplayError
End Sub


Private Sub FillMos()
   Dim prdoMOs As ADODB.Recordset
   Dim pstrItem As String
   Dim pintI As Integer
   Dim pstrStart As String
   Dim pstrSetup As String
   Dim pstrRun As String
   
   Dim iFakeSortColumn As Integer
   Dim dblSortDate As Double
   
   
   
   On Error GoTo DiaErr1
   MouseCursor ccHourglass
   
   'show jobs that can be logged out of
   If mblnRevise Then
'      sSql = "SELECT DISTINCT PARTNUM, ISRUN, RUNQTY, ISOP, ISMOSTART, " _
'             & "left(OPCOMT,255) as comments, RUNSTATUS,RUNPRIORITY,OPCENTER,WCNSHOP,ISEMPLOYEE, " _
'             & "PAPICLINK3,OPSUDATE,OPRUNDATE" & vbCrLf _
'             & "FROM IstcTable INNER JOIN " _
'             & "RnopTable ON IstcTable.ISRUN = RnopTable.OPRUN AND " & vbCrLf _
'             & "IstcTable.ISOP = RnopTable.OPNO AND IstcTable.ISMO = " _
'             & "RnopTable.OPREF INNER JOIN " _
'             & "PartTable ON IstcTable.ISMO = PartTable.PARTREF INNER JOIN " _
'             & "RunsTable ON IstcTable.ISMO = RunsTable.RUNREF AND " _
'             & "IstcTable.ISRUN = RunsTable.RUNNO " _
'             & "JOIN WcntTable on OpCenter = WCNREF" & vbCrLf _
'             & "AND OpShop = WCNSHOP" & vbCrLf _
'             & "WHERE ISEMPLOYEE = " & mempCurrentEmployee.intNumber & vbCrLf _
'             & "and ReadyToDelete = 0 and ISINDIRECT = 0" & vbCrLf _
'             & "ORDER BY OPSUDATE"

      sSql = "SELECT DISTINCT PARTNUM, ISRUN, RUNQTY, ISOP, " _
             & "left(OPCOMT,255) as comments, RUNSTATUS,RUNPRIORITY,OPCENTER,WCNSHOP,ISEMPLOYEE, " _
             & "PAPICLINK3,OPSUDATE,OPRUNDATE" & vbCrLf _
             & "FROM IstcTable INNER JOIN " _
             & "RnopTable ON IstcTable.ISRUN = RnopTable.OPRUN AND " & vbCrLf _
             & "IstcTable.ISOP = RnopTable.OPNO AND IstcTable.ISMO = " _
             & "RnopTable.OPREF INNER JOIN " _
             & "PartTable ON IstcTable.ISMO = PartTable.PARTREF INNER JOIN " _
             & "RunsTable ON IstcTable.ISMO = RunsTable.RUNREF AND " _
             & "IstcTable.ISRUN = RunsTable.RUNNO " _
             & "JOIN WcntTable on OpCenter = WCNREF" & vbCrLf _
             & "AND OpShop = WCNSHOP" & vbCrLf _
             & "WHERE ISEMPLOYEE = " & mempCurrentEmployee.intNumber & vbCrLf _
             & "and ReadyToDelete = 0 and ISINDIRECT = 0" & vbCrLf _
             & "ORDER BY OPSUDATE"

      
      'show jobs that can be logged into
   Else
      '       sSql = "SELECT DISTINCT PARTNUM, RUNNO, RUNQTY, OPNO, left(OPCOMT,255) as comments," & vbCrLf _
      '            & "RUNSTATUS,RUNPRIORITY,OPCENTER,WCNSHOP,ISEMPLOYEE,PAPICLINK3,OPSUDATE,OPRUNDATE" & vbCrLf _
      '            & "FROM RunsTable runs " & vbCrLf _
      '            & "JOIN PartTable parts ON runs.RUNREF = parts.PARTREF" & vbCrLf _
      '            & "JOIN RnopTable ops ON runs.RUNREF = ops.OPREF" & vbCrLf _
      '            & "AND runs.RUNNO = ops.OPRUN" & vbCrLf
      
'      sSql = "SELECT DISTINCT PARTNUM, RUNNO, RUNQTY, OPNO, left(OPCOMT,255) as comments," & vbCrLf _
'             & "RUNSTATUS,RUNPRIORITY,OPCENTER,WCNSHOP,PAPICLINK3,OPSUDATE,OPRUNDATE" & vbCrLf _
'             & "FROM RunsTable runs " & vbCrLf _
'             & "JOIN PartTable parts ON runs.RUNREF = parts.PARTREF" & vbCrLf _
'             & "JOIN RnopTable ops ON runs.RUNREF = ops.OPREF" & vbCrLf _
'             & "AND runs.RUNNO = ops.OPRUN" & vbCrLf
'
      sSql = "SELECT DISTINCT PARTNUM, RUNNO, RUNQTY, OPNO, left(OPCOMT,255) as comments," & vbCrLf _
             & "RUNSTATUS,RUNPRIORITY,OPCENTER,WCNSHOP,PAPICLINK3,OPSUDATE,OPRUNDATE" & vbCrLf
             
      'if previous op must be closed to log in, add that check here
      If denyLoginIfPriorOpOpen Then
         sSql = sSql & ",ISNULL((select OPCOMPLETE from RnopTable prev where prev.OPREF = ops.OPREF and prev.OPRUN = ops.OPRUN" & vbCrLf _
            & "and prev.OPNO = (select max(opno) from RnopTable ops3 where ops3.OPREF = ops.OPREF" & vbCrLf _
            & "and ops3.OPRUN = ops.OPRUN and ops3.OPNO < ops.OPNO)),1) as OKTOLOGIN" & vbCrLf
      Else
         sSql = sSql & ",1 as OKTOLOGIN" & vbCrLf
      End If
      
      sSql = sSql & "FROM RunsTable runs " & vbCrLf _
             & "JOIN PartTable parts ON runs.RUNREF = parts.PARTREF" & vbCrLf _
             & "JOIN RnopTable ops ON runs.RUNREF = ops.OPREF" & vbCrLf _
             & "AND runs.RUNNO = ops.OPRUN" & vbCrLf
      
'      sSql = "SELECT DISTINCT PARTNUM, RUNNO, RUNQTY, OPNO, left(OPCOMT,255) as comments," & vbCrLf _
'             & "RUNSTATUS,RUNPRIORITY,OPCENTER,WCNSHOP,PAPICLINK3,OPSUDATE,OPRUNDATE" & vbCrLf _
'             & "FROM RunsTable runs " & vbCrLf _
'             & "JOIN PartTable parts ON runs.RUNREF = parts.PARTREF" & vbCrLf _
'             & "JOIN RnopTable ops ON runs.RUNREF = ops.OPREF" & vbCrLf _
'             & "AND runs.RUNNO = ops.OPRUN" & vbCrLf
      
      'showing current ops only?
      If cmdOpenOps.Enabled Then
         sSql = sSql _
                & "AND runs.RUNOPCUR = ops.OPNO" & vbCrLf
      End If
      
      '        sSql = sSql _
      '            & "JOIN WcntTable wcs on OpCenter = WCNREF" & vbCrLf _
      '            & "AND OpShop = WCNSHOP" & vbCrLf _
      '            & "LEFT OUTER JOIN IstcTable ON ops.OPREF = IstcTable.ISMO" & vbCrLf _
      '            & "AND ops.OPRUN = IstcTable.ISRUN AND ops.OPNO = IstcTable.ISOP" & vbCrLf _
      '            & "WHERE (runs.RUNSTATUS <> 'SC' and runs.RUNSTATUS <> 'CL' and runs.RUNSTATUS <> 'CA')" & vbCrLf _
      '            & "AND OPCOMPLETE = 0" & vbCrLf
      
      '        sSql = sSql _
      '            & "JOIN WcntTable wcs on OpCenter = WCNREF" & vbCrLf _
      '            & "AND OpShop = WCNSHOP" & vbCrLf _
      '            & "WHERE (runs.RUNSTATUS <> 'SC' and runs.RUNSTATUS <> 'CL' and runs.RUNSTATUS <> 'CA')" & vbCrLf _
      '            & "AND OPCOMPLETE = 0" & vbCrLf
      '
      
      'include any ops that are not currently open for the employee
      sSql = sSql _
             & "JOIN WcntTable wcs on OpCenter = WCNREF" & vbCrLf _
             & "AND OpShop = WCNSHOP" & vbCrLf _
             & "LEFT JOIN IstcTable it on it.ISEMPLOYEE = " & mempCurrentEmployee.intNumber & vbCrLf _
             & "AND it.ISMO = PARTS.PARTREF" & vbCrLf _
             & "AND it.ISRUN = runs.RUNNO" & vbCrLf _
             & "AND it.ISOP = ops.OPNO" & vbCrLf _
             & "AND it.ISMOSTART IS NOT NULL" & vbCrLf _
             & "AND it.ISMOEND IS NULL" & vbCrLf _
             & "WHERE (runs.RUNSTATUS <> 'SC' and runs.RUNSTATUS <> 'CL' and runs.RUNSTATUS <> 'CA')" & vbCrLf _
             & "AND OPCOMPLETE = 0" & vbCrLf _
             & "AND it.ISMO IS NULL" & vbCrLf
      
      If mempCurrentEmployee.strCurShop <> "ALL" Then
         sSql = sSql & "AND OPSHOP = '" & Trim(mempCurrentEmployee.strCurShop) & "'" & vbCrLf
      End If
      
      If mempCurrentEmployee.strCurWC = "ALL" Then
         '            If mempCurrentEmployee.strCurShop = "ALL" Then
         '                'select all ops
         '            Else
         '                'select all ops for selected shop
         '                sSql = sSql & "and OPSHOP = '" & Trim(mempCurrentEmployee.strCurShop) & "'" & vbCrLf
         '            End If
      Else
         'just show the ops for the selected workcenter
         sSql = sSql & "AND OPCENTER = '" & Trim(mempCurrentEmployee.strCurWC) & "'" & vbCrLf
      End If
      
      sSql = sSql & "ORDER BY OPSUDATE"
   End If
   
   gblnSqlRows = clsADOCon.GetDataSet(sSql, prdoMOs, ES_KEYSET)
   
   
'Make sure all the fake sort columns are hidden
        grdMO.Cols = grdMO.Cols + 5
        For iFakeSortColumn = grdMO.Cols - 6 To grdMO.Cols - 1
            grdMO.ColWidth(iFakeSortColumn) = 0
        Next
        
        
        
   
   
   If gblnSqlRows Then
      With prdoMOs
         While Not .EOF
            pstrSetup = "" & !OPSUDATE
            pstrRun = "" & !OPRUNDATE
            
            If pstrSetup > pstrRun Then
               pstrStart = pstrRun
            Else
               pstrStart = pstrSetup
            End If
            
            If pstrStart <> "" Then
               pstrStart = Format(pstrStart, "m/d/yy")
            End If
            
            
            If pstrStart = "" Then dblSortDate = 0 Else dblSortDate = DateValue(pstrStart)
            
            pstrItem = " " & .Fields(0) & Chr(9) _
                       & .Fields(1) & Chr(9) _
                       & .Fields(2) & Chr(9) _
                       & .Fields(3) & Chr(9) _
                       & pstrStart & Chr(9) _
                       & .Fields(4) & Chr(9) _
                       & .Fields(5) & Chr(9) _
                       & .Fields(6) & Chr(9) _
                       & Trim(.Fields(7)) & Chr(9) _
                       & Trim(.Fields(8)) & Chr(9) _
                       & Trim(str(dblSortDate)) & Chr(9) _
                       & PadNumber(.Fields(1), 20) & Chr(9) _
                       & PadNumber(.Fields(2), 20) & Chr(9) _
                       & PadNumber(.Fields(3), 20) & Chr(9) _
                       & PadNumber(.Fields(6), 20) & Chr(9)
                       
            
            grdMO.AddItem pstrItem
            grdMO.Row = grdMO.Rows - 1
            grdMO.Col = 0
            grdMO.CellPictureAlignment = flexAlignRightCenter
            
            'change op color to red to prevent login if prev op open not allowed
            If Not mblnRevise Then
               grdMO.Col = 3
               If .Fields(12) = 1 Then
                  'grdMO.CellForeColor = ES_BLACK
               Else
                  grdMO.CellForeColor = ES_RED
               End If
            End If
            
            On Error Resume Next
            ' If bad picture path then just keep on going
            imgBuf.Picture = LoadPicture("" & Trim(!PAPICLINK3))
            Set grdMO.CellPicture = imgBuf
            Set imgBuf = Nothing
            On Error GoTo DiaErr1
            'End If
            .MoveNext
         Wend
         Set prdoMOs = Nothing
      End With

      MouseCursor ccDefault
   Else
      MouseCursor ccDefault
      SystemAlert SYSMSG9 & " " _
         & Trim(mempCurrentEmployee.strCurWC) & "."
   End If
   Exit Sub
   
DiaErr1:
   MouseCursor ccDefault
   gstrCurRoutine = "FillMos"
   Set prdoMOs = Nothing
   DisplayError
End Sub

Private Sub tmrScroll_Timer()
   If grdMO.Rows > 1 Then
      If grdMO.TopRow = 1 And mintScrollMode < 0 Then
         ' Stop we are at the top of the list
      Else
         grdMO.TopRow = grdMO.TopRow + mintScrollMode
      End If
   End If
End Sub

Private Sub FillLots(pstrPart As String)
   
   Dim prdoLot As ADODB.Recordset
   Dim pstrItem As String
   'Dim I      As Integer
   'Dim cQty   As Currency
   
   'On Error GoTo DiaErr1
   'Erase vLots
   'Es_LotsSelected = 0
   'Es_TotalLots = 0
   'iTotalItems = 0
   
   sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTADATE," _
          & "LOTREMAININGQTY " _
          & "FROM LohdTable WHERE (LOTPARTREF='" _
          & pstrPart & "' AND LOTREMAININGQTY>0) "
   'If lblLifo = "LIFO" Then sSql = sSql & " ORDER BY LOTNUMBER DESC"
   
   
   
   gblnSqlRows = clsADOCon.GetDataSet(sSql, prdoLot, ES_KEYSET)
   If gblnSqlRows Then
      With prdoLot
         Do Until .EOF
            pstrItem = " " & .Fields(0) & Chr(9) _
                       & " " & .Fields(1) & Chr(9) _
                       & Format(.Fields(2), "mm/dd/yy") _
                       & Chr(9) & .Fields(3) & Chr(9) & "0.000"
            grdMO.AddItem pstrItem
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set prdoLot = Nothing
   'lblTotLots = iTotalItems
   
End Sub


Private Sub FillPkList()
   Dim prdoPK As ADODB.Recordset
   Dim pintI As Integer
   Dim pstrItem As String
   
   On Error GoTo DiaErr1
   
   gstrCurRoutine = "FillPkList"
   
   sSql = "SELECT PARTNUM,PADESC,PKPQTY,PKUNITS,PASTDCOST," _
          & "PALOCATION,PKPARTREF,PKMOPART,PKMORUN,PKREV,PKTYPE," _
          & "PKAQTY,PKRECORD,PARTREF,PAPICLINK3 FROM PartTable,MopkTable " _
          & "WHERE PARTREF=PKPARTREF " _
          & "AND PKMOPART='" & mjobComplete.strPart & "' AND PKMORUN=" _
          & mjobComplete.lngRun & " AND PKTYPE=9 ORDER BY PARTREF"
   
   gblnSqlRows = clsADOCon.GetDataSet(sSql, prdoPK, ES_KEYSET)
   If gblnSqlRows Then
      With prdoPK
         While Not .EOF
            pstrItem = " " & .Fields(0) & Chr(9) _
                       & .Fields(1) & Chr(9) & .Fields(2) & Chr(9) _
                       & .Fields(3) & Chr(9) & .Fields(4) & Chr(9) _
                       & "No" & Chr(9) & !PASTDCOST
            grdMO.AddItem pstrItem
            
            grdMO.Row = grdMO.Rows - 1
            grdMO.Col = 0
            grdMO.CellPictureAlignment = flexAlignRightCenter
            imgBuf.Picture = LoadPicture("" & Trim(!PAPICLINK3))
            Set grdMO.CellPicture = imgBuf
            .MoveNext
         Wend
      End With
   Else
      SystemAlert SYSMSG11
   End If
   
DiaErr1:
   Set prdoPK = Nothing
   If Err <> 0 Then
      DisplayError
   End If
End Sub


Private Function MelterLogEnabled() As Boolean
    Dim rdoMel As ADODB.Recordset
    Dim iDL As Integer
    
    MelterLogEnabled = False
    sSql = "SELECT ENABLEMELTERSLOG FROM ComnTable WHERE COREF=1"
    bSqlRows = clsADOCon.GetDataSet(sSql, rdoMel, ES_FORWARD)
    If bSqlRows Then
        iDL = 0 & rdoMel!ENABLEMELTERSLOG
        If iDL = 1 Then MelterLogEnabled = True
    End If
    Set rdoMel = Nothing
End Function

Private Function MOComtEnabled() As Boolean
    Dim rdoMel As ADODB.Recordset
    Dim iDL As Integer
    
    MOComtEnabled = False
    sSql = "SELECT ISNULL(COALLOWMOCOMT, 0) COALLOWMOCOMT FROM ComnTable WHERE COREF=1"
    bSqlRows = clsADOCon.GetDataSet(sSql, rdoMel, ES_FORWARD)
    If bSqlRows Then
        iDL = 0 & rdoMel!COALLOWMOCOMT
        If iDL = 1 Then MOComtEnabled = True
    End If
    Set rdoMel = Nothing
End Function

