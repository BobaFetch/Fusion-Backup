VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLe04a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fiscal Years"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   5160
   Begin VB.CommandButton cmdApl 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4200
      TabIndex        =   56
      ToolTipText     =   "Apply Format To All Following Years"
      Top             =   2040
      Width           =   875
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   3720
      TabIndex        =   53
      Top             =   5760
      Width           =   255
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   3720
      TabIndex        =   49
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   3720
      TabIndex        =   45
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3720
      TabIndex        =   41
      Top             =   4680
      Width           =   255
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3720
      TabIndex        =   37
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   33
      Top             =   3960
      Width           =   255
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   29
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   25
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   21
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   17
      ToolTipText     =   "Delete Period"
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   13
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   9
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "é"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   5
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "ê"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   3480
      TabIndex        =   52
      ToolTipText     =   "New Period"
      Top             =   5760
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   3480
      TabIndex        =   48
      ToolTipText     =   "New Period"
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   3480
      TabIndex        =   44
      ToolTipText     =   "New Period"
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3480
      TabIndex        =   40
      ToolTipText     =   "New Period"
      Top             =   4680
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3480
      TabIndex        =   36
      ToolTipText     =   "New Period"
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3480
      TabIndex        =   32
      ToolTipText     =   "New Period"
      Top             =   3960
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   28
      ToolTipText     =   "New Period"
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   24
      ToolTipText     =   "New Period"
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   20
      ToolTipText     =   "New Period"
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   16
      ToolTipText     =   "New Period"
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   12
      ToolTipText     =   "New Period"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   8
      ToolTipText     =   "New Period"
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   4
      ToolTipText     =   "New Period"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton cmdDsl 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   4200
      TabIndex        =   54
      ToolTipText     =   "Select A Different Year"
      Top             =   1200
      Width           =   875
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "Select"
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      ToolTipText     =   "Update Structure And Associated Entries"
      Top             =   480
      Width           =   875
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   72
      Top             =   1080
      Width           =   4935
   End
   Begin VB.ComboBox cmbYer 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Tag             =   "1"
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox txtsDte 
      Height          =   315
      Index           =   12
      Left            =   960
      TabIndex        =   50
      Tag             =   "4"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.ComboBox txteDte 
      Height          =   315
      Index           =   12
      Left            =   2280
      TabIndex        =   51
      Tag             =   "4"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.ComboBox txtsDte 
      Height          =   315
      Index           =   11
      Left            =   960
      TabIndex        =   46
      Tag             =   "4"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.ComboBox txteDte 
      Height          =   315
      Index           =   11
      Left            =   2280
      TabIndex        =   47
      Tag             =   "4"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.ComboBox txtsDte 
      Height          =   315
      Index           =   10
      Left            =   960
      TabIndex        =   42
      Tag             =   "4"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.ComboBox txteDte 
      Height          =   315
      Index           =   10
      Left            =   2280
      TabIndex        =   43
      Tag             =   "4"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.ComboBox txtsDte 
      Height          =   315
      Index           =   9
      Left            =   960
      TabIndex        =   38
      Tag             =   "4"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.ComboBox txteDte 
      Height          =   315
      Index           =   9
      Left            =   2280
      TabIndex        =   39
      Tag             =   "4"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.ComboBox txtsDte 
      Height          =   315
      Index           =   8
      Left            =   960
      TabIndex        =   34
      Tag             =   "4"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.ComboBox txteDte 
      Height          =   315
      Index           =   8
      Left            =   2280
      TabIndex        =   35
      Tag             =   "4"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.ComboBox txtsDte 
      Height          =   315
      Index           =   7
      Left            =   960
      TabIndex        =   30
      Tag             =   "4"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.ComboBox txteDte 
      Height          =   315
      Index           =   7
      Left            =   2280
      TabIndex        =   31
      Tag             =   "4"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.ComboBox txtsDte 
      Height          =   315
      Index           =   6
      Left            =   960
      TabIndex        =   26
      Tag             =   "4"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ComboBox txteDte 
      Height          =   315
      Index           =   6
      Left            =   2280
      TabIndex        =   27
      Tag             =   "4"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ComboBox txtsDte 
      Height          =   315
      Index           =   5
      Left            =   960
      TabIndex        =   22
      Tag             =   "4"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ComboBox txteDte 
      Height          =   315
      Index           =   5
      Left            =   2280
      TabIndex        =   23
      Tag             =   "4"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ComboBox txtsDte 
      Height          =   315
      Index           =   4
      Left            =   960
      TabIndex        =   18
      Tag             =   "4"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ComboBox txteDte 
      Height          =   315
      Index           =   4
      Left            =   2280
      TabIndex        =   19
      Tag             =   "4"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ComboBox txtsDte 
      Height          =   315
      Index           =   3
      Left            =   960
      TabIndex        =   14
      Tag             =   "4"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ComboBox txteDte 
      Height          =   315
      Index           =   3
      Left            =   2280
      TabIndex        =   15
      Tag             =   "4"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ComboBox txtsDte 
      Height          =   315
      Index           =   2
      Left            =   960
      TabIndex        =   10
      Tag             =   "4"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox txteDte 
      Height          =   315
      Index           =   2
      Left            =   2280
      TabIndex        =   11
      Tag             =   "4"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox txtsDte 
      Height          =   315
      Index           =   1
      Left            =   960
      TabIndex        =   6
      Tag             =   "4"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox txteDte 
      Height          =   315
      Index           =   1
      Left            =   2280
      TabIndex        =   7
      Tag             =   "4"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox txtsDte 
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox txteDte 
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4200
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Update"
      Height          =   315
      Left            =   4200
      TabIndex        =   55
      ToolTipText     =   "Update Fiscal Periods"
      Top             =   1560
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4560
      Top             =   4800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6270
      FormDesignWidth =   5160
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   79
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
      PictureUp       =   "diaGLe04a.frx":0000
      PictureDn       =   "diaGLe04a.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End                 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   2280
      TabIndex        =   78
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start                "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   960
      TabIndex        =   77
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label lblStat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   76
      Top             =   6120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label lblPrd 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2760
      TabIndex        =   75
      Top             =   480
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periods:"
      Height          =   315
      Index           =   17
      Left            =   2040
      TabIndex        =   74
      Top             =   480
      Width           =   555
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   73
      Top             =   480
      Width           =   555
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      Height          =   315
      Index           =   15
      Left            =   360
      TabIndex        =   71
      Top             =   5760
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      Height          =   315
      Index           =   14
      Left            =   360
      TabIndex        =   70
      Top             =   5400
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      Height          =   315
      Index           =   13
      Left            =   360
      TabIndex        =   69
      Top             =   5040
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   315
      Index           =   12
      Left            =   360
      TabIndex        =   68
      Top             =   4680
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      Height          =   315
      Index           =   11
      Left            =   360
      TabIndex        =   67
      Top             =   4320
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   315
      Index           =   10
      Left            =   360
      TabIndex        =   66
      Top             =   3960
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   315
      Index           =   9
      Left            =   360
      TabIndex        =   65
      Top             =   3600
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   315
      Index           =   8
      Left            =   360
      TabIndex        =   64
      Top             =   3240
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   315
      Index           =   7
      Left            =   360
      TabIndex        =   63
      Top             =   2880
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   315
      Index           =   6
      Left            =   360
      TabIndex        =   62
      Top             =   2520
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   315
      Index           =   5
      Left            =   360
      TabIndex        =   61
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   315
      Index           =   4
      Left            =   360
      TabIndex        =   60
      Top             =   1815
      Width           =   195
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   315
      Index           =   1
      Left            =   360
      TabIndex        =   59
      Top             =   1470
      Width           =   195
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period      "
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
      Left            =   150
      TabIndex        =   58
      Top             =   1200
      Width           =   465
   End
End
Attribute VB_Name = "diaGLe04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'************************************************************************************
' diaGLe04a - Fiscal Year (Revise Periods)
'
' Notes:
'
' Created: 12/16/03 (JCW)
' Revisions:
'   01/07/04 (nth) Allow addition of fiscal periods.
'   01/08/04 (JCW) Fix Leap Year \ Allow periods to exceede Fiscal Year
'   01/08/04 (JCW) Prevent adding year exceeding Smalldatetime Capacity (2079)
'
'************************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodYears As Byte
Dim bUpdated As Byte
Dim bInvalid As Byte
Dim bGoodYear As Byte
Dim sMsg As String
Dim sMonths(14, 3) As String
Dim iRow As Integer
Dim rdoPrd As ADODB.Recordset

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


'************************************************************************************

Private Sub cmbYer_Click()
   bGoodYear = True
   GetPeriod
End Sub

Private Sub cmbYer_LostFocus()
   If Trim(cmbYer) <> "" Then
      cmbYer = Int(cmbYer)
      
      If cmbYer > 2076 Then 'Database datatype doesnt allow larger than 2079
         cmbYer = ""
         bGoodYear = False
      End If
      
      Select Case Len(cmbYer)
         Case 1, 2, 3
            cmbYer = 2000 + cmbYer
            GetPeriod
         Case 4
            GetPeriod
         Case Is > 4
            cmbYer = Left(cmbYer, 4)
      End Select
      
   Else
      bGoodYear = False
   End If
End Sub

Private Sub cmdApl_Click()
   ApplyToAll
End Sub

Private Sub Form_Deactivate()
   Unload Me
End Sub

Private Sub Form_Load()
   FormLoad Me
   sCurrForm = Caption
   FormatControls
   bOnLoad = True
   ManageBoxes 0
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set rdoPrd = Nothing
   Set diaGLe04a = Nothing
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdDel_Click(Index As Integer)
   lblPrd = Val(lblPrd) - 1
   txteDte(Index) = ""
   txtsDte(Index) = ""
   ManageBoxes 1
   txtsDte(Index - 1).SetFocus
End Sub

Private Sub cmdDsl_Click()
   ManageBoxes 0
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdNew_Click(Index As Integer)
   lblPrd = Val(lblPrd) + 1
   ManageBoxes 1
   txtsDte(Index + 1).SetFocus
End Sub

Private Sub cmdSel_Click()
   'We could just disable the button until a valid year is entered.. but then
   'the user would have no where to tab out to; we couldnt call get period
   If bGoodYear = True Then
      ManageBoxes 1
      LoadPeriods
   Else
      MsgBox "Enter A Valid Fiscal Year.", vbInformation, Caption
   End If
End Sub

Private Sub txtEDte_DropDown(Index As Integer)
   If Trim(txtsDte(Index)) <> "" And Trim(txteDte(Index)) = "" Then
      txteDte(Index) = GetMonthEnd(txtsDte(Index))
   End If
   ShowCalendar Me
End Sub

Private Sub txtEdte_LostFocus(Index As Integer)
   If Not bOnLoad Then
      If Trim(txteDte(Index)) <> "" Then
         txteDte(Index) = CheckDate(txteDte(Index))
      End If
   End If
End Sub

Private Sub txtSDte_DropDown(Index As Integer)
   If Index <> 0 Then
      If Trim(txteDte(Index - 1)) <> "" And Trim(txtsDte(Index)) = "" Then
         txtsDte(Index) = Format(CDate(txteDte(Index - 1)) + 1, "mm/dd/yy")
      End If
   Else
      txtsDte(0) = Format(Now, "1/1/" & cmbYer)
   End If
   ShowCalendar Me
End Sub

Private Sub txtSDte_LostFocus(Index As Integer)
   If Not bOnLoad Then
      If Trim(txtsDte(Index)) <> "" Then
         txtsDte(Index) = CheckDate(txtsDte(Index))
      End If
   End If
End Sub

Private Sub cmdUpd_Click()
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   Update (cmbYer)
   If bInvalid = 0 Then
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         MsgBox cmbYer & " Fiscal Periods Updated Successfully.", vbInformation, Caption
         lblStat = ""
         cmdDsl_Click
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "Error Updating Fiscal Periods.", vbInformation, Caption
         lblStat = ""
      End If
   Else
      bInvalid = 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      If CheckYears Then
         FillYearCombo
      Else
         InitializeYears
      End If
      ManageBoxes 0
      DoEvents
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub ApplyToAll() 'Loops through Update Function
   Dim i As Integer
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   'For i = cmbYer To Val(cmbYer.List(cmbYer.ListCount - 1))
   For i = 0 To cmbYer.ListCount - 1
      If bInvalid = 0 Then
         Update cmbYer.List(i)
      Else
         Exit For
      End If
   Next
   If bInvalid = 0 Then
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         MsgBox "All Fiscal Periods Updated Successfully.", vbInformation, Caption
         lblStat = ""
         cmdDsl_Click
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "Error Updating Fiscal Periods.", vbExclamation, Caption
         lblStat = ""
      End If
   Else
      bInvalid = 0
   End If
End Sub

Private Function FixLeap(sDate As String) As String
   ' Check for leap year
   Dim i As Integer
   Dim bMonth As Integer
   Dim bDay As Integer
   Dim bYear As Integer
   'bDay = Format(sDate, "dd")
   bDay = Right(Left(sDate, 5), 2)
   bYear = Right(sDate, 4)
   bMonth = Val(Left(sDate, 2))
   If bMonth = 2 Then
      For i = 1984 To 2100 Step 4
         If bYear = i Then
            'If bDay = 28 Then bDay = 29
            Exit For
         Else
            If i > bYear Then 'we can say that it is deffinately not leap year
               If bDay = 29 Then bDay = 28
               Exit For
            End If
         End If
      Next
   Else
      FixLeap = sDate
      Exit Function
   End If
   FixLeap = Format(bMonth, "00") & "/" & Format(bDay, "00") & "/" & bYear
End Function

Private Sub Update(iYeari) 'Updates Fiscal Periods
   Dim i As Integer
   Dim x As Integer
   Dim iValidType As Integer
   Dim iValidRow As Integer
   Dim sStart As String
   Dim sEnd As String
   Dim iSdiff As Integer
   Dim iEdiff As Integer
   
   On Error GoTo DiaErr1
   
   bInvalid = 0
   ValidDate iValidType, iValidRow
   Select Case iValidType
      Case 4
         lblStat = "Updating Fiscal Periods."
         iSdiff = Format(txtsDte(i), "yyyy") - cmbYer
         sSql = "UPDATE GlfyTable SET FYSTART = '" & FixLeap(Format(txtsDte(i), "mm/dd/" & iYeari + iSdiff)) & "',"
         For i = 0 To 12
            If Trim(txtsDte(i)) <> "" Then
               iSdiff = Format(txtsDte(i), "yyyy") - cmbYer
               iEdiff = Format(txteDte(i), "yyyy") - cmbYer
               sStart = FixLeap(Format(txtsDte(i), "mm/dd/" & iYeari + iSdiff))
               sEnd = FixLeap(Format(txteDte(i), "mm/dd/" & iYeari + iEdiff))
               sSql = sSql & "FYPERSTART" & CStr(i + 1) & "='" & sStart & "',"
               sSql = sSql & "FYPEREND" & CStr(i + 1) & "='" & sEnd & "',"
               x = x + 1
               If i = 12 Then
                  sSql = sSql & "FYEND ='" & sEnd & "', FYPERIODS=" & x & ","
               End If
            Else
               sSql = sSql & "FYEND ='" & sEnd & "', FYPERIODS=" & x & ","
               Exit For
            End If
         Next
         Do While x <= 12
            sSql = sSql & " FYPERSTART" & CStr(x + 1) & "=NULL,"
            sSql = sSql & " FYPEREND" & CStr(x + 1) & "=NULL,"
            x = x + 1
         Loop
         sSql = Left(sSql, Len(sSql) - 1)
         sSql = sSql & " WHERE FYYEAR = " & Trim(iYeari)
         clsADOCon.ExecuteSQL sSql
         
      Case 3
         MsgBox "All Fields Must be Completed.", vbInformation, Caption
         txtsDte(iValidRow).SetFocus
         bInvalid = 1
      Case 2
         MsgBox "Period Starting Date Must Be Greater Than Previous Ending Date.", vbInformation, Caption
         txtsDte(iValidRow).SetFocus
         bInvalid = 1
      Case 1
         MsgBox "Period Ending Date Must Be Greater Than Starting Date.", vbInformation, Caption
         txteDte(iValidRow).SetFocus
         bInvalid = 1
   End Select
   Exit Sub
   
DiaErr1:
   sProcName = "updatePeriods"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub LoadPeriods() 'Loads Existing Periods
   On Error GoTo DiaErr1
   Dim i As Integer
   Dim x As Integer
   With rdoPrd
      For i = 0 To 12
         txtsDte(i) = "" & Format(.Fields(x + 1), "mm/dd/yy")
         txteDte(i) = "" & Format(.Fields(x + 2), "mm/dd/yy")
         x = x + 2
      Next
   End With
   Exit Sub
DiaErr1:
   sProcName = "LoadPeriods"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillYearCombo() 'Gets Years; Fills Combo
   Dim rdoYrs As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT FYYEAR FROM GlfyTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoYrs)
   If bSqlRows Then
      With rdoYrs
         cmbYer.Clear
         While Not .EOF
            AddComboStr cmbYer.hWnd, "" & !FYYEAR
            .MoveNext
         Wend
         If cmbYer.ListCount > 0 Then cmbYer.ListIndex = 0
      End With
   End If
   Set rdoYrs = Nothing
   Exit Sub
DiaErr1:
   sProcName = "FillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub ManageBoxes(bSelect As Byte) 'Master User Interface Logic
   Dim i As Integer
   Dim bOpp As Byte
   On Error GoTo DiaErr1
   If bSelect = 0 Then
      bOpp = 1
   Else
      bOpp = 0
   End If
   For i = 1 To 13
      If bSelect = 1 Then
         If i <= Val(lblPrd) Then
            txtsDte(i - 1).enabled = True
            txteDte(i - 1).enabled = True
            If i = Val(lblPrd) Then
               cmdNew(i - 1).Visible = True
               cmdDel(i - 1).Visible = True
            Else
               cmdNew(i - 1).Visible = False
               cmdDel(i - 1).Visible = False
            End If
         Else
            If Val(lblPrd) <> 0 Then
               txtsDte(i - 1).enabled = False
               txteDte(i - 1).enabled = False
               cmdNew(i - 1).Visible = False
               cmdDel(i - 1).Visible = False
            Else
               txtsDte(i - 1).enabled = True
               txteDte(i - 1).enabled = True
               lblPrd = "1"
               cmdNew(0).Visible = True
               cmdDel(0).Visible = True
            End If
         End If
      Else
         txtsDte(i - 1).enabled = False
         txteDte(i - 1).enabled = False
         cmdNew(i - 1).Visible = False
         cmdDel(i - 1).Visible = False
         txtsDte(i - 1) = ""
         txteDte(i - 1) = ""
         If Not bOnLoad Then
            GetPeriod
         End If
      End If
   Next
   
   cmdDsl.enabled = bSelect
   cmdUpd.enabled = bSelect
   cmdApl.enabled = bSelect
   cmbYer.enabled = bOpp
   cmdSel.enabled = bOpp
   Exit Sub
DiaErr1:
   sProcName = "ManageBoxes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetPeriod() 'Dumps Data in resultset For Later Use
   Dim bResponse As Byte
   Dim itempYr As Integer
   On Error GoTo DiaErr1
   sSql = "SELECT FYPERIODS,FYPERSTART1,FYPEREND1, " _
          & "FYPERSTART2,FYPEREND2,FYPERSTART3,FYPEREND3, " _
          & "FYPERSTART4,FYPEREND4,FYPERSTART5,FYPEREND5, " _
          & "FYPERSTART6,FYPEREND6,FYPERSTART7,FYPEREND7, " _
          & "FYPERSTART8,FYPEREND8,FYPERSTART9,FYPEREND9, " _
          & "FYPERSTART10,FYPEREND10,FYPERSTART11,FYPEREND11, " _
          & "FYPERSTART12,FYPEREND12,FYPERSTART13,FYPEREND13 " _
          & " FROM GlfyTable WHERE " _
          & "FYYEAR = " & Trim(cmbYer)
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPrd)
   If bSqlRows Then
      bGoodYear = True
      With rdoPrd
         lblPrd = "" & .Fields(0)
         If Val("" & .Fields(0)) = 0 Then
            lblStat = "Fiscal Periods Not Established For " & cmbYer & "."
         Else
            lblStat = "Edit Fiscal Periods For " & cmbYer & "."
         End If
      End With
   Else
      bGoodYear = False
      ' no fiscal year found,  add it ?
      sMsg = cmbYer & " Wasn't Found. Add The Fiscal Year?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         
         clsADOCon.ADOErrNum = 0
         
         sSql = "INSERT INTO GlfyTable (FYYEAR,FYPERIODS) VALUES(" & cmbYer & ",0)"
         clsADOCon.ExecuteSQL sSql
         If clsADOCon.ADOErrNum = 0 Then
            itempYr = cmbYer
            bGoodYear = True
            FillYearCombo
            cmbYer = itempYr
            GetPeriod
         Else
            sMsg = "Cannot Add Fiscal Year."
            MsgBox sMsg, vbExclamation, Caption
         End If
      End If
   End If
   Exit Sub
DiaErr1:
   sProcName = "GetPeriod"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub InitializeYears() 'Sets Up initial year data
   Dim a As Integer
   Dim i As Integer
   Dim iYear As Integer
   On Error GoTo DiaErr1
   MouseCursor 13
   iYear = Format(Now, "yyyy")
   
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   For i = 0 To 10
      sSql = "INSERT INTO GlfyTable (FYYEAR,FYSTART,FYEND,FYPERIODS)" _
             & " VALUES(" & iYear + i & ",NULL,NULL," _
             & "0)"
      clsADOCon.ExecuteSQL sSql
   Next
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      lblStat = "Fiscal Years Initialized."
      lblStat.Refresh
      Sleep 500
      MouseCursor 0
      lblStat = ""
      FillYearCombo
   Else
      MouseCursor 0
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MsgBox "Unable To Establish Fiscal Years.", _
         vbExclamation, Caption
   End If
   Exit Sub
DiaErr1:
   sProcName = "InitialIze"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Function CheckYears() As Byte 'checks if initialization is neccesary
   Dim RdoFyr As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT FYYEAR FROM GlfyTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFyr, ES_FORWARD)
   If bSqlRows Then
      CheckYears = 1
      RdoFyr.Cancel
   Else
      CheckYears = 0
   End If
   Set RdoFyr = Nothing
   sProcName = "checkyears"
End Function

Private Sub ValidDate(iType As Integer, iRow As Integer) 'Function Used to determine Whether or not
   Dim i As Integer
   
   On Error GoTo DiaErr1 'to update the dates. (Bad Dates?)
   For i = 0 To Val(lblPrd) - 1
      If i > 0 Then
         If Trim(txtsDte(i)) <> "" And Trim(txteDte(i)) <> "" Then
            If CDate(txtsDte(i)) > CDate(txteDte(i - 1)) Then
               If CDate(txteDte(i)) > CDate(txtsDte(i)) Then
                  iType = 4
                  iRow = i
               Else
                  iType = 1
                  iRow = i
                  Exit Sub
               End If
            Else
               iType = 2
               iRow = i
               Exit Sub
            End If
         Else
            iType = 3
            iRow = i
            Exit Sub
         End If
      Else
         If Trim(txteDte(i)) <> "" And Trim(txtsDte(i)) <> "" Then
            If CDate(txteDte(i)) > CDate(txtsDte(i)) Then
               iType = 4
               iRow = i
            Else
               iType = 1
               iRow = i
               Exit Sub
            End If
         Else
            iType = 3
            iRow = i
            Exit Sub
         End If
      End If
   Next
   Exit Sub
DiaErr1:
   sProcName = "ValidDate"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
