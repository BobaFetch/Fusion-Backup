VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESe01c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Volume Discounts"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtQto 
      Height          =   285
      Index           =   11
      Left            =   7080
      TabIndex        =   45
      Tag             =   "2"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtPer 
      Height          =   285
      Index           =   11
      Left            =   5400
      TabIndex        =   46
      Tag             =   "2"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Index           =   11
      Left            =   7080
      TabIndex        =   47
      Tag             =   "2"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtQfr 
      Height          =   285
      Index           =   11
      Left            =   5400
      TabIndex        =   44
      Tag             =   "2"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtQfr 
      Height          =   285
      Index           =   10
      Left            =   5400
      TabIndex        =   40
      Tag             =   "2"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Index           =   10
      Left            =   7080
      TabIndex        =   43
      Tag             =   "2"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtPer 
      Height          =   285
      Index           =   10
      Left            =   5400
      TabIndex        =   42
      Tag             =   "2"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox txtQto 
      Height          =   285
      Index           =   10
      Left            =   7080
      TabIndex        =   41
      Tag             =   "2"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtQfr 
      Height          =   285
      Index           =   9
      Left            =   5400
      TabIndex        =   6
      Tag             =   "2"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Index           =   9
      Left            =   7080
      TabIndex        =   39
      Tag             =   "2"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtPer 
      Height          =   285
      Index           =   9
      Left            =   5400
      TabIndex        =   38
      Tag             =   "2"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtQto 
      Height          =   285
      Index           =   9
      Left            =   7080
      TabIndex        =   37
      Tag             =   "2"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtQfr 
      Height          =   285
      Index           =   8
      Left            =   5400
      TabIndex        =   33
      Tag             =   "2"
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Index           =   8
      Left            =   7080
      TabIndex        =   36
      Tag             =   "2"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtPer 
      Height          =   285
      Index           =   8
      Left            =   5400
      TabIndex        =   35
      Tag             =   "2"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtQto 
      Height          =   285
      Index           =   8
      Left            =   7080
      TabIndex        =   34
      Tag             =   "2"
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtQfr 
      Height          =   285
      Index           =   7
      Left            =   5400
      TabIndex        =   29
      Tag             =   "2"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Index           =   7
      Left            =   7080
      TabIndex        =   32
      Tag             =   "2"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtPer 
      Height          =   285
      Index           =   7
      Left            =   5400
      TabIndex        =   31
      Tag             =   "2"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtQto 
      Height          =   285
      Index           =   7
      Left            =   7080
      TabIndex        =   30
      Tag             =   "2"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtQto 
      Height          =   285
      Index           =   6
      Left            =   7080
      TabIndex        =   26
      Tag             =   "2"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtPer 
      Height          =   285
      Index           =   6
      Left            =   5400
      TabIndex        =   27
      Tag             =   "2"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Index           =   6
      Left            =   7080
      TabIndex        =   28
      Tag             =   "2"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtQfr 
      Height          =   285
      Index           =   6
      Left            =   5400
      TabIndex        =   25
      Tag             =   "2"
      Top             =   720
      Width           =   735
   End
   Begin VB.Frame z2 
      ForeColor       =   &H80000010&
      Height          =   5175
      Left            =   4200
      TabIndex        =   79
      Top             =   600
      Width           =   20
   End
   Begin VB.CheckBox optfrom 
      Caption         =   "from full"
      Height          =   255
      Left            =   5280
      TabIndex        =   75
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtQto 
      Height          =   285
      Index           =   5
      Left            =   2880
      TabIndex        =   22
      Tag             =   "2"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtPer 
      Height          =   285
      Index           =   5
      Left            =   1200
      TabIndex        =   23
      Tag             =   "2"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Index           =   5
      Left            =   2880
      TabIndex        =   24
      Tag             =   "2"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtQfr 
      Height          =   285
      Index           =   5
      Left            =   1200
      TabIndex        =   21
      Tag             =   "2"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtQto 
      Height          =   285
      Index           =   4
      Left            =   2880
      TabIndex        =   18
      Tag             =   "2"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtPer 
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   19
      Tag             =   "2"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Index           =   4
      Left            =   2880
      TabIndex        =   20
      Tag             =   "2"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtQfr 
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   17
      Tag             =   "2"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtQto 
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   14
      Tag             =   "2"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtPer 
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   15
      Tag             =   "2"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   16
      Tag             =   "2"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtQfr 
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   13
      Tag             =   "2"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtQto 
      Height          =   285
      Index           =   2
      Left            =   2880
      TabIndex        =   10
      Tag             =   "2"
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtPer 
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   11
      Tag             =   "2"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Index           =   2
      Left            =   2880
      TabIndex        =   12
      Tag             =   "2"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtQfr 
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   9
      Tag             =   "2"
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtQfr 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Tag             =   "2"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   8
      Tag             =   "2"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtPer 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Tag             =   "2"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtQto 
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   5
      Tag             =   "2"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtQfr 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Tag             =   "2"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   3
      Tag             =   "2"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtPer 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtQto 
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   1
      Tag             =   "2"
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7440
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   2280
      Top             =   5640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5880
      FormDesignWidth =   8475
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From 10"
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
      Index           =   50
      Left            =   4440
      TabIndex        =   103
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Index           =   49
      Left            =   6240
      TabIndex        =   102
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   255
      Index           =   48
      Left            =   6240
      TabIndex        =   101
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      Height          =   255
      Index           =   47
      Left            =   4440
      TabIndex        =   100
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      Height          =   255
      Index           =   46
      Left            =   4440
      TabIndex        =   99
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   255
      Index           =   45
      Left            =   6240
      TabIndex        =   98
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Index           =   44
      Left            =   6240
      TabIndex        =   97
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From 12"
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
      Index           =   43
      Left            =   4440
      TabIndex        =   96
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      Height          =   255
      Index           =   42
      Left            =   4440
      TabIndex        =   95
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   255
      Index           =   41
      Left            =   6240
      TabIndex        =   94
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Index           =   40
      Left            =   6240
      TabIndex        =   93
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From 9"
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
      Index           =   39
      Left            =   4440
      TabIndex        =   92
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      Height          =   255
      Index           =   38
      Left            =   4440
      TabIndex        =   91
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   255
      Index           =   37
      Left            =   6240
      TabIndex        =   90
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Index           =   36
      Left            =   6240
      TabIndex        =   89
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From 11"
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
      Index           =   35
      Left            =   4440
      TabIndex        =   88
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      Height          =   255
      Index           =   34
      Left            =   4440
      TabIndex        =   87
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   255
      Index           =   33
      Left            =   6240
      TabIndex        =   86
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Index           =   32
      Left            =   6240
      TabIndex        =   85
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From 8"
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
      Index           =   31
      Left            =   4440
      TabIndex        =   84
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From 7"
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
      Index           =   30
      Left            =   4440
      TabIndex        =   83
      Top             =   720
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Index           =   29
      Left            =   6240
      TabIndex        =   82
      Top             =   720
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   255
      Index           =   28
      Left            =   6240
      TabIndex        =   81
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      Height          =   255
      Index           =   27
      Left            =   4440
      TabIndex        =   80
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price"
      Height          =   255
      Index           =   25
      Left            =   3000
      TabIndex        =   78
      Top             =   0
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate"
      Height          =   255
      Index           =   24
      Left            =   240
      TabIndex        =   77
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblBid 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1200
      TabIndex        =   76
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label lblPrice 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   74
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From 6"
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
      Index           =   23
      Left            =   240
      TabIndex        =   73
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Index           =   22
      Left            =   2040
      TabIndex        =   72
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   255
      Index           =   21
      Left            =   2040
      TabIndex        =   71
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      Height          =   255
      Index           =   20
      Left            =   240
      TabIndex        =   70
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From 3"
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
      Index           =   19
      Left            =   240
      TabIndex        =   69
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Index           =   18
      Left            =   2040
      TabIndex        =   68
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   255
      Index           =   17
      Left            =   2040
      TabIndex        =   67
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   66
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From 5"
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
      Index           =   15
      Left            =   240
      TabIndex        =   65
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Index           =   14
      Left            =   2040
      TabIndex        =   64
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   255
      Index           =   13
      Left            =   2040
      TabIndex        =   63
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   62
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From 2"
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
      Index           =   11
      Left            =   240
      TabIndex        =   61
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   60
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   255
      Index           =   9
      Left            =   2040
      TabIndex        =   59
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   58
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   57
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   56
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   55
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From 4"
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
      Left            =   240
      TabIndex        =   54
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   53
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   52
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   51
      Top             =   720
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From 1"
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
      Left            =   240
      TabIndex        =   50
      Top             =   720
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Quantities/Prices"
      Height          =   255
      Index           =   26
      Left            =   240
      TabIndex        =   49
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "EstiESe01c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      If optFrom.Value = vbChecked Then
         lblPrice = EstiESe02a.lblUnitTotal
         lblBid = EstiESe02a.cmbCls & EstiESe02a.cmbBid _
                  & "-" & EstiESe02a.cmbCst
      Else
         lblPrice = EstiESe01a.txtPrc
         lblBid = EstiESe01a.cmbCls & EstiESe01a.cmbBid _
                  & "-" & EstiESe01a.cmbCst
      End If
      GetPrices
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   Move 600, 1300
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If optFrom.Value = vbChecked Then
      EstiESe02a.optFrom = vbUnchecked
   Else
      EstiESe01a.optFrom = vbUnchecked
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set EstiESe01c = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
   For b = 0 To 11
      txtQfr(b) = "0"
      txtQto(b) = "0"
      txtPrc(b) = "0.00"
      txtPer(b) = "0.00"
   Next
'   txtQfr(b) = "0"
'   txtQto(b) = "0"
'   txtPrc(b) = "0.00"
'   txtPer(b) = "0.00"
   
End Sub




Private Function GetPrices()
   On Error GoTo DiaErr1
   If RdoFull Is Nothing Then Exit Function
   If optFrom.Value = vbChecked Then
      With RdoFull
         txtQfr(0) = Format(!BIDQTYFROM1, "####0")
         txtQfr(1) = Format(!BIDQTYFROM2, "####0")
         txtQfr(2) = Format(!BIDQTYFROM3, "####0")
         txtQfr(3) = Format(!BIDQTYFROM4, "####0")
         txtQfr(4) = Format(!BIDQTYFROM5, "####0")
         txtQfr(5) = Format(!BIDQTYFROM6, "####0")
         txtQfr(6) = Format(!BIDQTYFROM7, "####0")
         txtQfr(7) = Format(!BIDQTYFROM8, "####0")
         txtQfr(8) = Format(!BIDQTYFROM9, "####0")
         txtQfr(9) = Format(!BIDQTYFROM10, "####0")
         txtQfr(10) = Format(!BIDQTYFROM11, "####0")
         txtQfr(11) = Format(!BIDQTYFROM12, "####0")
         
         txtQto(0) = Format(!BIDQTYTO1, "####0")
         txtQto(1) = Format(!BIDQTYTO2, "####0")
         txtQto(2) = Format(!BIDQTYTO3, "####0")
         txtQto(3) = Format(!BIDQTYTO4, "####0")
         txtQto(4) = Format(!BIDQTYTO5, "####0")
         txtQto(5) = Format(!BIDQTYTO6, "####0")
         txtQto(6) = Format(!BIDQTYTO7, "####0")
         txtQto(7) = Format(!BIDQTYTO8, "####0")
         txtQto(8) = Format(!BIDQTYTO9, "####0")
         txtQto(9) = Format(!BIDQTYTO10, "####0")
         txtQto(10) = Format(!BIDQTYTO11, "####0")
         txtQto(11) = Format(!BIDQTYTO12, "####0")
         
         txtPrc(0) = Format(!BIDQTYPRICE1, "####0.00")
         txtPrc(1) = Format(!BIDQTYPRICE2, "####0.00")
         txtPrc(2) = Format(!BIDQTYPRICE3, "####0.00")
         txtPrc(3) = Format(!BIDQTYPRICE4, "####0.00")
         txtPrc(4) = Format(!BIDQTYPRICE5, "####0.00")
         txtPrc(5) = Format(!BIDQTYPRICE6, "####0.00")
         txtPrc(6) = Format(!BIDQTYPRICE7, "####0.00")
         txtPrc(7) = Format(!BIDQTYPRICE8, "####0.00")
         txtPrc(8) = Format(!BIDQTYPRICE9, "####0.00")
         txtPrc(9) = Format(!BIDQTYPRICE10, "####0.00")
         txtPrc(10) = Format(!BIDQTYPRICE11, "####0.00")
         txtPrc(11) = Format(!BIDQTYPRICE12, "####0.00")
         
         txtPer(0) = Format(!BIDQTYDISC1, "####0.00")
         txtPer(1) = Format(!BIDQTYDISC2, "####0.00")
         txtPer(2) = Format(!BIDQTYDISC3, "####0.00")
         txtPer(3) = Format(!BIDQTYDISC4, "####0.00")
         txtPer(4) = Format(!BIDQTYDISC5, "####0.00")
         txtPer(5) = Format(!BIDQTYDISC6, "####0.00")
         txtPer(6) = Format(!BIDQTYDISC7, "####0.00")
         txtPer(7) = Format(!BIDQTYDISC8, "####0.00")
         txtPer(8) = Format(!BIDQTYDISC9, "####0.00")
         txtPer(9) = Format(!BIDQTYDISC10, "####0.00")
         txtPer(10) = Format(!BIDQTYDISC11, "####0.00")
         txtPer(11) = Format(!BIDQTYDISC12, "####0.00")
         
      End With
   Else
      With RdoBid
         txtQfr(0) = Format(!BIDQTYFROM1, "####0")
         txtQfr(1) = Format(!BIDQTYFROM2, "####0")
         txtQfr(2) = Format(!BIDQTYFROM3, "####0")
         txtQfr(3) = Format(!BIDQTYFROM4, "####0")
         txtQfr(4) = Format(!BIDQTYFROM5, "####0")
         txtQfr(5) = Format(!BIDQTYFROM6, "####0")
         txtQfr(6) = Format(!BIDQTYFROM7, "####0")
         txtQfr(7) = Format(!BIDQTYFROM8, "####0")
         txtQfr(8) = Format(!BIDQTYFROM9, "####0")
         txtQfr(9) = Format(!BIDQTYFROM10, "####0")
         txtQfr(10) = Format(!BIDQTYFROM11, "####0")
         txtQfr(11) = Format(!BIDQTYFROM12, "####0")
         
         txtQto(0) = Format(!BIDQTYTO1, "####0")
         txtQto(1) = Format(!BIDQTYTO2, "####0")
         txtQto(2) = Format(!BIDQTYTO3, "####0")
         txtQto(3) = Format(!BIDQTYTO4, "####0")
         txtQto(4) = Format(!BIDQTYTO5, "####0")
         txtQto(5) = Format(!BIDQTYTO6, "####0")
         txtQto(6) = Format(!BIDQTYTO7, "####0")
         txtQto(7) = Format(!BIDQTYTO8, "####0")
         txtQto(8) = Format(!BIDQTYTO9, "####0")
         txtQto(9) = Format(!BIDQTYTO10, "####0")
         txtQto(10) = Format(!BIDQTYTO11, "####0")
         txtQto(11) = Format(!BIDQTYTO12, "####0")
         
         txtPrc(0) = Format(!BIDQTYPRICE1, "####0.00")
         txtPrc(1) = Format(!BIDQTYPRICE2, "####0.00")
         txtPrc(2) = Format(!BIDQTYPRICE3, "####0.00")
         txtPrc(3) = Format(!BIDQTYPRICE4, "####0.00")
         txtPrc(4) = Format(!BIDQTYPRICE5, "####0.00")
         txtPrc(5) = Format(!BIDQTYPRICE6, "####0.00")
         txtPrc(6) = Format(!BIDQTYPRICE7, "####0.00")
         txtPrc(7) = Format(!BIDQTYPRICE8, "####0.00")
         txtPrc(8) = Format(!BIDQTYPRICE9, "####0.00")
         txtPrc(9) = Format(!BIDQTYPRICE10, "####0.00")
         txtPrc(10) = Format(!BIDQTYPRICE11, "####0.00")
         txtPrc(11) = Format(!BIDQTYPRICE12, "####0.00")
         
         txtPer(0) = Format(!BIDQTYDISC1, "####0.00")
         txtPer(1) = Format(!BIDQTYDISC2, "####0.00")
         txtPer(2) = Format(!BIDQTYDISC3, "####0.00")
         txtPer(3) = Format(!BIDQTYDISC4, "####0.00")
         txtPer(4) = Format(!BIDQTYDISC5, "####0.00")
         txtPer(5) = Format(!BIDQTYDISC6, "####0.00")
         txtPer(6) = Format(!BIDQTYDISC7, "####0.00")
         txtPer(7) = Format(!BIDQTYDISC8, "####0.00")
         txtPer(8) = Format(!BIDQTYDISC9, "####0.00")
         txtPer(9) = Format(!BIDQTYDISC10, "####0.00")
         txtPer(10) = Format(!BIDQTYDISC11, "####0.00")
         txtPer(11) = Format(!BIDQTYDISC12, "####0.00")
      End With
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getprices"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtPer_Change(Index As Integer)
   If Val(txtPer(Index)) <= 0 Then txtPer(Index) = "0.00"
   
End Sub

Private Sub txtPer_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtPer_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtPer_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtPer_LostFocus(Index As Integer)
   Dim cPer As Currency
   Dim cPrc As Currency
   txtPer(Index) = CheckLen(txtPer(Index), 6)
   txtPer(Index) = Format(Abs(Val(txtPer(Index))), "#0.00")
   cPer = Val(txtPer(Index)) / 100
   cPrc = Val(lblPrice)
   If cPrc > 0 Then
      If (cPer > 1) Then
         cPrc = cPrc * cPer
      Else
         cPrc = (cPrc * (1 - cPer))
      End If
      txtPrc(Index) = Format(cPrc, "####0.00")
   Else
      If Index = 0 Then txtPrc(Index) = lblPrice
   End If
   On Error Resume Next
   If optFrom.Value = vbChecked Then
      With RdoFull
         '.Edit
         Select Case Index
            Case 0
               !BIDQTYDISC1 = Val(txtPer(Index))
               !BIDQTYPRICE1 = Val(txtPrc(Index))
            Case 1
               !BIDQTYDISC2 = Val(txtPer(Index))
               !BIDQTYPRICE2 = Val(txtPrc(Index))
            Case 2
               !BIDQTYDISC3 = Val(txtPer(Index))
               !BIDQTYPRICE3 = Val(txtPrc(Index))
            Case 3
               !BIDQTYDISC4 = Val(txtPer(Index))
               !BIDQTYPRICE4 = Val(txtPrc(Index))
            Case 4
               !BIDQTYDISC5 = Val(txtPer(Index))
               !BIDQTYPRICE5 = Val(txtPrc(Index))
            Case 5
               !BIDQTYDISC6 = Val(txtPer(Index))
               !BIDQTYPRICE6 = Val(txtPrc(Index))
            Case 6
               !BIDQTYDISC7 = Val(txtPer(Index))
               !BIDQTYPRICE7 = Val(txtPrc(Index))
            Case 7
               !BIDQTYDISC8 = Val(txtPer(Index))
               !BIDQTYPRICE8 = Val(txtPrc(Index))
            Case 8
               !BIDQTYDISC9 = Val(txtPer(Index))
               !BIDQTYPRICE9 = Val(txtPrc(Index))
            Case 9
               !BIDQTYDISC10 = Val(txtPer(Index))
               !BIDQTYPRICE10 = Val(txtPrc(Index))
            Case 10
               !BIDQTYDISC11 = Val(txtPer(Index))
               !BIDQTYPRICE11 = Val(txtPrc(Index))
            Case Else
               !BIDQTYDISC12 = Val(txtPer(Index))
               !BIDQTYPRICE12 = Val(txtPrc(Index))
         End Select
         .Update
      End With
   Else
      With RdoBid
         '.Edit
         Select Case Index
            Case 0
               !BIDQTYDISC1 = Val(txtPer(Index))
               !BIDQTYPRICE1 = Val(txtPrc(Index))
            Case 1
               !BIDQTYDISC2 = Val(txtPer(Index))
               !BIDQTYPRICE2 = Val(txtPrc(Index))
            Case 2
               !BIDQTYDISC3 = Val(txtPer(Index))
               !BIDQTYPRICE3 = Val(txtPrc(Index))
            Case 3
               !BIDQTYDISC4 = Val(txtPer(Index))
               !BIDQTYPRICE4 = Val(txtPrc(Index))
            Case 4
               !BIDQTYDISC5 = Val(txtPer(Index))
               !BIDQTYPRICE5 = Val(txtPrc(Index))
            Case 5
               !BIDQTYDISC6 = Val(txtPer(Index))
               !BIDQTYPRICE6 = Val(txtPrc(Index))
            Case 6
               !BIDQTYDISC7 = Val(txtPer(Index))
               !BIDQTYPRICE7 = Val(txtPrc(Index))
            Case 7
               !BIDQTYDISC8 = Val(txtPer(Index))
               !BIDQTYPRICE8 = Val(txtPrc(Index))
            Case 8
               !BIDQTYDISC9 = Val(txtPer(Index))
               !BIDQTYPRICE9 = Val(txtPrc(Index))
            Case 9
               !BIDQTYDISC10 = Val(txtPer(Index))
               !BIDQTYPRICE10 = Val(txtPrc(Index))
            Case 10
               !BIDQTYDISC11 = Val(txtPer(Index))
               !BIDQTYPRICE11 = Val(txtPrc(Index))
            Case Else
               !BIDQTYDISC12 = Val(txtPer(Index))
               !BIDQTYPRICE12 = Val(txtPrc(Index))
         End Select
         .Update
      End With
   End If
   
End Sub

Private Sub txtPer_Validate(Index As Integer, Cancel As Boolean)
   Debug.Print "validate"
End Sub

Private Sub txtPrc_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtPrc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtPrc_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtPrc_LostFocus(Index As Integer)
   Dim cPer As Currency
   Dim cBPrc As Currency
   Dim cDPrc As Currency
   Dim cValPer As Currency
   
   txtPrc(Index) = CheckLen(txtPrc(Index), 10)
   txtPrc(Index) = Format(Abs(Val(txtPrc(Index))), "####0.00")
   cPer = Val(txtPrc(Index))
   cBPrc = Val(lblPrice)
   cDPrc = Val(txtPrc(Index))
   If cBPrc > 0 And cDPrc > 0 Then
      cValPer = cDPrc / cBPrc
      If (cValPer > 1) Then
         cPer = cDPrc / cBPrc
      Else
         cPer = 1 - (cDPrc / cBPrc)
      End If
      txtPer(Index) = Format(cPer * 100, "#0.00")
   Else
      If cPer <= 0 Then
         If Index = 0 Then
            txtPrc(Index) = lblPrice
            txtPrc(Index) = "0.00"
         End If
      End If
   End If
   On Error Resume Next
   If optFrom.Value = vbChecked Then
      With RdoFull
         '.Edit
         Select Case Index
            Case 0
               !BIDQTYPRICE1 = Val(txtPrc(Index))
               !BIDQTYDISC1 = Val(txtPer(Index))
            Case 1
               !BIDQTYPRICE2 = Val(txtPrc(Index))
               !BIDQTYDISC2 = Val(txtPer(Index))
            Case 2
               !BIDQTYPRICE3 = Val(txtPrc(Index))
               !BIDQTYDISC3 = Val(txtPer(Index))
            Case 3
               !BIDQTYPRICE4 = Val(txtPrc(Index))
               !BIDQTYDISC4 = Val(txtPer(Index))
            Case 4
               !BIDQTYPRICE5 = Val(txtPrc(Index))
               !BIDQTYDISC5 = Val(txtPer(Index))
            Case 5
               !BIDQTYPRICE6 = Val(txtPrc(Index))
               !BIDQTYDISC6 = Val(txtPer(Index))
            Case 6
               !BIDQTYPRICE7 = Val(txtPrc(Index))
               !BIDQTYDISC7 = Val(txtPer(Index))
            Case 7
               !BIDQTYPRICE8 = Val(txtPrc(Index))
               !BIDQTYDISC8 = Val(txtPer(Index))
            Case 8
               !BIDQTYPRICE9 = Val(txtPrc(Index))
               !BIDQTYDISC9 = Val(txtPer(Index))
            Case 9
               !BIDQTYPRICE10 = Val(txtPrc(Index))
               !BIDQTYDISC10 = Val(txtPer(Index))
            Case 10
               !BIDQTYPRICE11 = Val(txtPrc(Index))
               !BIDQTYDISC11 = Val(txtPer(Index))
            Case Else
               !BIDQTYPRICE12 = Val(txtPrc(Index))
               !BIDQTYDISC12 = Val(txtPer(Index))
         End Select
         .Update
      End With
   Else
      With RdoBid
         '.Edit
         Select Case Index
            Case 0
               !BIDQTYPRICE1 = Val(txtPrc(Index))
               !BIDQTYDISC1 = Val(txtPer(Index))
            Case 1
               !BIDQTYPRICE2 = Val(txtPrc(Index))
               !BIDQTYDISC2 = Val(txtPer(Index))
            Case 2
               !BIDQTYPRICE3 = Val(txtPrc(Index))
               !BIDQTYDISC3 = Val(txtPer(Index))
            Case 3
               !BIDQTYPRICE4 = Val(txtPrc(Index))
               !BIDQTYDISC4 = Val(txtPer(Index))
            Case 4
               !BIDQTYPRICE5 = Val(txtPrc(Index))
               !BIDQTYDISC5 = Val(txtPer(Index))
            Case 5
               !BIDQTYPRICE6 = Val(txtPrc(Index))
               !BIDQTYDISC6 = Val(txtPer(Index))
            Case 6
               !BIDQTYPRICE7 = Val(txtPrc(Index))
               !BIDQTYDISC7 = Val(txtPer(Index))
            Case 7
               !BIDQTYPRICE8 = Val(txtPrc(Index))
               !BIDQTYDISC8 = Val(txtPer(Index))
            Case 8
               !BIDQTYPRICE9 = Val(txtPrc(Index))
               !BIDQTYDISC9 = Val(txtPer(Index))
            Case 9
               !BIDQTYPRICE10 = Val(txtPrc(Index))
               !BIDQTYDISC10 = Val(txtPer(Index))
            Case 10
               !BIDQTYPRICE11 = Val(txtPrc(Index))
               !BIDQTYDISC11 = Val(txtPer(Index))
            Case Else
               !BIDQTYPRICE12 = Val(txtPrc(Index))
               !BIDQTYDISC12 = Val(txtPer(Index))
         End Select
         .Update
      End With
   End If
   
End Sub

Private Sub txtQfr_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtQfr_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtQfr_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtQfr_LostFocus(Index As Integer)
   txtQfr(Index) = CheckLen(txtQfr(Index), 6)
   txtQfr(Index) = Format(Abs(Val(txtQfr(Index))), "####0")
   On Error Resume Next
   If optFrom.Value = vbChecked Then
      With RdoFull
         '.Edit
         Select Case Index
            Case 0
               !BIDQTYFROM1 = Val(txtQfr(Index))
            Case 1
               !BIDQTYFROM2 = Val(txtQfr(Index))
            Case 2
               !BIDQTYFROM3 = Val(txtQfr(Index))
            Case 3
               !BIDQTYFROM4 = Val(txtQfr(Index))
            Case 4
               !BIDQTYFROM5 = Val(txtQfr(Index))
            Case 5
               !BIDQTYFROM6 = Val(txtQfr(Index))
            Case 6
               !BIDQTYFROM7 = Val(txtQfr(Index))
            Case 7
               !BIDQTYFROM8 = Val(txtQfr(Index))
            Case 8
               !BIDQTYFROM9 = Val(txtQfr(Index))
            Case 9
               !BIDQTYFROM10 = Val(txtQfr(Index))
            Case 10
               !BIDQTYFROM11 = Val(txtQfr(Index))
            Case Else
               !BIDQTYFROM12 = Val(txtQfr(Index))
         End Select
         .Update
      End With
   Else
      With RdoBid
         '.Edit
         Select Case Index
            Case 0
               !BIDQTYFROM1 = Val(txtQfr(Index))
            Case 1
               !BIDQTYFROM2 = Val(txtQfr(Index))
            Case 2
               !BIDQTYFROM3 = Val(txtQfr(Index))
            Case 3
               !BIDQTYFROM4 = Val(txtQfr(Index))
            Case 4
               !BIDQTYFROM5 = Val(txtQfr(Index))
            Case 5
               !BIDQTYFROM6 = Val(txtQfr(Index))
            Case 6
               !BIDQTYFROM7 = Val(txtQfr(Index))
            Case 7
               !BIDQTYFROM8 = Val(txtQfr(Index))
            Case 8
               !BIDQTYFROM9 = Val(txtQfr(Index))
            Case 9
               !BIDQTYFROM10 = Val(txtQfr(Index))
            Case 10
               !BIDQTYFROM11 = Val(txtQfr(Index))
            Case Else
               !BIDQTYFROM12 = Val(txtQfr(Index))
         End Select
         .Update
      End With
   End If
   
End Sub

Private Sub txtQto_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtQto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtQto_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtQto_LostFocus(Index As Integer)
   txtQto(Index) = CheckLen(txtQto(Index), 6)
   txtQto(Index) = Format(Abs(Val(txtQto(Index))), "####0")
   On Error Resume Next
   If optFrom.Value = vbChecked Then
      With RdoFull
         '.Edit
         Select Case Index
            Case 0
               !BIDQTYTO1 = Val(txtQto(Index))
            Case 1
               !BIDQTYTO2 = Val(txtQto(Index))
            Case 2
               !BIDQTYTO3 = Val(txtQto(Index))
            Case 3
               !BIDQTYTO4 = Val(txtQto(Index))
            Case 4
               !BIDQTYTO5 = Val(txtQto(Index))
            Case 5
               !BIDQTYTO6 = Val(txtQto(Index))
            Case 6
               !BIDQTYTO7 = Val(txtQto(Index))
            Case 7
               !BIDQTYTO8 = Val(txtQto(Index))
            Case 8
               !BIDQTYTO9 = Val(txtQto(Index))
            Case 9
               !BIDQTYTO10 = Val(txtQto(Index))
            Case 10
               !BIDQTYTO11 = Val(txtQto(Index))
            Case Else
               !BIDQTYTO12 = Val(txtQto(Index))
         End Select
         .Update
      End With
   Else
      With RdoBid
         '.Edit
         Select Case Index
            Case 0
               !BIDQTYTO1 = Val(txtQto(Index))
            Case 1
               !BIDQTYTO2 = Val(txtQto(Index))
            Case 2
               !BIDQTYTO3 = Val(txtQto(Index))
            Case 3
               !BIDQTYTO4 = Val(txtQto(Index))
            Case 4
               !BIDQTYTO5 = Val(txtQto(Index))
            Case 5
               !BIDQTYTO6 = Val(txtQto(Index))
            Case 6
               !BIDQTYTO7 = Val(txtQto(Index))
            Case 7
               !BIDQTYTO8 = Val(txtQto(Index))
            Case 8
               !BIDQTYTO9 = Val(txtQto(Index))
            Case 9
               !BIDQTYTO10 = Val(txtQto(Index))
            Case 10
               !BIDQTYTO11 = Val(txtQto(Index))
            Case Else
               !BIDQTYTO12 = Val(txtQto(Index))
         End Select
         .Update
      End With
   End If
   
End Sub


