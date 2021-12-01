VERSION 5.00
Begin VB.Form CapaCPp10b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work Center Load Analysis"
   ClientHeight    =   8400
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   8280
   Icon            =   "CapaCPp10b.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPp10b.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   87
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton optPrn 
      DownPicture     =   "CapaCPp10b.frx":0AB8
      Height          =   415
      Left            =   7560
      Picture         =   "CapaCPp10b.frx":0C42
      Style           =   1  'Graphical
      TabIndex        =   78
      TabStop         =   0   'False
      ToolTipText     =   "Print This Form"
      Top             =   120
      Width           =   415
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   3720
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   8640
      Width           =   1065
   End
   Begin VB.Label CntAvailable 
      BackStyle       =   0  'Transparent
      Caption         =   "Available: "
      Height          =   252
      Index           =   9
      Left            =   5280
      TabIndex        =   86
      Top             =   7560
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRequired 
      BackStyle       =   0  'Transparent
      Caption         =   "Required: "
      Height          =   252
      Index           =   9
      Left            =   3840
      TabIndex        =   85
      Top             =   7560
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label FreeHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   9
      Left            =   4884
      TabIndex        =   84
      ToolTipText     =   "Calendar Hours Recorded"
      Top             =   7320
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label UsedHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   9
      Left            =   1884
      TabIndex        =   83
      ToolTipText     =   "Hours Required For Setup And Cycle Times"
      Top             =   7320
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label WorkCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   9
      Left            =   120
      TabIndex        =   82
      ToolTipText     =   "Work Center"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label WCDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   9
      Left            =   360
      TabIndex        =   81
      ToolTipText     =   "Work Center"
      Top             =   7584
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.Label CntRemaining 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining:"
      Height          =   252
      Index           =   9
      Left            =   6600
      TabIndex        =   80
      Top             =   7560
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center Load Analysis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   79
      Top             =   120
      Width           =   7092
   End
   Begin VB.Label Shop4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remain"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   6600
      TabIndex        =   77
      ToolTipText     =   "Work Center Calendar Hours Less MO Operation Requirements"
      Top             =   7920
      Width           =   1332
   End
   Begin VB.Label Shop3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Avail"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5040
      TabIndex        =   76
      ToolTipText     =   "Totals From Work Center Calendars"
      Top             =   7920
      Width           =   1452
   End
   Begin VB.Label Shop2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reqired"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3480
      TabIndex        =   75
      ToolTipText     =   "Required (MO Operations) Hours"
      Top             =   7920
      Width           =   1452
   End
   Begin VB.Label Shop1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shop Totals:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   74
      Top             =   7920
      Width           =   1572
   End
   Begin VB.Label CntRemaining 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining:"
      Height          =   252
      Index           =   8
      Left            =   6600
      TabIndex        =   72
      Top             =   6960
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRemaining 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining:"
      Height          =   252
      Index           =   7
      Left            =   6600
      TabIndex        =   71
      Top             =   6360
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRemaining 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining:"
      Height          =   252
      Index           =   6
      Left            =   6600
      TabIndex        =   70
      Top             =   5760
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRemaining 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining:"
      Height          =   252
      Index           =   5
      Left            =   6600
      TabIndex        =   69
      Top             =   5040
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRemaining 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining:"
      Height          =   252
      Index           =   4
      Left            =   6600
      TabIndex        =   68
      Top             =   4320
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRemaining 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining:"
      Height          =   252
      Index           =   3
      Left            =   6600
      TabIndex        =   67
      Top             =   3600
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRemaining 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining:"
      Height          =   252
      Index           =   2
      Left            =   6600
      TabIndex        =   66
      Top             =   2880
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRemaining 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining:"
      Height          =   252
      Index           =   1
      Left            =   6600
      TabIndex        =   65
      Top             =   2160
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRemaining 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining:"
      Height          =   252
      Index           =   0
      Left            =   6600
      TabIndex        =   64
      Top             =   1440
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Center 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   6360
      TabIndex        =   63
      ToolTipText     =   "Work Center(s)"
      Top             =   840
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   252
      Index           =   5
      Left            =   5280
      TabIndex        =   62
      ToolTipText     =   "From Work Center Calendars"
      Top             =   840
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   252
      Index           =   4
      Left            =   3240
      TabIndex        =   61
      ToolTipText     =   "From Work Center Calendars"
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Shop 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   3960
      TabIndex        =   60
      ToolTipText     =   "Shop Selected"
      Top             =   840
      Width           =   1092
   End
   Begin VB.Label lblThrough 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   2280
      TabIndex        =   59
      ToolTipText     =   "Date Selected"
      Top             =   840
      Width           =   732
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Though"
      Height          =   252
      Index           =   3
      Left            =   1560
      TabIndex        =   58
      ToolTipText     =   "From Work Center Calendars"
      Top             =   840
      Width           =   612
   End
   Begin VB.Label lblToday 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   720
      TabIndex        =   57
      ToolTipText     =   "Always Today"
      Top             =   840
      Width           =   732
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From "
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   56
      ToolTipText     =   "From Work Center Calendars"
      Top             =   840
      Width           =   612
   End
   Begin VB.Label WCDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   8
      Left            =   360
      TabIndex        =   55
      ToolTipText     =   "Work Center"
      Top             =   6984
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.Label WorkCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   8
      Left            =   120
      TabIndex        =   54
      ToolTipText     =   "Work Center"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label UsedHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   8
      Left            =   1884
      TabIndex        =   53
      ToolTipText     =   "Hours Required For Setup And Cycle Times"
      Top             =   6720
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label FreeHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   8
      Left            =   4884
      TabIndex        =   52
      ToolTipText     =   "Calendar Hours Recorded"
      Top             =   6720
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label CntRequired 
      BackStyle       =   0  'Transparent
      Caption         =   "Required: "
      Height          =   252
      Index           =   8
      Left            =   3840
      TabIndex        =   51
      Top             =   6960
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label CntAvailable 
      BackStyle       =   0  'Transparent
      Caption         =   "Available: "
      Height          =   252
      Index           =   8
      Left            =   5280
      TabIndex        =   50
      Top             =   6960
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntAvailable 
      BackStyle       =   0  'Transparent
      Caption         =   "Available: "
      Height          =   252
      Index           =   7
      Left            =   5280
      TabIndex        =   49
      Top             =   6360
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRequired 
      BackStyle       =   0  'Transparent
      Caption         =   "Required: "
      Height          =   252
      Index           =   7
      Left            =   3840
      TabIndex        =   48
      Top             =   6360
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label CntAvailable 
      BackStyle       =   0  'Transparent
      Caption         =   "Available: "
      Height          =   252
      Index           =   6
      Left            =   5280
      TabIndex        =   47
      Top             =   5760
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRequired 
      BackStyle       =   0  'Transparent
      Caption         =   "Required: "
      Height          =   252
      Index           =   6
      Left            =   3840
      TabIndex        =   46
      Top             =   5760
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label CntAvailable 
      BackStyle       =   0  'Transparent
      Caption         =   "Available: "
      Height          =   252
      Index           =   5
      Left            =   5280
      TabIndex        =   45
      Top             =   5040
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRequired 
      BackStyle       =   0  'Transparent
      Caption         =   "Required: "
      Height          =   252
      Index           =   5
      Left            =   3840
      TabIndex        =   44
      Top             =   5040
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label CntAvailable 
      BackStyle       =   0  'Transparent
      Caption         =   "Available: "
      Height          =   252
      Index           =   4
      Left            =   5280
      TabIndex        =   43
      Top             =   4320
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRequired 
      BackStyle       =   0  'Transparent
      Caption         =   "Required: "
      Height          =   252
      Index           =   4
      Left            =   3840
      TabIndex        =   42
      Top             =   4320
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label CntAvailable 
      BackStyle       =   0  'Transparent
      Caption         =   "Available: "
      Height          =   252
      Index           =   3
      Left            =   5280
      TabIndex        =   41
      Top             =   3600
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRequired 
      BackStyle       =   0  'Transparent
      Caption         =   "Required: "
      Height          =   252
      Index           =   3
      Left            =   3840
      TabIndex        =   40
      Top             =   3600
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label CntAvailable 
      BackStyle       =   0  'Transparent
      Caption         =   "Available: "
      Height          =   252
      Index           =   2
      Left            =   5280
      TabIndex        =   39
      Top             =   2880
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRequired 
      BackStyle       =   0  'Transparent
      Caption         =   "Required: "
      Height          =   252
      Index           =   2
      Left            =   3720
      TabIndex        =   38
      Top             =   2880
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label CntAvailable 
      BackStyle       =   0  'Transparent
      Caption         =   "Available: "
      Height          =   252
      Index           =   1
      Left            =   5280
      TabIndex        =   37
      Top             =   2160
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRequired 
      BackStyle       =   0  'Transparent
      Caption         =   "Required: "
      Height          =   252
      Index           =   1
      Left            =   3720
      TabIndex        =   36
      Top             =   2160
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label CntAvailable 
      BackStyle       =   0  'Transparent
      Caption         =   "Available: "
      Height          =   252
      Index           =   0
      Left            =   5280
      TabIndex        =   35
      Top             =   1440
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CntRequired 
      BackStyle       =   0  'Transparent
      Caption         =   "Required: "
      Height          =   252
      Index           =   0
      Left            =   3720
      TabIndex        =   34
      Top             =   1440
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Calendar Hours Do Not Include Past Hours"
      Height          =   252
      Index           =   1
      Left            =   3240
      TabIndex        =   33
      ToolTipText     =   "From Work Center Calendars"
      Top             =   480
      Width           =   3372
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Required Times Include Past Due "
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   32
      ToolTipText     =   "From Manufacturing Orders"
      Top             =   480
      Width           =   3132
   End
   Begin VB.Label FreeHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   7
      Left            =   4884
      TabIndex        =   31
      ToolTipText     =   "Calendar Hours Recorded"
      Top             =   6120
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label UsedHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   7
      Left            =   1884
      TabIndex        =   30
      ToolTipText     =   "Hours Required For Setup And Cycle Times"
      Top             =   6120
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label WorkCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   7
      Left            =   120
      TabIndex        =   29
      ToolTipText     =   "Work Center"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label WCDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   7
      Left            =   360
      TabIndex        =   28
      ToolTipText     =   "Work Center"
      Top             =   6384
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.Label WCDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   6
      Left            =   360
      TabIndex        =   27
      ToolTipText     =   "Work Center"
      Top             =   5784
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.Label WorkCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   26
      ToolTipText     =   "Work Center"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label UsedHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   6
      Left            =   1884
      TabIndex        =   25
      ToolTipText     =   "Hours Required For Setup And Cycle Times"
      Top             =   5520
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label FreeHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   6
      Left            =   4884
      TabIndex        =   24
      ToolTipText     =   "Calendar Hours Recorded"
      Top             =   5520
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label WCDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   5
      Left            =   360
      TabIndex        =   23
      ToolTipText     =   "Work Center"
      Top             =   5064
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.Label WorkCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "Work Center"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label UsedHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   5
      Left            =   1884
      TabIndex        =   21
      ToolTipText     =   "Hours Required For Setup And Cycle Times"
      Top             =   4800
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label FreeHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   5
      Left            =   4884
      TabIndex        =   20
      ToolTipText     =   "Calendar Hours Recorded"
      Top             =   4800
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label WCDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   360
      TabIndex        =   19
      ToolTipText     =   "Work Center"
      Top             =   4344
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.Label WorkCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Work Center"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label UsedHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   4
      Left            =   1884
      TabIndex        =   17
      ToolTipText     =   "Hours Required For Setup And Cycle Times"
      Top             =   4080
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label FreeHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   4
      Left            =   4884
      TabIndex        =   16
      ToolTipText     =   "Calendar Hours Recorded"
      Top             =   4080
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label WCDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   360
      TabIndex        =   15
      ToolTipText     =   "Work Center"
      Top             =   3624
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.Label WorkCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Work Center"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label UsedHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   3
      Left            =   1884
      TabIndex        =   13
      ToolTipText     =   "Hours Required For Setup And Cycle Times"
      Top             =   3360
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label FreeHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   3
      Left            =   4884
      TabIndex        =   12
      ToolTipText     =   "Calendar Hours Recorded"
      Top             =   3360
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label WCDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   360
      TabIndex        =   11
      ToolTipText     =   "Work Center"
      Top             =   2904
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.Label WorkCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Work Center"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label UsedHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   2
      Left            =   1884
      TabIndex        =   9
      ToolTipText     =   "Hours Required For Setup And Cycle Times"
      Top             =   2640
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label FreeHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   2
      Left            =   4884
      TabIndex        =   8
      ToolTipText     =   "Calendar Hours Recorded"
      Top             =   2640
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label WCDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "Work Center"
      Top             =   2184
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.Label WorkCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Work Center"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label UsedHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   1884
      TabIndex        =   5
      ToolTipText     =   "Hours Required For Setup And Cycle Times"
      Top             =   1920
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label FreeHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   4884
      TabIndex        =   4
      ToolTipText     =   "Calendar Hours Recorded"
      Top             =   1920
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label WCDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   3
      ToolTipText     =   "Work Center"
      Top             =   1464
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.Label FreeHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   4884
      TabIndex        =   2
      ToolTipText     =   "Calendar Hours Recorded"
      Top             =   1200
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label UsedHours 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   1884
      TabIndex        =   1
      ToolTipText     =   "Hours Required For Setup And Cycle Times"
      Top             =   1200
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label WorkCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Work Center"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1572
   End
End
Attribute VB_Name = "CapaCPp10b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'8/22/06 New
'1/2/07 Corrected End Date in query 7.1.9
'2/6/07 Corrected Array error (GetChartInformation) 7.2.5
Option Explicit
Dim bOnLoad As Byte
Dim sWorkCenters(10, 6) As String
Const LabelWidth = 6000

Private Sub cmdCan_Click()
   Form_Deactivate
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4230
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   If bOnLoad = 1 Then
      lblToday = Format(Now, "mm/dd/yy")
      GetChartInformation
      bOnLoad = 0
   End If
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Load()
   'FormLoad Me, ES_DONTLIST
   Move 1000, 1000
   Me.BackColor = vbWhite
   Shop1.BorderStyle = 0
   Shop2.BorderStyle = 0
   Shop3.BorderStyle = 0
   Shop4.BorderStyle = 0
   Shop2.Caption = "0"
   Shop3.Caption = "0"
   Shop4.Caption = "0"
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub



Private Sub GetChartInformation()
   Dim RdoCapa As ADODB.Recordset
   Dim bLabelIdx As Byte
   Dim bNoCal As Byte
   Dim iRow As Integer
   Dim CenterHrs As Long
   Dim CalHrs As Long
   Dim HrsPerc As Currency
   
   Dim sCenter As String
   Dim sShop As String
   
   MouseCursor 13
   If Center <> "ALL" Then sCenter = Compress(Center)
   sShop = Compress(Shop)
   
   Erase sWorkCenters
   On Error GoTo DiaErr1
   sSql = "SELECT WCNREF,WCNNUM,WCNDESC,WCNSHOP FROM WcntTable WHERE " _
          & "(WCNSERVICE=0 AND WCNSHOP='" & sShop & "' AND WCNREF LIKE '" _
          & sCenter & "%') ORDER BY WCNREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCapa, ES_FORWARD)
   If bSqlRows Then
      With RdoCapa
         Do Until .EOF
            WorkCenter(bLabelIdx).Visible = True
            WCDescription(bLabelIdx).Visible = True
            WCDescription(bLabelIdx).BorderStyle = 0
            UsedHours(bLabelIdx).Visible = True
            FreeHours(bLabelIdx).Visible = True
            WorkCenter(bLabelIdx) = "" & Trim(!WCNNUM)
            sWorkCenters(bLabelIdx, 0) = "" & Trim(!WCNREF)
            WCDescription(bLabelIdx) = "" & Trim(!WCNDESC)
            CntRequired(bLabelIdx).Visible = True
            CntAvailable(bLabelIdx).Visible = True
            CntRemaining(bLabelIdx).Visible = True
            .MoveNext
            bLabelIdx = bLabelIdx + 1
            If bLabelIdx > 9 Then Exit Do
         Loop
         ClearResultSet RdoCapa
      End With
   End If
   If bLabelIdx = 0 Then
      MouseCursor 0
      MsgBox "There Is No Information Available to Chart.", _
         vbInformation, Caption
      Exit Sub
   End If
   Shop1.Top = WCDescription(bLabelIdx - 1).Top + 320
   Shop2.Top = WCDescription(bLabelIdx - 1).Top + 320
   Shop3.Top = WCDescription(bLabelIdx - 1).Top + 320
   Shop4.Top = WCDescription(bLabelIdx - 1).Top + 320
   
   For iRow = 0 To bLabelIdx - 1
      sSql = "select SUM(OPSUHRS + (OPUNITHRS*RUNREMAININGQTY)) As CHours FROM RnopTable," _
             & "RunsTable where (OPREF=RUNREF AND OPRUN=RUNNO AND OPCOMPLETE=0 " _
             & "AND OPCENTER='" & sWorkCenters(iRow, 0) & "' AND OPSHOP='" _
             & sShop & "' AND  OPSCHEDDATE <= '" & Format(lblThrough, "mm-dd-yyyy") & "')"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCapa, ES_FORWARD)
      If bSqlRows Then
         With RdoCapa
            If Not IsNull(!cHours) Then
               sWorkCenters(iRow, 1) = str$(Int(!cHours))
            Else
               sWorkCenters(iRow, 1) = "0"
            End If
            CntRequired(iRow).Caption = "Required: " & sWorkCenters(iRow, 1)
         End With
      End If
   Next
   For iRow = 0 To bLabelIdx - 1
      sSql = "SELECT SUM(WCCSHH1+WCCSHH2+WCCSHH3+WCCSHH4) AS CalHours " _
             & "FROM WcclTable WHERE (WCCDATE BETWEEN '" & Format(lblToday, "mm-dd-yyyy") & "' AND " _
             & "'" & Format(lblThrough, "mm-dd-yyyy") & "' AND WCCCENTER='" & sWorkCenters(iRow, 0) _
             & "' AND WCCSHOP='" & sShop & "')"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCapa, ES_FORWARD)
      If bSqlRows Then
         With RdoCapa
            If Not IsNull(!CalHours) Then
               sWorkCenters(iRow, 2) = str$(Int(!CalHours))
            Else
               sWorkCenters(iRow, 2) = "0"
               bNoCal = bNoCal + 1
            End If
            CntAvailable(iRow).Caption = "Available: " & sWorkCenters(iRow, 2)
         End With
      Else
         bNoCal = bNoCal + 1
      End If
   Next
   For iRow = 0 To bLabelIdx - 1
      CenterHrs = Val(sWorkCenters(iRow, 1))
      CalHrs = Val(sWorkCenters(iRow, 2))
      CntRemaining(iRow) = "Remaining: " & str$(CalHrs - CenterHrs)
      Shop2.Caption = str$(Val(Shop2.Caption) + CenterHrs)
      Shop3.Caption = str$(Val(Shop3.Caption) + CalHrs)
      Shop4.Caption = str$(Val(Shop4.Caption) + (CalHrs - CenterHrs))
      If CenterHrs = 0 Then
         UsedHours(iRow).Visible = False
         FreeHours(iRow).Left = UsedHours(iRow).Left
         FreeHours(iRow).Width = 6000
         FreeHours(iRow).Caption = "100.00% Available"
      ElseIf CalHrs = 0 Then
         FreeHours(iRow).Visible = False
         UsedHours(iRow).Width = 6000
         UsedHours(iRow).Caption = "0.00% Available"
      Else
         If CenterHrs >= CalHrs Then
            FreeHours(iRow).Visible = False
            UsedHours(iRow).Width = 6000
         Else
            If CalHrs = 0 Then CalHrs = 1
            HrsPerc = CenterHrs / CalHrs
            UsedHours(iRow).Width = LabelWidth * HrsPerc
            FreeHours(iRow).Left = UsedHours(iRow).Left + UsedHours(iRow).Width
            FreeHours(iRow).Width = LabelWidth - UsedHours(iRow).Width
            FreeHours(iRow).Caption = Format((100 - HrsPerc * 100), "#0.00") & "% Available"
         End If
      End If
      
   Next
   If bNoCal > 0 Then
      MsgBox "There Are (" & Trim(str$(bNoCal)) & " Work Centers Without Calendars.", _
         vbInformation, Caption
   End If
   Set RdoCapa = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "getchartinfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set CapaCPp10b = Nothing
   
End Sub


Private Sub optPrn_Click()
   Dim bResponse As Byte
   bResponse = MsgBox("Print This Form?", ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      cmdHlp.Visible = False
      optPrn.Visible = False
      Sleep 500
      PrintForm
   End If
   optPrn.Visible = True
   cmdHlp.Visible = True
   optPrn.Value = False
   
End Sub
