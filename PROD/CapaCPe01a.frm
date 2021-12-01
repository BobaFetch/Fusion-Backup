VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form CapaCPe01a 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work Centers"
   ClientHeight    =   8100
   ClientLeft      =   2415
   ClientTop       =   1185
   ClientWidth     =   9195
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   4203
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8100
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   Begin VB.Frame tabFrame 
      Height          =   6015
      Index           =   1
      Left            =   5880
      TabIndex        =   45
      Top             =   1920
      Width           =   8655
      Begin VB.TextBox txtRe4 
         Height          =   285
         Index           =   6
         Left            =   2370
         TabIndex        =   146
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   5400
         Width           =   735
      End
      Begin VB.TextBox txtHr4 
         Height          =   285
         Index           =   6
         Left            =   1575
         TabIndex        =   145
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   5400
         Width           =   705
      End
      Begin VB.TextBox txtSt4 
         Height          =   285
         Index           =   6
         Left            =   720
         TabIndex        =   144
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   5400
         Width           =   720
      End
      Begin VB.TextBox txtRe3 
         Height          =   285
         Index           =   6
         Left            =   2370
         TabIndex        =   143
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   5040
         Width           =   735
      End
      Begin VB.TextBox txtHr3 
         Height          =   285
         Index           =   6
         Left            =   1575
         TabIndex        =   142
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   5040
         Width           =   705
      End
      Begin VB.TextBox txtSt3 
         Height          =   285
         Index           =   6
         Left            =   720
         TabIndex        =   141
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   5040
         Width           =   720
      End
      Begin VB.TextBox txtRe2 
         Height          =   285
         Index           =   6
         Left            =   2370
         TabIndex        =   140
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   4680
         Width           =   735
      End
      Begin VB.TextBox txtHr2 
         Height          =   285
         Index           =   6
         Left            =   1575
         TabIndex        =   139
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   4680
         Width           =   705
      End
      Begin VB.TextBox txtSt2 
         Height          =   285
         Index           =   6
         Left            =   720
         TabIndex        =   138
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   4680
         Width           =   720
      End
      Begin VB.TextBox txtRe1 
         Height          =   285
         Index           =   6
         Left            =   2370
         TabIndex        =   137
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox txtHr1 
         Height          =   285
         Index           =   6
         Left            =   1575
         TabIndex        =   136
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   4320
         Width           =   705
      End
      Begin VB.TextBox txtSt1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   720
         TabIndex        =   135
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   4320
         Width           =   720
      End
      Begin VB.TextBox txtRe4 
         Height          =   285
         Index           =   5
         Left            =   7650
         TabIndex        =   134
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtHr4 
         Height          =   285
         Index           =   5
         Left            =   6855
         TabIndex        =   133
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   3720
         Width           =   705
      End
      Begin VB.TextBox txtSt4 
         Height          =   285
         Index           =   5
         Left            =   6000
         TabIndex        =   132
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   3720
         Width           =   720
      End
      Begin VB.TextBox txtRe3 
         Height          =   285
         Index           =   5
         Left            =   7650
         TabIndex        =   131
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtHr3 
         Height          =   285
         Index           =   5
         Left            =   6855
         TabIndex        =   130
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   3360
         Width           =   705
      End
      Begin VB.TextBox txtSt3 
         Height          =   285
         Index           =   5
         Left            =   6000
         TabIndex        =   129
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   3360
         Width           =   720
      End
      Begin VB.TextBox txtRe2 
         Height          =   285
         Index           =   5
         Left            =   7650
         TabIndex        =   128
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtHr2 
         Height          =   285
         Index           =   5
         Left            =   6855
         TabIndex        =   127
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   3000
         Width           =   705
      End
      Begin VB.TextBox txtSt2 
         Height          =   285
         Index           =   5
         Left            =   6000
         TabIndex        =   126
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   3000
         Width           =   720
      End
      Begin VB.TextBox txtRe1 
         Height          =   285
         Index           =   5
         Left            =   7650
         TabIndex        =   125
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtHr1 
         Height          =   285
         Index           =   5
         Left            =   6855
         TabIndex        =   124
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   2640
         Width           =   705
      End
      Begin VB.TextBox txtSt1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   6000
         TabIndex        =   123
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   2640
         Width           =   720
      End
      Begin VB.TextBox txtRe4 
         Height          =   285
         Index           =   4
         Left            =   5010
         TabIndex        =   122
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtHr4 
         Height          =   285
         Index           =   4
         Left            =   4215
         TabIndex        =   121
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   3720
         Width           =   705
      End
      Begin VB.TextBox txtSt4 
         Height          =   285
         Index           =   4
         Left            =   3360
         TabIndex        =   120
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   3720
         Width           =   720
      End
      Begin VB.TextBox txtRe3 
         Height          =   285
         Index           =   4
         Left            =   5010
         TabIndex        =   119
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtHr3 
         Height          =   285
         Index           =   4
         Left            =   4215
         TabIndex        =   118
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   3360
         Width           =   705
      End
      Begin VB.TextBox txtSt3 
         Height          =   285
         Index           =   4
         Left            =   3360
         TabIndex        =   117
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   3360
         Width           =   720
      End
      Begin VB.TextBox txtRe2 
         Height          =   285
         Index           =   4
         Left            =   5010
         TabIndex        =   116
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtHr2 
         Height          =   285
         Index           =   4
         Left            =   4215
         TabIndex        =   115
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   3000
         Width           =   705
      End
      Begin VB.TextBox txtSt2 
         Height          =   285
         Index           =   4
         Left            =   3360
         TabIndex        =   114
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   3000
         Width           =   720
      End
      Begin VB.TextBox txtRe1 
         Height          =   285
         Index           =   4
         Left            =   5010
         TabIndex        =   113
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtHr1 
         Height          =   285
         Index           =   4
         Left            =   4215
         TabIndex        =   112
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   2640
         Width           =   705
      End
      Begin VB.TextBox txtSt1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   3360
         TabIndex        =   111
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   2640
         Width           =   720
      End
      Begin VB.TextBox txtRe4 
         Height          =   285
         Index           =   3
         Left            =   2370
         TabIndex        =   110
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtHr4 
         Height          =   285
         Index           =   3
         Left            =   1575
         TabIndex        =   109
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   3720
         Width           =   705
      End
      Begin VB.TextBox txtSt4 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   108
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   3720
         Width           =   720
      End
      Begin VB.TextBox txtRe3 
         Height          =   285
         Index           =   3
         Left            =   2370
         TabIndex        =   107
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtHr3 
         Height          =   285
         Index           =   3
         Left            =   1575
         TabIndex        =   106
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   3360
         Width           =   705
      End
      Begin VB.TextBox txtSt3 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   105
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   3360
         Width           =   720
      End
      Begin VB.TextBox txtRe2 
         Height          =   285
         Index           =   3
         Left            =   2370
         TabIndex        =   104
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtHr2 
         Height          =   285
         Index           =   3
         Left            =   1575
         TabIndex        =   103
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   3000
         Width           =   705
      End
      Begin VB.TextBox txtSt2 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   102
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   3000
         Width           =   720
      End
      Begin VB.TextBox txtRe1 
         Height          =   285
         Index           =   3
         Left            =   2370
         TabIndex        =   101
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtHr1 
         Height          =   285
         Index           =   3
         Left            =   1575
         TabIndex        =   100
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   2640
         Width           =   705
      End
      Begin VB.TextBox txtSt1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   99
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   2640
         Width           =   720
      End
      Begin VB.TextBox txtRe4 
         Height          =   285
         Index           =   2
         Left            =   7650
         TabIndex        =   98
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtHr4 
         Height          =   285
         Index           =   2
         Left            =   6855
         TabIndex        =   97
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   2040
         Width           =   705
      End
      Begin VB.TextBox txtSt4 
         Height          =   285
         Index           =   2
         Left            =   6000
         TabIndex        =   96
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   2040
         Width           =   720
      End
      Begin VB.TextBox txtRe3 
         Height          =   285
         Index           =   2
         Left            =   7650
         TabIndex        =   95
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtHr3 
         Height          =   285
         Index           =   2
         Left            =   6855
         TabIndex        =   94
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   1680
         Width           =   705
      End
      Begin VB.TextBox txtSt3 
         Height          =   285
         Index           =   2
         Left            =   6000
         TabIndex        =   93
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   1680
         Width           =   720
      End
      Begin VB.TextBox txtRe2 
         Height          =   285
         Index           =   2
         Left            =   7650
         TabIndex        =   92
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtHr2 
         Height          =   285
         Index           =   2
         Left            =   6855
         TabIndex        =   91
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   1320
         Width           =   705
      End
      Begin VB.TextBox txtSt2 
         Height          =   285
         Index           =   2
         Left            =   6000
         TabIndex        =   90
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   1320
         Width           =   720
      End
      Begin VB.TextBox txtRe1 
         Height          =   285
         Index           =   2
         Left            =   7650
         TabIndex        =   89
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtHr1 
         Height          =   285
         Index           =   2
         Left            =   6855
         TabIndex        =   88
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox txtSt1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   6000
         TabIndex        =   87
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   960
         Width           =   720
      End
      Begin VB.TextBox txtRe4 
         Height          =   285
         Index           =   1
         Left            =   5010
         TabIndex        =   86
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtHr4 
         Height          =   285
         Index           =   1
         Left            =   4215
         TabIndex        =   85
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   2040
         Width           =   705
      End
      Begin VB.TextBox txtSt4 
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   84
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   2040
         Width           =   720
      End
      Begin VB.TextBox txtRe3 
         Height          =   285
         Index           =   1
         Left            =   5010
         TabIndex        =   83
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtHr3 
         Height          =   285
         Index           =   1
         Left            =   4215
         TabIndex        =   82
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   1680
         Width           =   705
      End
      Begin VB.TextBox txtSt3 
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   81
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   1680
         Width           =   720
      End
      Begin VB.TextBox txtRe2 
         Height          =   285
         Index           =   1
         Left            =   5010
         TabIndex        =   80
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtHr2 
         Height          =   285
         Index           =   1
         Left            =   4215
         TabIndex        =   79
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   1320
         Width           =   705
      End
      Begin VB.TextBox txtSt2 
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   78
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   1320
         Width           =   720
      End
      Begin VB.TextBox txtRe1 
         Height          =   285
         Index           =   1
         Left            =   5010
         TabIndex        =   77
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtHr1 
         Height          =   285
         Index           =   1
         Left            =   4215
         TabIndex        =   76
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox txtSt1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   75
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   960
         Width           =   720
      End
      Begin VB.CommandButton cmdUpd 
         Caption         =   "&Apply"
         Height          =   375
         Left            =   7440
         TabIndex        =   46
         ToolTipText     =   "Update Work Center Shifts for Mon-Fri"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtSt1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   750
         TabIndex        =   14
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   960
         Width           =   720
      End
      Begin VB.TextBox txtHr1 
         Height          =   285
         Index           =   0
         Left            =   1605
         TabIndex        =   15
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox txtRe1 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   16
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtSt2 
         Height          =   285
         Index           =   0
         Left            =   750
         TabIndex        =   17
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   1320
         Width           =   720
      End
      Begin VB.TextBox txtHr2 
         Height          =   285
         Index           =   0
         Left            =   1605
         TabIndex        =   18
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   1320
         Width           =   705
      End
      Begin VB.TextBox txtRe2 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   19
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtSt3 
         Height          =   285
         Index           =   0
         Left            =   750
         TabIndex        =   20
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   1680
         Width           =   720
      End
      Begin VB.TextBox txtHr3 
         Height          =   285
         Index           =   0
         Left            =   1605
         TabIndex        =   21
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   1680
         Width           =   705
      End
      Begin VB.TextBox txtRe3 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   22
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtSt4 
         Height          =   285
         Index           =   0
         Left            =   750
         TabIndex        =   23
         Tag             =   "5"
         ToolTipText     =   "Time As 8:00a or 12:00p (blank for none)"
         Top             =   2040
         Width           =   720
      End
      Begin VB.TextBox txtHr4 
         Height          =   285
         Index           =   0
         Left            =   1605
         TabIndex        =   24
         Tag             =   "1"
         ToolTipText     =   "Hours Available"
         Top             =   2040
         Width           =   705
      End
      Begin VB.TextBox txtRe4 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   25
         Tag             =   "1"
         ToolTipText     =   "Total Resources For Today"
         Top             =   2040
         Width           =   735
      End
      Begin VB.ComboBox cmbDay 
         Height          =   315
         Left            =   5310
         TabIndex        =   13
         Tag             =   "8"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Update Monday Through Friday With Monday Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   40
         Left            =   2760
         TabIndex        =   147
         Top             =   240
         Width           =   4665
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tuesday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   39
         Left            =   6000
         TabIndex        =   74
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Wednesday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   38
         Left            =   720
         TabIndex        =   73
         Top             =   2400
         Width           =   1065
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Thursday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   37
         Left            =   3360
         TabIndex        =   72
         Top             =   2400
         Width           =   1065
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Friday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   20
         Left            =   6000
         TabIndex        =   71
         Top             =   2400
         Width           =   1065
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Saturday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   720
         TabIndex        =   70
         Top             =   4080
         Width           =   1065
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Monday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   3360
         TabIndex        =   69
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sunday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   36
         Left            =   720
         TabIndex        =   68
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 4:"
         Height          =   285
         Index           =   35
         Left            =   120
         TabIndex        =   67
         Top             =   5400
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 3:"
         Height          =   285
         Index           =   34
         Left            =   120
         TabIndex        =   66
         Top             =   5040
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 2:"
         Height          =   285
         Index           =   33
         Left            =   120
         TabIndex        =   65
         Top             =   4680
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 1:"
         Height          =   285
         Index           =   32
         Left            =   120
         TabIndex        =   64
         Top             =   4320
         Width           =   585
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 4:"
         Height          =   285
         Index           =   31
         Left            =   120
         TabIndex        =   63
         Top             =   3720
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 3:"
         Height          =   285
         Index           =   30
         Left            =   120
         TabIndex        =   62
         Top             =   3360
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 2:"
         Height          =   285
         Index           =   29
         Left            =   120
         TabIndex        =   61
         Top             =   3000
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 1:"
         Height          =   285
         Index           =   28
         Left            =   120
         TabIndex        =   60
         Top             =   2640
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Resources"
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
         Index           =   27
         Left            =   7770
         TabIndex        =   59
         Top             =   4440
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hours       "
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
         Index           =   26
         Left            =   6975
         TabIndex        =   58
         Top             =   4440
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Time"
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
         Index           =   25
         Left            =   6120
         TabIndex        =   57
         Top             =   4440
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Resources"
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
         Index           =   24
         Left            =   5130
         TabIndex        =   56
         Top             =   4440
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hours       "
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
         Index           =   23
         Left            =   4335
         TabIndex        =   55
         Top             =   4440
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Time"
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
         Index           =   22
         Left            =   3480
         TabIndex        =   54
         Top             =   4440
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblCenter 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   52
         Top             =   5160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 1:"
         Height          =   285
         Index           =   14
         Left            =   120
         TabIndex        =   51
         Top             =   960
         Width           =   585
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 2:"
         Height          =   285
         Index           =   15
         Left            =   120
         TabIndex        =   50
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 3:"
         Height          =   285
         Index           =   16
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 4:"
         Height          =   285
         Index           =   17
         Left            =   120
         TabIndex        =   48
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label z1 
         Caption         =   "Week Day"
         Height          =   285
         Index           =   21
         Left            =   3720
         TabIndex        =   47
         Top             =   4920
         Visible         =   0   'False
         Width           =   1185
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   3372
      Index           =   0
      Left            =   0
      TabIndex        =   32
      Top             =   1920
      Width           =   5556
      Begin VB.ComboBox cmbAct 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1716
         Sorted          =   -1  'True
         TabIndex        =   3
         Tag             =   "3"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox optSrv 
         Alignment       =   1  'Right Justify
         Caption         =   "Services Work Center"
         Height          =   285
         Left            =   2820
         TabIndex        =   12
         Top             =   2760
         Width           =   2355
      End
      Begin VB.TextBox txtEte 
         Height          =   285
         Left            =   1716
         TabIndex        =   11
         Tag             =   "1"
         ToolTipText     =   "Default Estimating Rate"
         Top             =   2736
         Width           =   825
      End
      Begin VB.TextBox txtUnt 
         Height          =   285
         Left            =   4380
         TabIndex        =   10
         Tag             =   "1"
         Top             =   2400
         Width           =   825
      End
      Begin VB.TextBox txtSet 
         Height          =   285
         Left            =   1716
         TabIndex        =   9
         Tag             =   "1"
         ToolTipText     =   "Setup Hours"
         Top             =   2376
         Width           =   825
      End
      Begin VB.TextBox txtMdy 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4380
         TabIndex        =   8
         Tag             =   "1"
         ToolTipText     =   "Move Hours"
         Top             =   2040
         Width           =   825
      End
      Begin VB.TextBox txtQdy 
         Height          =   285
         Left            =   1716
         TabIndex        =   7
         Tag             =   "1"
         ToolTipText     =   "Hours In Queue"
         Top             =   2016
         Width           =   825
      End
      Begin VB.TextBox txtSte 
         Height          =   285
         Left            =   1716
         TabIndex        =   6
         Tag             =   "1"
         ToolTipText     =   "Standard Labor Rate"
         Top             =   1656
         Width           =   825
      End
      Begin VB.TextBox txtPoh 
         Height          =   285
         Left            =   1716
         TabIndex        =   5
         Tag             =   "1"
         ToolTipText     =   "Overhead By Percentage Of Labor"
         Top             =   1296
         Width           =   825
      End
      Begin VB.TextBox txtFoh 
         Height          =   285
         Left            =   1716
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "Fixed Overhead"
         Top             =   936
         Width           =   825
      End
      Begin VB.Label lblActdsc 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   "
         Height          =   288
         Left            =   1716
         TabIndex        =   44
         Top             =   612
         Width           =   3132
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Estimating Rate"
         Height          =   288
         Index           =   13
         Left            =   120
         TabIndex        =   43
         Top             =   2736
         Width           =   1272
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Default Unit Hrs"
         Height          =   288
         Index           =   12
         Left            =   2880
         TabIndex        =   42
         Top             =   2400
         Width           =   1392
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Default Setup Hours"
         Height          =   288
         Index           =   11
         Left            =   120
         TabIndex        =   41
         Top             =   2400
         Width           =   1512
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Move Hours"
         Height          =   288
         Index           =   10
         Left            =   2856
         TabIndex        =   40
         Top             =   2016
         Width           =   1392
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Queue Hours"
         Height          =   288
         Index           =   9
         Left            =   120
         TabIndex        =   39
         Top             =   2016
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Rate"
         Height          =   288
         Index           =   8
         Left            =   120
         TabIndex        =   38
         Top             =   1656
         Width           =   1272
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "% Percentage For This Work Center"
         Height          =   288
         Index           =   7
         Left            =   2616
         TabIndex        =   37
         Top             =   1296
         Width           =   2796
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Overhead"
         Height          =   288
         Index           =   6
         Left            =   120
         TabIndex        =   36
         Top             =   1296
         Width           =   1092
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate For This Work Center"
         Height          =   288
         Index           =   5
         Left            =   2616
         TabIndex        =   35
         Top             =   960
         Width           =   2592
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fixed Overhead"
         Height          =   288
         Index           =   4
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1272
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Account"
         Height          =   192
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1332
      End
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   31
      Top             =   1560
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11456
      TabWidthStyle   =   2
      TabFixedWidth   =   1411
      TabFixedHeight  =   473
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Shifts"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   288
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter A New Work Center"
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1050
      Width           =   3075
   End
   Begin VB.ComboBox cmbShp 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   1800
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select From List"
      Top             =   360
      Width           =   1815
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8640
      Top             =   600
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8100
      FormDesignWidth =   9195
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   8280
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin VB.Label Label1 
      Caption         =   "More >>>>"
      Height          =   252
      Left            =   4440
      TabIndex        =   53
      Top             =   1400
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   375
      Index           =   0
      Left            =   300
      TabIndex        =   29
      ToolTipText     =   "Select Or Enter A New Work Center"
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   300
      TabIndex        =   28
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Assign To Shop"
      Height          =   195
      Index           =   2
      Left            =   300
      TabIndex        =   27
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "CapaCPe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'8/8/06 Replaced SSTab
'8/22/06 Fixed boxes where no Work Centers are found
'2/6/07 Fixed Warning in Account Box 7.2.5
Option Explicit

Dim AdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim ADOParameter2 As ADODB.Parameter
Dim RdoWcn As ADODB.Recordset

Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bGoodShop As Byte
Dim bGoodCenter As Boolean
Dim bNewCenter As Byte
Dim bChanged As Byte


Dim iIndex As Integer
Dim sOldCenter As String

'Days/Shift factors
Dim ShiftStart(8, 6) As String
Dim ShiftHours(8, 6) As Currency
Dim ShiftMults(8, 6) As Currency

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Const WKDAY_MON_INDEX As Integer = 1
Private Const WKDAY_SUN_INDEX As Integer = 0
Private Const NUM_WKDAYS As Integer = 7


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   ES_TimeFormat = GetTimeFormat()
   
End Sub

Private Sub cmbAct_Click()
   FindAccount Me
   
End Sub

Private Sub txtHr1_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtHr1_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtHr1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtHr1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtHr1_LostFocus(Index As Integer)
   txtHr1(Index) = CheckLen(txtHr1(Index), 4)
   If Val(txtHr1(Index)) > 12 Then txtHr1(Index) = "12.0"
   txtHr1(Index) = Format(Val(txtHr1(Index)), "#0.0")
   
End Sub


Private Sub txtHr2_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtHr2_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtHr2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtHr2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtHr2_LostFocus(Index As Integer)
   txtHr2(Index) = CheckLen(txtHr2(Index), 4)
   If Val(txtHr2(Index)) > 12 Then txtHr2(Index) = "12.0"
   txtHr2(Index) = Format(Val(txtHr2(Index)), "#0.0")
   
End Sub


Private Sub txtHr3_Change(Index As Integer)
   bChanged = 1
End Sub

Private Sub txtHr3_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtHr3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtHr3_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtHr3_LostFocus(Index As Integer)
   txtHr3(Index) = CheckLen(txtHr3(Index), 4)
   If Val(txtHr3(Index)) > 12 Then txtHr3(Index) = "12.0"
   txtHr3(Index) = Format(Val(txtHr3(Index)), "#0.0")
   
End Sub


Private Sub txtHr4_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtHr4_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtHr4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtHr4_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtHr4_LostFocus(Index As Integer)
   txtHr4(Index) = CheckLen(txtHr4(Index), 4)
   If Val(txtHr4(Index)) > 12 Then txtHr4(Index) = "12.0"
   txtHr4(Index) = Format(Val(txtHr4(Index)), "#0.0")
   
End Sub


Private Sub txtSt1_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtSt1_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtSt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtSt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub

Private Sub txtSt1_LostFocus(Index As Integer)
   txtSt1(Index) = CheckLen(txtSt1(Index), 6)
   Dim tc As New ClassTimeCharge
   txtSt1(Index) = tc.GetTime(txtSt1(Index))
   
End Sub


Private Sub txtSt2_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtSt2_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtSt2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtSt2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub

Private Sub txtSt2_LostFocus(Index As Integer)
   txtSt2(Index) = CheckLen(txtSt2(Index), 6)
   Dim tc As New ClassTimeCharge
   txtSt2(Index) = tc.GetTime(txtSt2(Index))
   
End Sub


Private Sub txtSt3_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtSt3_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtSt3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtSt3_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub

Private Sub txtSt3_LostFocus(Index As Integer)
   txtSt3(Index) = CheckLen(txtSt3(Index), 6)
   Dim tc As New ClassTimeCharge
   txtSt3(Index) = tc.GetTime(txtSt3(Index))
   
End Sub


Private Sub txtSt4_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtSt4_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtSt4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtSt4_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub

Private Sub txtSt4_LostFocus(Index As Integer)
   txtSt4(Index) = CheckLen(txtSt4(Index), 6)
   Dim tc As New ClassTimeCharge
   txtSt4(Index) = tc.GetTime(txtSt4(Index))
   
End Sub

Private Sub txtRe1_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtRe1_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtRe1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtRe1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub

Private Sub txtRe1_LostFocus(Index As Integer)
   Dim res As Double
   txtRe1(Index) = CheckLen(txtRe1(Index), 5)
   res = ConvertToHours(txtRe1(Index))
   txtRe1(Index) = Format(Abs(res), "##0.0")

End Sub


Private Sub txtRe2_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtRe2_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtRe2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtRe2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub

Private Sub txtRe2_LostFocus(Index As Integer)
    Dim res As Double
   txtRe2(Index) = CheckLen(txtRe2(Index), 5)
   res = ConvertToHours(txtRe2(Index))
   txtRe2(Index) = Format(Abs(res), "##0.0")
   
End Sub

Private Function ConvertToHours(strHours As String) As Double
    Dim time() As String
    Dim Value As Double
    
    time = Split(strHours, ":")
        
        
    If (UBound(time) < 1) Then
        ConvertToHours = strHours
    Else
        If (strHours = "") Then
            ConvertToHours = 0
        Else
            If (Trim(time(0)) = "") Then time(0) = 0
            If (Trim(time(1)) = "") Then time(1) = 0
            
            Value = CDbl(time(0)) + CDbl(time(1) / 60)
            ConvertToHours = Value
        End If
    End If
End Function


Private Sub txtRe3_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtRe3_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtRe3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtRe3_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub

Private Sub txtRe3_LostFocus(Index As Integer)
   Dim res As Double
   txtRe3(Index) = CheckLen(txtRe3(Index), 5)
   res = ConvertToHours(txtRe3(Index))
   txtRe3(Index) = Format(Abs(res), "##0.0")
End Sub


Private Sub txtRe4_Change(Index As Integer)
   bChanged = 1
   
End Sub

Private Sub txtRe4_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtRe4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtRe4_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub

Private Sub txtRe4_LostFocus(Index As Integer)
   Dim res As Double
   txtRe4(Index) = CheckLen(txtRe4(Index), 5)
   res = ConvertToHours(txtRe4(Index))
   txtRe4(Index) = Format(Abs(res), "##0.0")
   
End Sub



Private Sub cmbAct_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   Dim sAccount As String
   
   On Error Resume Next
   If Len(Trim(cmbAct)) > 0 Then
      For iList = 0 To cmbAct.ListCount - 1
         If cmbAct = cmbAct.List(iList) Then b = True
      Next
      If b = 0 Then
         Beep
         cmbAct = "" & Trim(RdoWcn!WCNACCT)
      End If
      FindAccount Me
   End If
   
   RdoWcn!WCNACCT = "" & Compress(cmbAct)
   RdoWcn.Update
   If Err > 0 Then ValidateEdit
   MouseCursor 0
   
End Sub


Private Sub cmbShp_Click()
       
   If bChanged Then UpdateWc
   bChanged = False
   tab1.Enabled = False
   FillWorkCenters
   
End Sub

Private Sub cmbWcn_Click()
   If bChanged Then UpdateWc
   bChanged = False
   bGoodCenter = GetCenter(True)
   
End Sub

Private Sub cmbWcn_LostFocus()
   cmbWcn = CheckLen(cmbWcn, 12)
   If bCancel = 1 Then Exit Sub
   If Len(cmbWcn) = 0 Then
      bGoodCenter = False
      cmdCan.SetFocus
      Exit Sub
   Else
      bGoodCenter = GetCenter(True)
   End If
   If Not bGoodCenter Then
      AddCenter
   Else
      bNewCenter = False
      On Error Resume Next
      txtDsc.SetFocus
   End If
   MouseCursor 0
End Sub

Private Sub cmdCan_Click()

   If bChanged Then UpdateWc
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   bCancel = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4203
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim iList As Integer
   Dim sMsg As String
   
   sMsg = "Update Monday Through Friday With " & vbCr _
          & "With The Current Settings?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      'Start Time
      If bGoodCenter Then
         MouseCursor 13
         
         On Error Resume Next
         RdoWcn!WCNMONSH1 = "" & txtSt1(WKDAY_MON_INDEX)
         RdoWcn!WCNTUESH1 = "" & txtSt1(WKDAY_MON_INDEX)
         RdoWcn!WCNWEDSH1 = "" & txtSt1(WKDAY_MON_INDEX)
         RdoWcn!WCNTHUSH1 = "" & txtSt1(WKDAY_MON_INDEX)
         RdoWcn!WCNFRISH1 = "" & txtSt1(WKDAY_MON_INDEX)
         
         RdoWcn!WCNMONSH2 = "" & txtSt2(WKDAY_MON_INDEX)
         RdoWcn!WCNTUESH2 = "" & txtSt2(WKDAY_MON_INDEX)
         RdoWcn!WCNWEDSH2 = "" & txtSt2(WKDAY_MON_INDEX)
         RdoWcn!WCNTHUSH2 = "" & txtSt2(WKDAY_MON_INDEX)
         RdoWcn!WCNFRISH2 = "" & txtSt2(WKDAY_MON_INDEX)
         
         RdoWcn!WCNMONSH3 = "" & txtSt3(WKDAY_MON_INDEX)
         RdoWcn!WCNTUESH3 = "" & txtSt3(WKDAY_MON_INDEX)
         RdoWcn!WCNWEDSH3 = "" & txtSt3(WKDAY_MON_INDEX)
         RdoWcn!WCNTHUSH3 = "" & txtSt3(WKDAY_MON_INDEX)
         RdoWcn!WCNFRISH3 = "" & txtSt3(WKDAY_MON_INDEX)
         
         RdoWcn!WCNMONSH4 = "" & txtSt4(WKDAY_MON_INDEX)
         RdoWcn!WCNTUESH4 = "" & txtSt4(WKDAY_MON_INDEX)
         RdoWcn!WCNWEDSH4 = "" & txtSt4(WKDAY_MON_INDEX)
         RdoWcn!WCNTHUSH4 = "" & txtSt4(WKDAY_MON_INDEX)
         RdoWcn!WCNFRISH4 = "" & txtSt4(WKDAY_MON_INDEX)
         
         'Hours
         RdoWcn!WCNMONHR1 = 0 + Val(txtHr1(WKDAY_MON_INDEX))
         RdoWcn!WCNTUEHR1 = 0 + Val(txtHr1(WKDAY_MON_INDEX))
         RdoWcn!WCNWEDHR1 = 0 + Val(txtHr1(WKDAY_MON_INDEX))
         RdoWcn!WCNTHUHR1 = 0 + Val(txtHr1(WKDAY_MON_INDEX))
         RdoWcn!WCNFRIHR1 = 0 + Val(txtHr1(WKDAY_MON_INDEX))
         
         RdoWcn!WCNMONHR2 = 0 + Val(txtHr2(WKDAY_MON_INDEX))
         RdoWcn!WCNTUEHR2 = 0 + Val(txtHr2(WKDAY_MON_INDEX))
         RdoWcn!WCNWEDHR2 = 0 + Val(txtHr2(WKDAY_MON_INDEX))
         RdoWcn!WCNTHUHR2 = 0 + Val(txtHr2(WKDAY_MON_INDEX))
         RdoWcn!WCNFRIHR2 = 0 + Val(txtHr2(WKDAY_MON_INDEX))
         
         RdoWcn!WCNMONHR3 = 0 + Val(txtHr3(WKDAY_MON_INDEX))
         RdoWcn!WCNTUEHR3 = 0 + Val(txtHr3(WKDAY_MON_INDEX))
         RdoWcn!WCNWEDHR3 = 0 + Val(txtHr3(WKDAY_MON_INDEX))
         RdoWcn!WCNTHUHR3 = 0 + Val(txtHr3(WKDAY_MON_INDEX))
         RdoWcn!WCNFRIHR3 = 0 + Val(txtHr3(WKDAY_MON_INDEX))
         
         RdoWcn!WCNMONHR4 = 0 + Val(txtHr4(WKDAY_MON_INDEX))
         RdoWcn!WCNTUEHR4 = 0 + Val(txtHr4(WKDAY_MON_INDEX))
         RdoWcn!WCNWEDHR4 = 0 + Val(txtHr4(WKDAY_MON_INDEX))
         RdoWcn!WCNTHUHR4 = 0 + Val(txtHr4(WKDAY_MON_INDEX))
         RdoWcn!WCNFRIHR4 = 0 + Val(txtHr4(WKDAY_MON_INDEX))
         
         'Resources
         RdoWcn!WCNMONMU1 = 0 + Val(txtRe1(WKDAY_MON_INDEX))
         RdoWcn!WCNTUEMU1 = 0 + Val(txtRe1(WKDAY_MON_INDEX))
         RdoWcn!WCNWEDMU1 = 0 + Val(txtRe1(WKDAY_MON_INDEX))
         RdoWcn!WCNTHUMU1 = 0 + Val(txtRe1(WKDAY_MON_INDEX))
         RdoWcn!WCNFRIMU1 = 0 + Val(txtRe1(WKDAY_MON_INDEX))
         
         RdoWcn!WCNMONMU2 = 0 + Val(txtRe2(WKDAY_MON_INDEX))
         RdoWcn!WCNTUEMU2 = 0 + Val(txtRe2(WKDAY_MON_INDEX))
         RdoWcn!WCNWEDMU2 = 0 + Val(txtRe2(WKDAY_MON_INDEX))
         RdoWcn!WCNTHUMU2 = 0 + Val(txtRe2(WKDAY_MON_INDEX))
         RdoWcn!WCNFRIMU2 = 0 + Val(txtRe2(WKDAY_MON_INDEX))
         
         RdoWcn!WCNMONMU3 = 0 + Val(txtRe3(WKDAY_MON_INDEX))
         RdoWcn!WCNTUEMU3 = 0 + Val(txtRe3(WKDAY_MON_INDEX))
         RdoWcn!WCNWEDMU3 = 0 + Val(txtRe3(WKDAY_MON_INDEX))
         RdoWcn!WCNTHUMU3 = 0 + Val(txtRe3(WKDAY_MON_INDEX))
         RdoWcn!WCNFRIMU3 = 0 + Val(txtRe3(WKDAY_MON_INDEX))
         
         RdoWcn!WCNMONMU4 = 0 + Val(txtRe4(WKDAY_MON_INDEX))
         RdoWcn!WCNTUEMU4 = 0 + Val(txtRe4(WKDAY_MON_INDEX))
         RdoWcn!WCNWEDMU4 = 0 + Val(txtRe4(WKDAY_MON_INDEX))
         RdoWcn!WCNTHUMU4 = 0 + Val(txtRe4(WKDAY_MON_INDEX))
         RdoWcn!WCNFRIMU4 = 0 + Val(txtRe4(WKDAY_MON_INDEX))
         RdoWcn.Update
         If Err > 0 Then ValidateEdit
      End If
      
      For iList = WKDAY_MON_INDEX To (NUM_WKDAYS - 2) ' Friday
         ShiftStart(iList, 1) = "" & txtSt1(WKDAY_MON_INDEX)
         ShiftStart(iList, 2) = "" & txtSt2(WKDAY_MON_INDEX)
         ShiftStart(iList, 3) = "" & txtSt3(WKDAY_MON_INDEX)
         ShiftStart(iList, 4) = "" & txtSt4(WKDAY_MON_INDEX)
         
        txtSt1(iList) = "" & Trim(ShiftStart(iList, 1))
        txtSt2(iList) = "" & Trim(ShiftStart(iList, 2))
        txtSt3(iList) = "" & Trim(ShiftStart(iList, 3))
        txtSt4(iList) = "" & Trim(ShiftStart(iList, 4))
         
         ShiftHours(iList, 1) = 0 + Val(txtHr1(WKDAY_MON_INDEX))
         ShiftHours(iList, 2) = 0 + Val(txtHr2(WKDAY_MON_INDEX))
         ShiftHours(iList, 3) = 0 + Val(txtHr3(WKDAY_MON_INDEX))
         ShiftHours(iList, 4) = 0 + Val(txtHr4(WKDAY_MON_INDEX))
            
        txtHr1(iList) = Format(ShiftHours(iList, 1), "#0.0")
        txtHr2(iList) = Format(ShiftHours(iList, 2), "#0.0")
        txtHr3(iList) = Format(ShiftHours(iList, 3), "#0.0")
        txtHr4(iList) = Format(ShiftHours(iList, 4), "#0.0")
            
         
         ShiftMults(iList, 1) = 0 + Val(txtRe1(WKDAY_MON_INDEX))
         ShiftMults(iList, 2) = 0 + Val(txtRe2(WKDAY_MON_INDEX))
         ShiftMults(iList, 3) = 0 + Val(txtRe3(WKDAY_MON_INDEX))
         ShiftMults(iList, 4) = 0 + Val(txtRe4(WKDAY_MON_INDEX))
      
            txtRe1(iList) = Format(ShiftMults(iList, 1), "#0.0")
            txtRe2(iList) = Format(ShiftMults(iList, 2), "#0.0")
            txtRe3(iList) = Format(ShiftMults(iList, 3), "#0.0")
            txtRe4(iList) = Format(ShiftMults(iList, 4), "#0.0")
      Next
      MouseCursor 0
      SysMsg "Shifts Updated.", True, Me
   Else
      CancelTrans
   End If
   
   On Error Resume Next
End Sub




Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      ManageBoxes False
      FillAccounts
      FillShops
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   On Error Resume Next
   'Tab1.Tab = 0
   'Me.Width = 5760
   tabFrame(0).BorderStyle = 0
   tabFrame(1).BorderStyle = 0
   tabFrame(0).Visible = True
   tabFrame(1).Visible = False
   tabFrame(0).Left = 10
   tabFrame(1).Left = 10
   
   sSql = "SELECT * FROM WcntTable WHERE WCNREF= ? AND WCNSHOP= ? "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 12
   
   Set ADOParameter2 = New ADODB.Parameter
   ADOParameter2.Type = adChar
   ADOParameter2.SIZE = 12
   
   AdoQry.Parameters.Append AdoParameter1
   AdoQry.Parameters.Append ADOParameter2
   
   bNewCenter = False
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter1 = Nothing
   Set ADOParameter2 = Nothing
   Set AdoQry = Nothing
   Set RdoWcn = Nothing
   Set CapaCPe01a = Nothing
   
End Sub



Private Function GetShop() As Byte
   Dim sShop As String
   Dim RdoWcn2 As ADODB.Recordset
   sShop = Compress(cmbShp)
   GetShop = False
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM ShopTable WHERE SHPREF='" & sShop & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoWcn2, ES_FORWARD)
   If bSqlRows Then
      With RdoWcn2
         GetShop = True
         cmbShp = "" & Trim(!SHPNUM)
         If bNewCenter Then
            txtEte = Format(0 + !SHPESTRATE, ES_QuantityDataFormat)
            cmbAct = "" & Trim(!SHPACCT)
            FindAccount Me
            txtQdy = Format(!SHPQHRS, "##0.000")
            txtMdy = Format(!SHPMHRS, "##0.000")
            txtSet = Format(!SHPSUHRS, "##0.000")
            txtUnt = Format(!SHPUNITHRS, ES_TimeFormat)
            txtFoh = Format(!SHPOHTOTAL, ES_QuantityDataFormat)
            txtPoh = Format(!SHPOHRATE, ES_QuantityDataFormat)
            txtSte = Format(!SHPOHRATE, ES_QuantityDataFormat)
            optSrv.Value = !SHPSERVICE
            bNewCenter = False
            RdoWcn!WCNACCT = "" & cmbAct
            RdoWcn!WCNESTRATE = Val(txtEte)
            RdoWcn!WCNOHFIXED = Val(txtFoh)
            RdoWcn!WCNMHRS = Val(txtMdy)
            RdoWcn!WCNOHPCT = Val(txtPoh)
            RdoWcn!WCNQHRS = Val(txtQdy)
            RdoWcn!WCNSUHRS = Val(txtSet)
            RdoWcn!WCNUNITHRS = Val(txtUnt)
            RdoWcn!WCNSERVICE = optSrv.Value
            RdoWcn.Update
         End If
         ClearResultSet RdoWcn2
      End With
   Else
      GetShop = False
   End If
   On Error Resume Next
   Set RdoWcn2 = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getshop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub FillWorkCenters()
   cmbWcn.Clear
   txtDsc = ""
   bGoodCenter = False
   On Error GoTo DiaErr1
   sSql = "Qry_FillWorkCenters '" & Compress(cmbShp) & "'"
   LoadComboBox cmbWcn
   If cmbWcn.ListCount > 0 Then
      cmbWcn = cmbWcn.List(0)
      ' bGoodCenter = GetCenter(True)
   Else
      cmbAct = ""
      txtFoh = "0.000"
      txtPoh = "0.000"
      txtSte = "0.00"
      txtQdy = "0.000"
      txtSet = "0.000"
      txtEte = "0.000"
      txtMdy = "0.000"
      txtUnt = Format(0, ES_TimeFormat)
      optSrv.Value = vbUnchecked
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillworkcenters"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillShops()
   On Error GoTo DiaErr1
   sSql = "Qry_FillShops "
   LoadComboBox cmbShp
   tab1.Enabled = False
   If bSqlRows Then cmbShp = cmbShp.List(0)
   FillWorkCenters
   Exit Sub
   
DiaErr1:
   sProcName = "fillshops"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblActdsc_Change()
   If Trim(lblActdsc) = "*** Account Wasn't Found ***" Then
      lblActdsc.ForeColor = ES_RED
   Else
      lblActdsc.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub optSrv_Click()
   If bGoodCenter Then
      On Error Resume Next
      RdoWcn!WCNSERVICE = optSrv.Value
      RdoWcn.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub tab1_Click()
   On Error Resume Next
   If tab1.SelectedItem.Index = 1 Then
      tabFrame(0).Visible = True
      tabFrame(1).Visible = False
      cmbAct.SetFocus
   Else
      tabFrame(1).Visible = True
      tabFrame(0).Visible = False
   End If
   
End Sub

Private Sub Tab1_GotFocus()
   On Error Resume Next
End Sub



Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   If bGoodCenter Then
      On Error Resume Next
      RdoWcn!WCNDESC = "" & txtDsc
      RdoWcn.Update
      If Err > 0 Then ValidateEdit
   End If
   lblCenter = cmbWcn & " - " & txtDsc
   
End Sub

Private Sub txtEte_LostFocus()
   txtEte = CheckLen(txtEte, 7)
   txtEte = Format(Abs(Val(txtEte)), ES_QuantityDataFormat)
   If bGoodCenter Then
      On Error Resume Next
      RdoWcn!WCNESTRATE = Val(txtEte)
      RdoWcn.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtFoh_LostFocus()
   txtFoh = CheckLen(txtFoh, 7)
   txtFoh = Format(Abs(Val(txtFoh)), ES_QuantityDataFormat)
   If bGoodCenter Then
      On Error Resume Next
      RdoWcn!WCNOHFIXED = Val(txtFoh)
      RdoWcn!WCNACCT = "" & Compress(cmbAct)
      RdoWcn.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtMdy_LostFocus()
   txtMdy = CheckLen(txtMdy, 7)
   txtMdy = Format(Abs(Val(txtMdy)), ES_QuantityDataFormat)
   If bGoodCenter Then
      On Error Resume Next
      RdoWcn!WCNMHRS = Val(txtMdy)
      RdoWcn.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtPoh_LostFocus()
   txtPoh = CheckLen(txtPoh, 7)
   txtPoh = Format(Abs(Val(txtPoh)), ES_QuantityDataFormat)
   If bGoodCenter Then
      On Error Resume Next
      RdoWcn!WCNOHPCT = Val(txtPoh)
      RdoWcn.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtQdy_LostFocus()
   txtQdy = CheckLen(txtQdy, 7)
   txtQdy = Format(Abs(Val(txtQdy)), ES_QuantityDataFormat)
   If bGoodCenter Then
      On Error Resume Next
      RdoWcn!WCNQHRS = Val(txtQdy)
      RdoWcn.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub txtSet_LostFocus()
   txtSet = CheckLen(txtSet, 7)
   txtSet = Format(Abs(Val(txtSet)), ES_QuantityDataFormat)
   If bGoodCenter Then
      On Error Resume Next
      RdoWcn!WCNSUHRS = Val(txtSet)
      RdoWcn.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtSte_LostFocus()
   txtSte = CheckLen(txtSte, 7)
   txtSte = Format(Abs(Val(txtSte)), ES_QuantityDataFormat)
   If bGoodCenter Then
      On Error Resume Next
      RdoWcn!WCNSTDRATE = Val(txtSte)
      RdoWcn.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtUnt_LostFocus()
   txtUnt = CheckLen(txtUnt, 8)
   txtUnt = Format(Abs(Val(txtUnt)), ES_TimeFormat)
   If bGoodCenter Then
      On Error Resume Next
      RdoWcn!WCNUNITHRS = Val(txtUnt)
      RdoWcn.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Function GetCenter(bOpen As Byte) As Boolean
   Erase ShiftStart
   Erase ShiftHours
   Erase ShiftMults
   
   Dim iList As Integer
   
   GetCenter = False
   On Error GoTo DiaErr1
   AdoParameter1.Value = Compress(cmbWcn)
   ADOParameter2.Value = Compress(cmbShp)
   
   bSqlRows = clsADOCon.GetQuerySet(RdoWcn, AdoQry, ES_KEYSET, True, 1)
   If bSqlRows Then
      With RdoWcn
         GetCenter = True
         cmbWcn = "" & Trim(!WCNNUM)
         txtDsc = "" & Trim(!WCNDESC)
         cmbAct = "" & Trim(!WCNACCT)
         FindAccount Me
         If cmbAct = "" Then lblActdsc = ""
         txtFoh = Format(!WCNOHFIXED, ES_QuantityDataFormat)
         txtPoh = Format(!WCNOHPCT, "##0.000")
         txtSte = Format(!WCNSTDRATE, "##0.000")
         txtQdy = Format(!WCNQHRS, "##0.000")
         txtMdy = Format(!WCNMHRS, "##0.000")
         txtSet = Format(!WCNSUHRS, "##0.000")
         txtUnt = Format(!WCNUNITHRS, ES_TimeFormat)
         txtEte = Format(!WCNESTRATE, "##0.000")
         optSrv.Value = !WCNSERVICE
         lblCenter = cmbWcn & " - " & txtDsc
         ShiftStart(1, 1) = "" & Trim(!WCNSUNSH1)
         ShiftStart(1, 2) = "" & Trim(!WCNSUNSH2)
         ShiftStart(1, 3) = "" & Trim(!WCNSUNSH3)
         ShiftStart(1, 4) = "" & Trim(!WCNSUNSH4)
         
         ShiftStart(2, 1) = "" & Trim(!WCNMONSH1)
         ShiftStart(2, 2) = "" & Trim(!WCNMONSH2)
         ShiftStart(2, 3) = "" & Trim(!WCNMONSH3)
         ShiftStart(2, 4) = "" & Trim(!WCNMONSH4)
         
         ShiftStart(3, 1) = "" & Trim(!WCNTUESH1)
         ShiftStart(3, 2) = "" & Trim(!WCNTUESH2)
         ShiftStart(3, 3) = "" & Trim(!WCNTUESH3)
         ShiftStart(3, 4) = "" & Trim(!WCNTUESH4)
         
         ShiftStart(4, 1) = "" & Trim(!WCNWEDSH1)
         ShiftStart(4, 2) = "" & Trim(!WCNWEDSH2)
         ShiftStart(4, 3) = "" & Trim(!WCNWEDSH3)
         ShiftStart(4, 4) = "" & Trim(!WCNWEDSH4)
         
         ShiftStart(5, 1) = "" & Trim(!WCNTHUSH1)
         ShiftStart(5, 2) = "" & Trim(!WCNTHUSH2)
         ShiftStart(5, 3) = "" & Trim(!WCNTHUSH3)
         ShiftStart(5, 4) = "" & Trim(!WCNTHUSH4)
         
         ShiftStart(6, 1) = "" & Trim(!WCNFRISH1)
         ShiftStart(6, 2) = "" & Trim(!WCNFRISH2)
         ShiftStart(6, 3) = "" & Trim(!WCNFRISH3)
         ShiftStart(6, 4) = "" & Trim(!WCNFRISH4)
         
         ShiftStart(7, 1) = "" & Trim(!WCNSATSH1)
         ShiftStart(7, 2) = "" & Trim(!WCNSATSH2)
         ShiftStart(7, 3) = "" & Trim(!WCNSATSH3)
         ShiftStart(7, 4) = "" & Trim(!WCNSATSH4)
         
         ShiftHours(1, 1) = Format(!WCNSUNHR1, "#0.0")
         ShiftHours(1, 2) = Format(!WCNSUNHR2, "#0.0")
         ShiftHours(1, 3) = Format(!WCNSUNHR3, "#0.0")
         ShiftHours(1, 4) = Format(!WCNSUNHR4, "#0.0")
         
         ShiftHours(2, 1) = Format(!WCNMONHR1, "#0.0")
         ShiftHours(2, 2) = Format(!WCNMONHR2, "#0.0")
         ShiftHours(2, 3) = Format(!WCNMONHR3, "#0.0")
         ShiftHours(2, 4) = Format(!WCNMONHR4, "#0.0")
         
         ShiftHours(3, 1) = Format(!WCNTUEHR1, "#0.0")
         ShiftHours(3, 2) = Format(!WCNTUEHR2, "#0.0")
         ShiftHours(3, 3) = Format(!WCNTUEHR3, "#0.0")
         ShiftHours(3, 4) = Format(!WCNTUEHR4, "#0.0")
         
         ShiftHours(4, 1) = Format(!WCNWEDHR1, "#0.0")
         ShiftHours(4, 2) = Format(!WCNWEDHR2, "#0.0")
         ShiftHours(4, 3) = Format(!WCNWEDHR3, "#0.0")
         ShiftHours(4, 4) = Format(!WCNWEDHR4, "#0.0")
         
         ShiftHours(5, 1) = Format(!WCNTHUHR1, "#0.0")
         ShiftHours(5, 2) = Format(!WCNTHUHR2, "#0.0")
         ShiftHours(5, 3) = Format(!WCNTHUHR3, "#0.0")
         ShiftHours(5, 4) = Format(!WCNTHUHR4, "#0.0")
         
         ShiftHours(6, 1) = Format(!WCNFRIHR1, "#0.0")
         ShiftHours(6, 2) = Format(!WCNFRIHR2, "#0.0")
         ShiftHours(6, 3) = Format(!WCNFRIHR3, "#0.0")
         ShiftHours(6, 4) = Format(!WCNFRIHR4, "#0.0")
         
         ShiftHours(7, 1) = Format(!WCNSATHR1, "#0.0")
         ShiftHours(7, 2) = Format(!WCNSATHR2, "#0.0")
         ShiftHours(7, 3) = Format(!WCNSATHR3, "#0.0")
         ShiftHours(7, 4) = Format(!WCNSATHR4, "#0.0")
         
         ShiftMults(1, 1) = Format(!WCNSUNMU1, "#0.0")
         ShiftMults(1, 2) = Format(!WCNSUNMU2, "#0.0")
         ShiftMults(1, 3) = Format(!WCNSUNMU3, "#0.0")
         ShiftMults(1, 4) = Format(!WCNSUNMU4, "#0.0")
         
         ShiftMults(2, 1) = Format(!WCNMONMU1, "#0.0")
         ShiftMults(2, 2) = Format(!WCNMONMU2, "#0.0")
         ShiftMults(2, 3) = Format(!WCNMONMU3, "#0.0")
         ShiftMults(2, 4) = Format(!WCNMONMU4, "#0.0")
         
         ShiftMults(3, 1) = Format(!WCNTUEMU1, "#0.0")
         ShiftMults(3, 2) = Format(!WCNTUEMU2, "#0.0")
         ShiftMults(3, 3) = Format(!WCNTUEMU3, "#0.0")
         ShiftMults(3, 4) = Format(!WCNTUEMU4, "#0.0")
         
         ShiftMults(4, 1) = Format(!WCNWEDMU1, "#0.0")
         ShiftMults(4, 2) = Format(!WCNWEDMU2, "#0.0")
         ShiftMults(4, 3) = Format(!WCNWEDMU3, "#0.0")
         ShiftMults(4, 4) = Format(!WCNWEDMU4, "#0.0")
         
         ShiftMults(5, 1) = Format(!WCNTHUMU1, "#0.0")
         ShiftMults(5, 2) = Format(!WCNTHUMU2, "#0.0")
         ShiftMults(5, 3) = Format(!WCNTHUMU3, "#0.0")
         ShiftMults(5, 4) = Format(!WCNTHUMU4, "#0.0")
         
         ShiftMults(6, 1) = Format(!WCNFRIMU1, "#0.0")
         ShiftMults(6, 2) = Format(!WCNFRIMU2, "#0.0")
         ShiftMults(6, 3) = Format(!WCNFRIMU3, "#0.0")
         ShiftMults(6, 4) = Format(!WCNFRIMU4, "#0.0")
         
         ShiftMults(7, 1) = Format(!WCNSATMU1, "#0.0")
         ShiftMults(7, 2) = Format(!WCNSATMU2, "#0.0")
         ShiftMults(7, 3) = Format(!WCNSATMU3, "#0.0")
         ShiftMults(7, 4) = Format(!WCNSATMU4, "#0.0")

        For iList = 0 To (NUM_WKDAYS - 1)
            txtSt1(iList) = "" & Trim(ShiftStart((iList + 1), 1))
            txtSt2(iList) = "" & Trim(ShiftStart((iList + 1), 2))
            txtSt3(iList) = "" & Trim(ShiftStart((iList + 1), 3))
            txtSt4(iList) = "" & Trim(ShiftStart((iList + 1), 4))
            
            txtSt1(iList).ToolTipText = "Shift Start-Enter As 8.00a"
            txtSt2(iList).ToolTipText = "Shift Start-Enter As 8.00a"
            txtSt3(iList).ToolTipText = "Shift Start-Enter As 8.00a"
            txtSt4(iList).ToolTipText = "Shift Start-Enter As 8.00a"
            
            txtHr1(iList) = Format(ShiftHours((iList + 1), 1), "#0.0")
            txtHr2(iList) = Format(ShiftHours((iList + 1), 2), "#0.0")
            txtHr3(iList) = Format(ShiftHours((iList + 1), 3), "#0.0")
            txtHr4(iList) = Format(ShiftHours((iList + 1), 4), "#0.0")
            
            txtHr1(iList).ToolTipText = "Shift Hours Enter As 2.5"
            txtHr2(iList).ToolTipText = "Shift Hours Enter As 2.5"
            txtHr3(iList).ToolTipText = "Shift Hours Enter As 2.5"
            txtHr4(iList).ToolTipText = "Shift Hours Enter As 2.5"
        
            txtRe1(iList) = Format(ShiftMults((iList + 1), 1), "#0.0")
            txtRe2(iList) = Format(ShiftMults((iList + 1), 2), "#0.0")
            txtRe3(iList) = Format(ShiftMults((iList + 1), 3), "#0.0")
            txtRe4(iList) = Format(ShiftMults((iList + 1), 4), "#0.0")
        
            txtRe1(iList).ToolTipText = "Resource Hours Enter As 1.0"
            txtRe2(iList).ToolTipText = "Resource Hours Enter As 1.0"
            txtRe3(iList).ToolTipText = "Resource Hours Enter As 1.0"
            txtRe4(iList).ToolTipText = "Resource Hours Enter As 1.0"
        Next
        
'         txtSt1 = ShiftStart(iIndex, 1)
'         txtSt2 = ShiftStart(iIndex, 2)
'         txtSt3 = ShiftStart(iIndex, 3)
'         txtSt4 = ShiftStart(iIndex, 4)
'
'         txtHr1 = Format(ShiftHours(iIndex, 1), "#0.0")
'         txtHr2 = Format(ShiftHours(iIndex, 2), "#0.0")
'         txtHr3 = Format(ShiftHours(iIndex, 3), "#0.0")
'         txtHr4 = Format(ShiftHours(iIndex, 4), "#0.0")
'
'         txtRe1 = Format(ShiftMults(iIndex, 1), "#0.0")
'         txtRe2 = Format(ShiftMults(iIndex, 2), "#0.0")
'         txtRe3 = Format(ShiftMults(iIndex, 3), "#0.0")
'         txtRe4 = Format(ShiftMults(iIndex, 4), "#0.0")
         If cmbAct.ListCount > 0 Then cmbAct.Enabled = True
      End With
      ManageBoxes True
   Else
      cmbAct.Enabled = False
      GetCenter = False
      cmbAct = ""
      txtDsc = ""
      txtFoh = ""
      txtPoh = ""
      txtSte = ""
      txtQdy = ""
      txtMdy = ""
      txtSet = ""
      txtUnt = ""
      txtEte = ""
      optSrv.Value = vbUnchecked
'      txtSt1 = ""
'      txtSt2 = ""
'      txtSt3 = ""
'      txtSt4 = ""
'      txtHr1 = "0"
'      txtHr2 = "0"
'      txtHr3 = "0"
'      txtHr4 = "0"
'      txtRe1 = ""
'      txtRe2 = ""
'      txtRe3 = ""
'      txtRe4 = ""
      ManageBoxes False
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getcenter"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddCenter()
   Dim sShop As String
   Dim sNewCenter As String
   Dim Response As Integer
   
   Response = MsgBox(cmbWcn & " Wasn't Found. Add It?", ES_YESQUESTION, Caption)
   If Response = vbNo Then
      bGoodCenter = False
      On Error Resume Next
      If cmbWcn.ListCount > 0 Then cmbWcn = cmbWcn.List(0)
      cmbWcn.SetFocus
      Width = Width + 10
      Exit Sub
   End If
   Response = IllegalCharacters(cmbWcn)
   If Response > 0 Then
      MsgBox "The Work Center Contains An Illegal " & Chr$(Response) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   
   MouseCursor 11
   sShop = Compress(cmbShp)
   sNewCenter = Compress(cmbWcn)
   
   On Error Resume Next
   RdoWcn.Close
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "Select * FROM WcntTable"
   Set RdoWcn = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   RdoWcn.AddNew
   RdoWcn!WCNREF = sNewCenter
   RdoWcn!WCNNUM = cmbWcn
   RdoWcn!WCNSHOP = "" & sShop
   RdoWcn!WCNMONHR1 = 8
   RdoWcn!WCNTUEHR1 = 8
   RdoWcn!WCNWEDHR1 = 8
   RdoWcn!WCNTHUHR1 = 8
   RdoWcn!WCNFRIHR1 = 8
   RdoWcn!WCNMONMU1 = 1
   RdoWcn!WCNTUEMU1 = 1
   RdoWcn!WCNWEDMU1 = 1
   RdoWcn!WCNTHUMU1 = 1
   RdoWcn!WCNFRIMU1 = 1
   RdoWcn.Update
   MouseCursor 0
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      AddComboStr cmbWcn.hwnd, cmbWcn
      Set RdoWcn = Nothing
      bGoodCenter = GetCenter(True)
      bNewCenter = True
      tab1.Enabled = True
      'Tab1.Tab = 0
      txtDsc.SetFocus
      SysMsg cmbWcn & " Added.", True, Me
   Else
      clsADOCon.RollbackTrans
      MsgBox "Could Not Successfully Add The Work Center", _
         vbExclamation, Caption
   End If
   
   Exit Sub
   
DiaErr1:
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   MouseCursor 0
   On Error Resume Next
   clsADOCon.RollbackTrans
   MsgBox CurrError.Description & vbCr & "Couldn't Add Work Center.", vbExclamation, Caption
   cmbWcn.SetFocus
   
End Sub


Private Sub FillAccounts()
   On Error GoTo DiaErr1
   sSql = "Qry_FillLowAccounts"
   LoadComboBox cmbAct
   If cmbAct.ListCount > 0 Then
      cmbAct = cmbAct.List(0)
      'cmbAct.Enabled = True
      FindAccount Me
   Else
      cmbAct.Enabled = False
      cmbAct = "No Accounts."
   End If
   Exit Sub
   
DiaErr1:
   On Error GoTo 0
   
End Sub

Private Sub ManageBoxes(bEnabled As Boolean)
   On Error Resume Next
   Dim iList As Integer
   For iList = 0 To Controls.Count - 1
      If Err > 0 And (TypeOf Controls(iList) Is TextBox Or _
                      TypeOf Controls(iList) Is ComboBox Or TypeOf Controls(iList) Is MaskEdBox) Then
         If Controls(iList).TabIndex > 2 Then Controls(iList).Enabled = bEnabled
      End If
   Next
   If cmbAct.ListCount > 0 Then cmbAct.Enabled = bEnabled
   tab1.Enabled = bEnabled
   
End Sub


Private Sub UpdateWc()
    If bGoodCenter Then
        MouseCursor 13
        On Error Resume Next
        
        RdoWcn!WCNSUNSH1 = "" & txtSt1(WKDAY_SUN_INDEX)
        RdoWcn!WCNMONSH1 = "" & txtSt1(WKDAY_SUN_INDEX + 1)
        RdoWcn!WCNTUESH1 = "" & txtSt1(WKDAY_SUN_INDEX + 2)
        RdoWcn!WCNWEDSH1 = "" & txtSt1(WKDAY_SUN_INDEX + 3)
        RdoWcn!WCNTHUSH1 = "" & txtSt1(WKDAY_SUN_INDEX + 4)
        RdoWcn!WCNFRISH1 = "" & txtSt1(WKDAY_SUN_INDEX + 5)
        RdoWcn!WCNSATSH1 = "" & txtSt1(WKDAY_SUN_INDEX + 6)
        
        RdoWcn!WCNSUNSH2 = "" & txtSt2(WKDAY_SUN_INDEX)
        RdoWcn!WCNMONSH2 = "" & txtSt2(WKDAY_SUN_INDEX + 1)
        RdoWcn!WCNTUESH2 = "" & txtSt2(WKDAY_SUN_INDEX + 2)
        RdoWcn!WCNWEDSH2 = "" & txtSt2(WKDAY_SUN_INDEX + 3)
        RdoWcn!WCNTHUSH2 = "" & txtSt2(WKDAY_SUN_INDEX + 4)
        RdoWcn!WCNFRISH2 = "" & txtSt2(WKDAY_SUN_INDEX + 5)
        RdoWcn!WCNSATSH2 = "" & txtSt2(WKDAY_SUN_INDEX + 6)
        
        RdoWcn!WCNSUNSH3 = "" & txtSt3(WKDAY_SUN_INDEX)
        RdoWcn!WCNMONSH3 = "" & txtSt3(WKDAY_SUN_INDEX + 1)
        RdoWcn!WCNTUESH3 = "" & txtSt3(WKDAY_SUN_INDEX + 2)
        RdoWcn!WCNWEDSH3 = "" & txtSt3(WKDAY_SUN_INDEX + 3)
        RdoWcn!WCNTHUSH3 = "" & txtSt3(WKDAY_SUN_INDEX + 4)
        RdoWcn!WCNFRISH3 = "" & txtSt3(WKDAY_SUN_INDEX + 5)
        RdoWcn!WCNSATSH3 = "" & txtSt3(WKDAY_SUN_INDEX + 6)
        
        RdoWcn!WCNSUNSH4 = "" & txtSt4(WKDAY_SUN_INDEX)
        RdoWcn!WCNMONSH4 = "" & txtSt4(WKDAY_SUN_INDEX + 1)
        RdoWcn!WCNTUESH4 = "" & txtSt4(WKDAY_SUN_INDEX + 2)
        RdoWcn!WCNWEDSH4 = "" & txtSt4(WKDAY_SUN_INDEX + 3)
        RdoWcn!WCNTHUSH4 = "" & txtSt4(WKDAY_SUN_INDEX + 4)
        RdoWcn!WCNFRISH4 = "" & txtSt4(WKDAY_SUN_INDEX + 5)
        RdoWcn!WCNSATSH4 = "" & txtSt4(WKDAY_SUN_INDEX + 6)
        
        
        
        'Hours
        RdoWcn!WCNSUNHR1 = 0 + Val(txtHr1(WKDAY_SUN_INDEX))
        RdoWcn!WCNMONHR1 = 0 + Val(txtHr1(WKDAY_SUN_INDEX + 1))
        RdoWcn!WCNTUEHR1 = 0 + Val(txtHr1(WKDAY_SUN_INDEX + 2))
        RdoWcn!WCNWEDHR1 = 0 + Val(txtHr1(WKDAY_SUN_INDEX + 3))
        RdoWcn!WCNTHUHR1 = 0 + Val(txtHr1(WKDAY_SUN_INDEX + 4))
        RdoWcn!WCNFRIHR1 = 0 + Val(txtHr1(WKDAY_SUN_INDEX + 5))
        RdoWcn!WCNSATHR1 = 0 + Val(txtHr1(WKDAY_SUN_INDEX + 6))
        
        RdoWcn!WCNSUNHR2 = 0 + Val(txtHr2(WKDAY_SUN_INDEX))
        RdoWcn!WCNMONHR2 = 0 + Val(txtHr2(WKDAY_SUN_INDEX + 1))
        RdoWcn!WCNTUEHR2 = 0 + Val(txtHr2(WKDAY_SUN_INDEX + 2))
        RdoWcn!WCNWEDHR2 = 0 + Val(txtHr2(WKDAY_SUN_INDEX + 3))
        RdoWcn!WCNTHUHR2 = 0 + Val(txtHr2(WKDAY_SUN_INDEX + 4))
        RdoWcn!WCNFRIHR2 = 0 + Val(txtHr2(WKDAY_SUN_INDEX + 5))
        RdoWcn!WCNSATHR2 = 0 + Val(txtHr2(WKDAY_SUN_INDEX + 6))
        
        RdoWcn!WCNSUNHR3 = 0 + Val(txtHr3(WKDAY_SUN_INDEX))
        RdoWcn!WCNMONHR3 = 0 + Val(txtHr3(WKDAY_SUN_INDEX + 1))
        RdoWcn!WCNTUEHR3 = 0 + Val(txtHr3(WKDAY_SUN_INDEX + 2))
        RdoWcn!WCNWEDHR3 = 0 + Val(txtHr3(WKDAY_SUN_INDEX + 3))
        RdoWcn!WCNTHUHR3 = 0 + Val(txtHr3(WKDAY_SUN_INDEX + 4))
        RdoWcn!WCNFRIHR3 = 0 + Val(txtHr3(WKDAY_SUN_INDEX + 5))
        RdoWcn!WCNSATHR3 = 0 + Val(txtHr3(WKDAY_SUN_INDEX + 6))
        
        RdoWcn!WCNSUNHR4 = 0 + Val(txtHr4(WKDAY_SUN_INDEX))
        RdoWcn!WCNMONHR4 = 0 + Val(txtHr4(WKDAY_SUN_INDEX + 1))
        RdoWcn!WCNTUEHR4 = 0 + Val(txtHr4(WKDAY_SUN_INDEX + 2))
        RdoWcn!WCNWEDHR4 = 0 + Val(txtHr4(WKDAY_SUN_INDEX + 3))
        RdoWcn!WCNTHUHR4 = 0 + Val(txtHr4(WKDAY_SUN_INDEX + 4))
        RdoWcn!WCNFRIHR4 = 0 + Val(txtHr4(WKDAY_SUN_INDEX + 5))
        RdoWcn!WCNSATHR4 = 0 + Val(txtHr4(WKDAY_SUN_INDEX + 6))
        
        'Resources
        RdoWcn!WCNSUNMU1 = 0 + Val(txtRe1(WKDAY_SUN_INDEX))
        RdoWcn!WCNMONMU1 = 0 + Val(txtRe1(WKDAY_SUN_INDEX + 1))
        RdoWcn!WCNTUEMU1 = 0 + Val(txtRe1(WKDAY_SUN_INDEX + 2))
        RdoWcn!WCNWEDMU1 = 0 + Val(txtRe1(WKDAY_SUN_INDEX + 3))
        RdoWcn!WCNTHUMU1 = 0 + Val(txtRe1(WKDAY_SUN_INDEX + 4))
        RdoWcn!WCNFRIMU1 = 0 + Val(txtRe1(WKDAY_SUN_INDEX + 5))
        RdoWcn!WCNSATMU1 = 0 + Val(txtRe1(WKDAY_SUN_INDEX + 6))
        
        RdoWcn!WCNSUNMU2 = 0 + Val(txtRe2(WKDAY_SUN_INDEX))
        RdoWcn!WCNMONMU2 = 0 + Val(txtRe2(WKDAY_SUN_INDEX + 1))
        RdoWcn!WCNTUEMU2 = 0 + Val(txtRe2(WKDAY_SUN_INDEX + 2))
        RdoWcn!WCNWEDMU2 = 0 + Val(txtRe2(WKDAY_SUN_INDEX + 3))
        RdoWcn!WCNTHUMU2 = 0 + Val(txtRe2(WKDAY_SUN_INDEX + 4))
        RdoWcn!WCNFRIMU2 = 0 + Val(txtRe2(WKDAY_SUN_INDEX + 5))
        RdoWcn!WCNSATMU2 = 0 + Val(txtRe2(WKDAY_SUN_INDEX + 6))
        
        RdoWcn!WCNSUNMU3 = 0 + Val(txtRe3(WKDAY_SUN_INDEX))
        RdoWcn!WCNMONMU3 = 0 + Val(txtRe3(WKDAY_SUN_INDEX + 1))
        RdoWcn!WCNTUEMU3 = 0 + Val(txtRe3(WKDAY_SUN_INDEX + 2))
        RdoWcn!WCNWEDMU3 = 0 + Val(txtRe3(WKDAY_SUN_INDEX + 3))
        RdoWcn!WCNTHUMU3 = 0 + Val(txtRe3(WKDAY_SUN_INDEX + 4))
        RdoWcn!WCNFRIMU3 = 0 + Val(txtRe3(WKDAY_SUN_INDEX + 5))
        RdoWcn!WCNSATMU3 = 0 + Val(txtRe3(WKDAY_SUN_INDEX + 6))
        
        RdoWcn!WCNSUNMU4 = 0 + Val(txtRe4(WKDAY_SUN_INDEX))
        RdoWcn!WCNMONMU4 = 0 + Val(txtRe4(WKDAY_SUN_INDEX + 1))
        RdoWcn!WCNTUEMU4 = 0 + Val(txtRe4(WKDAY_SUN_INDEX + 2))
        RdoWcn!WCNWEDMU4 = 0 + Val(txtRe4(WKDAY_SUN_INDEX + 3))
        RdoWcn!WCNTHUMU4 = 0 + Val(txtRe4(WKDAY_SUN_INDEX + 4))
        RdoWcn!WCNFRIMU4 = 0 + Val(txtRe4(WKDAY_SUN_INDEX + 5))
        RdoWcn!WCNSATMU4 = 0 + Val(txtRe4(WKDAY_SUN_INDEX + 6))
        
        RdoWcn.Update
        MouseCursor 0
        If Err > 0 Then ValidateEdit
    
    End If

End Sub
