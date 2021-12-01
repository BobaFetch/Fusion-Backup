VERSION 5.00
Begin VB.Form CapaCPe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company Calendar"
   ClientHeight    =   7470
   ClientLeft      =   1950
   ClientTop       =   645
   ClientWidth     =   8160
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   4202
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbYer 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Year"
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cmbMon 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "Month"
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox fraThur 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   6
      Left            =   4608
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   390
      TabStop         =   0   'False
      Top             =   6168
      Width           =   1092
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   127
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   126
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   125
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   124
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
   End
   Begin VB.PictureBox fraTue 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   6
      Left            =   2352
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   389
      TabStop         =   0   'False
      Top             =   6168
      Width           =   1092
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   76
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   77
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   78
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   79
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
   End
   Begin VB.PictureBox fraWed 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   6
      Left            =   3480
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   388
      TabStop         =   0   'False
      Top             =   6168
      Width           =   1092
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   100
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   101
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   102
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   103
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
   End
   Begin VB.PictureBox fraFri 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   6
      Left            =   5736
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   387
      TabStop         =   0   'False
      Top             =   6168
      Width           =   1092
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   148
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   149
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   150
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   151
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
   End
   Begin VB.PictureBox fraSat 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   6
      Left            =   6876
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   386
      TabStop         =   0   'False
      Top             =   6168
      Width           =   1092
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   172
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   173
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   174
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   175
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
   End
   Begin VB.PictureBox fraSun 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   1
      Left            =   120
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   376
      TabStop         =   0   'False
      Top             =   840
      Width           =   1092
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   580
         TabIndex        =   384
         Top             =   760
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   383
         Top             =   760
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   580
         TabIndex        =   382
         Top             =   560
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   381
         Top             =   560
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   580
         TabIndex        =   380
         Top             =   340
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   379
         Top             =   340
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   580
         TabIndex        =   378
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   377
         Top             =   120
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   680
         TabIndex        =   385
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox fraSun 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   2
      Left            =   120
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   366
      Top             =   1896
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   0
         TabIndex        =   374
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   580
         TabIndex        =   373
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   0
         TabIndex        =   372
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   580
         TabIndex        =   371
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   0
         TabIndex        =   370
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   580
         TabIndex        =   369
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   580
         TabIndex        =   368
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   0
         TabIndex        =   367
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   12
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   13
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   14
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   15
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   7
         Left            =   680
         TabIndex        =   375
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraSun 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   3
      Left            =   120
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   356
      Top             =   2964
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   0
         TabIndex        =   364
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   580
         TabIndex        =   363
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   0
         TabIndex        =   362
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   580
         TabIndex        =   361
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   0
         TabIndex        =   360
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   580
         TabIndex        =   359
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   580
         TabIndex        =   358
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   0
         TabIndex        =   357
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   16
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   17
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   18
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   19
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   14
         Left            =   680
         TabIndex        =   365
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraSun 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   4
      Left            =   120
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   346
      Top             =   4032
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   0
         TabIndex        =   354
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   580
         TabIndex        =   353
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   0
         TabIndex        =   352
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   580
         TabIndex        =   351
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   0
         TabIndex        =   350
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   580
         TabIndex        =   349
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   580
         TabIndex        =   348
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   0
         TabIndex        =   347
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   20
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   21
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   22
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   23
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   21
         Left            =   680
         TabIndex        =   355
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraSun 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   5
      Left            =   120
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   336
      Top             =   5100
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   0
         TabIndex        =   344
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   580
         TabIndex        =   343
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   0
         TabIndex        =   342
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   580
         TabIndex        =   341
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   0
         TabIndex        =   340
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   580
         TabIndex        =   339
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   580
         TabIndex        =   338
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   0
         TabIndex        =   337
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   24
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   25
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   26
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   27
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   28
         Left            =   680
         TabIndex        =   345
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraMon 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   1
      Left            =   1236
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   326
      TabStop         =   0   'False
      Top             =   840
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   334
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   580
         TabIndex        =   333
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   332
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   580
         TabIndex        =   331
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   330
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   580
         TabIndex        =   329
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   580
         TabIndex        =   328
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   327
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   28
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   29
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   30
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   31
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   1
         Left            =   680
         TabIndex        =   335
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox fraMon 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   2
      Left            =   1236
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   316
      Top             =   1896
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   0
         TabIndex        =   324
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   580
         TabIndex        =   323
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   0
         TabIndex        =   322
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   580
         TabIndex        =   321
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   0
         TabIndex        =   320
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   580
         TabIndex        =   319
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   580
         TabIndex        =   318
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   0
         TabIndex        =   317
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   36
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   37
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   38
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   39
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   8
         Left            =   680
         TabIndex        =   325
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraMon 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   3
      Left            =   1236
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   306
      Top             =   2964
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   0
         TabIndex        =   314
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   580
         TabIndex        =   313
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   0
         TabIndex        =   312
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   580
         TabIndex        =   311
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   0
         TabIndex        =   310
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   580
         TabIndex        =   309
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   580
         TabIndex        =   308
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   0
         TabIndex        =   307
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   40
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   41
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   42
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   43
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   15
         Left            =   680
         TabIndex        =   315
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraMon 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   4
      Left            =   1236
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   296
      Top             =   4032
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   0
         TabIndex        =   304
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   580
         TabIndex        =   303
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   0
         TabIndex        =   302
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   580
         TabIndex        =   301
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   0
         TabIndex        =   300
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   580
         TabIndex        =   299
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   580
         TabIndex        =   298
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   0
         TabIndex        =   297
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   44
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   45
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   46
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   47
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   22
         Left            =   680
         TabIndex        =   305
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraMon 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   5
      Left            =   1236
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   286
      Top             =   5100
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   0
         TabIndex        =   294
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   580
         TabIndex        =   293
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   0
         TabIndex        =   292
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   580
         TabIndex        =   291
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   0
         TabIndex        =   290
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   580
         TabIndex        =   289
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   580
         TabIndex        =   288
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   0
         TabIndex        =   287
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   48
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   49
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   50
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   51
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   29
         Left            =   680
         TabIndex        =   295
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraTue 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   1
      Left            =   2352
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   276
      TabStop         =   0   'False
      Top             =   840
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   284
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   580
         TabIndex        =   283
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   282
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   580
         TabIndex        =   281
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   280
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   580
         TabIndex        =   279
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   580
         TabIndex        =   278
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   277
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   52
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   53
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   54
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   55
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   132
         Index           =   2
         Left            =   720
         TabIndex        =   285
         Top             =   0
         Width           =   372
      End
   End
   Begin VB.PictureBox fraTue 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   2
      Left            =   2352
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   266
      Top             =   1896
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   0
         TabIndex        =   274
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   580
         TabIndex        =   273
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   0
         TabIndex        =   272
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   580
         TabIndex        =   271
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   0
         TabIndex        =   270
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   580
         TabIndex        =   269
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   580
         TabIndex        =   268
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   0
         TabIndex        =   267
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   60
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   61
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   62
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   63
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   9
         Left            =   680
         TabIndex        =   275
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraTue 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   3
      Left            =   2352
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   256
      Top             =   2964
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   0
         TabIndex        =   264
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   580
         TabIndex        =   263
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   0
         TabIndex        =   262
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   580
         TabIndex        =   261
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   0
         TabIndex        =   260
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   580
         TabIndex        =   259
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   580
         TabIndex        =   258
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   0
         TabIndex        =   257
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   64
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   65
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   66
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   67
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   16
         Left            =   680
         TabIndex        =   265
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraTue 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   4
      Left            =   2352
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   246
      Top             =   4032
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   0
         TabIndex        =   254
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   576
         TabIndex        =   253
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   0
         TabIndex        =   252
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   576
         TabIndex        =   251
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   0
         TabIndex        =   250
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   576
         TabIndex        =   249
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   576
         TabIndex        =   248
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   0
         TabIndex        =   247
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   68
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   69
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   70
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   71
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   23
         Left            =   680
         TabIndex        =   255
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraTue 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   5
      Left            =   2352
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   236
      Top             =   5100
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   0
         TabIndex        =   244
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   580
         TabIndex        =   243
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   0
         TabIndex        =   242
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   580
         TabIndex        =   241
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   0
         TabIndex        =   240
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   580
         TabIndex        =   239
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   580
         TabIndex        =   238
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   0
         TabIndex        =   237
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   72
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   73
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   74
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   75
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   30
         Left            =   680
         TabIndex        =   245
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraWed 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   1
      Left            =   3480
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   226
      TabStop         =   0   'False
      Top             =   840
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   234
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   580
         TabIndex        =   233
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   232
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   580
         TabIndex        =   231
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   230
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   580
         TabIndex        =   229
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   580
         TabIndex        =   228
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   227
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   56
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   57
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   58
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   59
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   3
         Left            =   680
         TabIndex        =   235
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraWed 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   2
      Left            =   3480
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   216
      Top             =   1896
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   0
         TabIndex        =   224
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   580
         TabIndex        =   223
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   0
         TabIndex        =   222
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   580
         TabIndex        =   221
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   0
         TabIndex        =   220
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   580
         TabIndex        =   219
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   580
         TabIndex        =   218
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   0
         TabIndex        =   217
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   84
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   85
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   86
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   87
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   10
         Left            =   680
         TabIndex        =   225
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraWed 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   3
      Left            =   3480
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   206
      Top             =   2964
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   0
         TabIndex        =   214
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   580
         TabIndex        =   213
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   0
         TabIndex        =   212
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   580
         TabIndex        =   211
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   0
         TabIndex        =   210
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   580
         TabIndex        =   209
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   580
         TabIndex        =   208
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   0
         TabIndex        =   207
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   88
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   89
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   90
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   91
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   17
         Left            =   680
         TabIndex        =   215
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraWed 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   4
      Left            =   3480
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   196
      Top             =   4032
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   0
         TabIndex        =   204
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   580
         TabIndex        =   203
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   0
         TabIndex        =   202
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   580
         TabIndex        =   201
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   0
         TabIndex        =   200
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   580
         TabIndex        =   199
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   580
         TabIndex        =   198
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   0
         TabIndex        =   197
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   92
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   93
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   94
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   95
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   24
         Left            =   680
         TabIndex        =   205
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraWed 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   5
      Left            =   3480
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   186
      Top             =   5100
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   0
         TabIndex        =   194
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   580
         TabIndex        =   193
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   0
         TabIndex        =   192
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   580
         TabIndex        =   191
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   0
         TabIndex        =   190
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   580
         TabIndex        =   189
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   580
         TabIndex        =   188
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   0
         TabIndex        =   187
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   96
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   97
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   98
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   99
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   31
         Left            =   680
         TabIndex        =   195
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraSun 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   6
      Left            =   120
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   176
      Top             =   6168
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   0
         TabIndex        =   184
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   580
         TabIndex        =   183
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   0
         TabIndex        =   182
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   580
         TabIndex        =   181
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   0
         TabIndex        =   180
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   580
         TabIndex        =   179
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   580
         TabIndex        =   178
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   0
         TabIndex        =   177
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   104
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   105
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   106
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   107
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   35
         Left            =   680
         TabIndex        =   185
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraMon 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   6
      Left            =   1236
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   166
      Top             =   6168
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   36
         Left            =   0
         TabIndex        =   174
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   36
         Left            =   580
         TabIndex        =   173
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   36
         Left            =   0
         TabIndex        =   172
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   36
         Left            =   580
         TabIndex        =   171
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   36
         Left            =   0
         TabIndex        =   170
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   36
         Left            =   580
         TabIndex        =   169
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   36
         Left            =   580
         TabIndex        =   168
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   36
         Left            =   0
         TabIndex        =   167
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   108
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   109
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   110
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   111
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   36
         Left            =   680
         TabIndex        =   175
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraThur 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   1
      Left            =   4608
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   840
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   164
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   580
         TabIndex        =   163
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   162
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   580
         TabIndex        =   161
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   160
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   580
         TabIndex        =   159
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   580
         TabIndex        =   158
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   157
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   83
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   82
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   81
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   80
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   4
         Left            =   680
         TabIndex        =   165
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraThur 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   2
      Left            =   4608
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   146
      Top             =   1896
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   0
         TabIndex        =   154
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   580
         TabIndex        =   153
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   0
         TabIndex        =   152
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   580
         TabIndex        =   151
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   0
         TabIndex        =   150
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   580
         TabIndex        =   149
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   580
         TabIndex        =   148
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   0
         TabIndex        =   147
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   32
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   33
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   34
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   35
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   132
         Index           =   11
         Left            =   600
         TabIndex        =   155
         Top             =   0
         Width           =   372
      End
   End
   Begin VB.PictureBox fraThur 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   3
      Left            =   4608
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   136
      Top             =   2964
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   0
         TabIndex        =   144
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   580
         TabIndex        =   143
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   0
         TabIndex        =   142
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   580
         TabIndex        =   141
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   0
         TabIndex        =   140
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   580
         TabIndex        =   139
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   580
         TabIndex        =   138
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   0
         TabIndex        =   137
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   112
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   113
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   114
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   115
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   18
         Left            =   680
         TabIndex        =   145
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraThur 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   4
      Left            =   4608
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   126
      Top             =   4032
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   0
         TabIndex        =   134
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   580
         TabIndex        =   133
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   0
         TabIndex        =   132
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   580
         TabIndex        =   131
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   0
         TabIndex        =   130
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   580
         TabIndex        =   129
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   580
         TabIndex        =   128
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   0
         TabIndex        =   127
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   116
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   117
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   118
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   119
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   25
         Left            =   680
         TabIndex        =   135
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraThur 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   5
      Left            =   4608
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   116
      Top             =   5100
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   0
         TabIndex        =   124
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   580
         TabIndex        =   123
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   0
         TabIndex        =   122
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   580
         TabIndex        =   121
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   0
         TabIndex        =   120
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   580
         TabIndex        =   119
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   580
         TabIndex        =   118
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   0
         TabIndex        =   117
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   120
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   121
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   122
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   123
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   32
         Left            =   680
         TabIndex        =   125
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraFri 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   2
      Left            =   5760
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   106
      Top             =   1896
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   0
         TabIndex        =   114
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   580
         TabIndex        =   113
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   0
         TabIndex        =   112
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   580
         TabIndex        =   111
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   0
         TabIndex        =   110
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   580
         TabIndex        =   109
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   580
         TabIndex        =   108
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   0
         TabIndex        =   107
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   132
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   133
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   134
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   135
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   12
         Left            =   680
         TabIndex        =   115
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraFri 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   3
      Left            =   5760
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   96
      Top             =   2964
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   0
         TabIndex        =   104
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   576
         TabIndex        =   103
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   0
         TabIndex        =   102
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   576
         TabIndex        =   101
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   0
         TabIndex        =   100
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   576
         TabIndex        =   99
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   576
         TabIndex        =   98
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   0
         TabIndex        =   97
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   136
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   137
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   138
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   139
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   19
         Left            =   680
         TabIndex        =   105
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraFri 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   4
      Left            =   5760
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   86
      Top             =   4032
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   0
         TabIndex        =   94
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   580
         TabIndex        =   93
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   0
         TabIndex        =   92
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   580
         TabIndex        =   91
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   0
         TabIndex        =   90
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   580
         TabIndex        =   89
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   580
         TabIndex        =   88
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   0
         TabIndex        =   87
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   140
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   141
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   142
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   143
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   26
         Left            =   680
         TabIndex        =   95
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraFri 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   5
      Left            =   5736
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   76
      Top             =   5100
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   0
         TabIndex        =   84
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   580
         TabIndex        =   83
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   0
         TabIndex        =   82
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   580
         TabIndex        =   81
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   0
         TabIndex        =   80
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   580
         TabIndex        =   79
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   580
         TabIndex        =   78
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   0
         TabIndex        =   77
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   144
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   145
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   146
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   147
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   33
         Left            =   680
         TabIndex        =   85
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraSat 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   1
      Left            =   6876
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   840
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   74
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   580
         TabIndex        =   73
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   72
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   580
         TabIndex        =   71
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   70
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   580
         TabIndex        =   69
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   580
         TabIndex        =   68
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   67
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   128
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   129
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   130
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   131
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   6
         Left            =   680
         TabIndex        =   75
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraSat 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   2
      Left            =   6876
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   56
      Top             =   1896
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   0
         TabIndex        =   64
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   580
         TabIndex        =   63
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   0
         TabIndex        =   62
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   580
         TabIndex        =   61
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   0
         TabIndex        =   60
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   580
         TabIndex        =   59
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   580
         TabIndex        =   58
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   0
         TabIndex        =   57
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   156
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   157
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   158
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   159
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1620
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   13
         Left            =   680
         TabIndex        =   65
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraSat 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   3
      Left            =   6876
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   46
      Top             =   2952
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   0
         TabIndex        =   54
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   580
         TabIndex        =   53
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   0
         TabIndex        =   52
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   580
         TabIndex        =   51
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   0
         TabIndex        =   50
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   580
         TabIndex        =   49
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   580
         TabIndex        =   48
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   0
         TabIndex        =   47
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   160
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   161
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   162
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   163
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   20
         Left            =   680
         TabIndex        =   55
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraSat 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   4
      Left            =   6876
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   36
      Top             =   4032
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   0
         TabIndex        =   44
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   580
         TabIndex        =   43
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   0
         TabIndex        =   42
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   580
         TabIndex        =   41
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   0
         TabIndex        =   40
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   580
         TabIndex        =   39
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   580
         TabIndex        =   38
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   0
         TabIndex        =   37
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   164
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   165
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   166
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   167
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   27
         Left            =   680
         TabIndex        =   45
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraSat 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   5
      Left            =   6876
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   26
      Top             =   5100
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   0
         TabIndex        =   34
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   580
         TabIndex        =   33
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   0
         TabIndex        =   32
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   580
         TabIndex        =   31
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   0
         TabIndex        =   30
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   580
         TabIndex        =   29
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   580
         TabIndex        =   28
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   0
         TabIndex        =   27
         Top             =   765
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   168
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   169
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   170
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   171
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   34
         Left            =   680
         TabIndex        =   35
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraFri 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   1
      Left            =   5736
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   840
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   24
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   580
         TabIndex        =   23
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   22
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   580
         TabIndex        =   21
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   20
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   580
         TabIndex        =   19
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   580
         TabIndex        =   18
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   17
         Top             =   768
         Width           =   580
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   11
         X1              =   0
         X2              =   1092
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line z6 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   1080
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   9
         X1              =   0
         X2              =   1092
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line z6 
         BorderColor     =   &H00808080&
         Index           =   8
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   5
         Left            =   680
         TabIndex        =   25
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPe04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdTem 
      Caption         =   "&Template"
      Height          =   350
      Left            =   5400
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Reload Current Template"
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6240
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Press To Save or Update Calendar"
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   800
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      Height          =   252
      Index           =   8
      Left            =   3360
      TabIndex        =   12
      Top             =   120
      Width           =   960
   End
   Begin VB.Label lblFrom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   3960
      TabIndex        =   11
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saturday"
      Height          =   255
      Index           =   7
      Left            =   6876
      TabIndex        =   10
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Friday"
      Height          =   255
      Index           =   6
      Left            =   5736
      TabIndex        =   9
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thursday"
      Height          =   255
      Index           =   5
      Left            =   4608
      TabIndex        =   8
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wednesday"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   7
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tuesday"
      Height          =   255
      Index           =   3
      Left            =   2352
      TabIndex        =   6
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monday"
      Height          =   255
      Index           =   2
      Left            =   1236
      TabIndex        =   5
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sunday"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Month/Year"
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "CapaCPe04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'6/12/06 Revised ToolTipText
Option Explicit
Dim bOnLoad As Byte
Dim bGoodCalendar As Byte
Dim bGoodTemplate As Byte
Dim iStartDay As Integer
Dim vShifts(8, 9)

Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter


Private Sub cmbMon_Click()
   cmdSave.Enabled = True
   GetThisMonth
   
End Sub

Private Sub cmbMon_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub cmbYer_Click()
   cmdSave.Enabled = True
   GetThisMonth
   
End Sub

Private Sub cmbYer_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4202
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSave_Click()
   Dim iList As Integer
   Dim K As Integer
   If Not bGoodCalendar Then
      MsgBox "The Calendar Can't Be Saved.", vbInformation, Caption
      Exit Sub
   End If
   MouseCursor 11
   cmdSave.Enabled = False
   On Error Resume Next
   clsADOCon.ExecuteSQL "DELETE FROM CoclTable WHERE COCREF='" & cmbMon & "-" & cmbYer & "'"
   On Error GoTo CccalSv1
   For iList = 0 To 6
      If lblDte(iList).Visible Then Exit For
   Next
   iStartDay = iList
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   For iList = iStartDay To 36
      If Not lblDte(iList).Visible Then Exit For
      K = K + 1
      sSql = "INSERT INTO CoclTable (COCREF,COCDAY," _
             & "COCSHS1,COCSHS2,COCSHS3,COCSHS4," _
             & "COCSHT1,COCSHT2,COCSHT3,COCSHT4) " _
             & "VALUES('" & cmbMon & "-" & cmbYer & "'," _
             & str(K) & ",'" _
             & txtS1s(iList) & "','" _
             & txtS2s(iList) & "','" _
             & txtS3s(iList) & "','" _
             & txtS4s(iList) & "'," _
             & Val(txtS1t(iList)) & "," _
             & Val(txtS2t(iList)) & "," _
             & Val(txtS3t(iList)) & "," _
             & Val(txtS4t(iList)) & ") "
      clsADOCon.ExecuteSQL sSql
   Next
   MouseCursor 0
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      SysMsg "The Calendar Was Saved.", True, Me
      lblFrom = "Saved Calendar"
   Else
      clsADOCon.RollbackTrans
      MsgBox "Unable To Successfully Update The Calendar", _
         vbExclamation, Caption
   End If
   cmdSave.Enabled = True
   Exit Sub
   
CccalSv1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume CccalSv2
CccalSv2:
   MouseCursor 0
   On Error Resume Next
   clsADOCon.RollbackTrans
   MsgBox Trim(str(CurrError.Number)) & vbCr & CurrError.Description, vbExclamation, Caption
   
End Sub

Private Sub cmdTem_Click()
   Dim sNewShop As String
   Dim bResponse As Byte
   If bGoodCalendar Then
      bResponse = MsgBox("Refill From Current Template?", ES_YESQUESTION, Caption)
      If bResponse = vbYes Then bGoodCalendar = GetTheCalendar(True)
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      LoadCombos
      bOnLoad = 0
   End If
   
End Sub

Private Sub Form_Load()
   Dim A As Integer
   Dim iList As Integer
   FormLoad Me, ES_LIST, ES_DONTRESIZE
   Move 0, 0
   A = -1
   For iList = 0 To 36
      txtS1s(iList).ToolTipText = "Shift Hours-Enter As 8.0"
      A = A + 1
      txtS1s(iList).TabIndex = A
      
      txtS1t(iList).ToolTipText = "Shift Resources As 2.5"
      A = A + 1
      txtS1t(iList).TabIndex = A
      
      txtS2s(iList).ToolTipText = "Shift Hours-Enter As 8.0"
      A = A + 1
      txtS2s(iList).TabIndex = A
      
      txtS2t(iList).ToolTipText = "Shift Resources As 2.5"
      A = A + 1
      txtS2t(iList).TabIndex = A
      
      txtS3s(iList).ToolTipText = "Shift Hours-Enter As 8.0"
      A = A + 1
      txtS3s(iList).TabIndex = A
      
      txtS3t(iList).ToolTipText = "Shift Resources As 2.5"
      A = A + 1
      txtS3t(iList).TabIndex = A
      
      txtS4s(iList).ToolTipText = "Shift Hours-Enter As 8.0"
      A = A + 1
      txtS4s(iList).TabIndex = A
      
      txtS4t(iList).ToolTipText = "Shift Resources As 2.5"
      A = A + 1
      txtS4t(iList).TabIndex = A
   Next
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set CapaCPe04a = Nothing
   
End Sub




Private Sub txtS1s_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtS1s_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtS1s_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub


Private Sub txtS1s_LostFocus(Index As Integer)
   txtS1s(Index) = CheckLen(txtS1s(Index), 6)
   Dim tc As New ClassTimeCharge
   txtS1s(Index) = tc.GetTime(txtS1s(Index))
   
End Sub

Private Sub txtS1t_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtS1t_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtS1t_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtS1t_LostFocus(Index As Integer)
   txtS1t(Index) = CheckLen(txtS1t(Index), 4)
   txtS1t(Index) = Format(Abs(Val(txtS1t(Index))), "#0.0")
   
End Sub

Private Sub txtS2s_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtS2s_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtS2s_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub


Private Sub txtS2s_LostFocus(Index As Integer)
   txtS2s(Index) = CheckLen(txtS2s(Index), 6)
   Dim tc As New ClassTimeCharge
   txtS2s(Index) = tc.GetTime(txtS2s(Index))
   
End Sub

Private Sub txtS2t_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtS2t_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtS2t_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtS2t_LostFocus(Index As Integer)
   txtS2t(Index) = CheckLen(txtS2t(Index), 4)
   txtS2t(Index) = Format(Abs(Val(txtS2t(Index))), "#0.0")
   
End Sub

Private Sub txtS3s_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtS3s_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtS3s_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub

Private Sub txtS3s_LostFocus(Index As Integer)
   txtS3s(Index) = CheckLen(txtS3s(Index), 6)
   Dim tc As New ClassTimeCharge
   txtS3s(Index) = tc.GetTime(txtS3s(Index))
   
End Sub

Private Sub txtS3t_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtS3t_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtS3t_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtS3t_LostFocus(Index As Integer)
   txtS3t(Index) = CheckLen(txtS3t(Index), 4)
   txtS3t(Index) = Format(Abs(Val(txtS3t(Index))), "#0.0")
   
End Sub

Private Sub txtS4s_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtS4s_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtS4s_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyTime KeyAscii
   
End Sub

Private Sub txtS4s_LostFocus(Index As Integer)
   txtS4s(Index) = CheckLen(txtS4s(Index), 6)
   Dim tc As New ClassTimeCharge
   txtS4s(Index) = tc.GetTime(txtS4s(Index))
   
End Sub

Private Sub txtS4t_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtS4t_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtS4t_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtS4t_LostFocus(Index As Integer)
   txtS4t(Index) = CheckLen(txtS4t(Index), 4)
   txtS4t(Index) = Format(Abs(Val(txtS4t(Index))), "#0.0")
   
End Sub



Private Sub GetThisMonth()
   Dim A As Integer
   Dim iList As Integer
   Dim n As Integer
   Dim ThisMonth As Date
   Dim iStartMonth As Integer
   sSql = Format$(cmbMon) & " 1, " & Format$(cmbYer)
   ThisMonth = DateValue(sSql)
   iStartMonth = Format(ThisMonth, "w")
   For iList = 0 To 4
      lblDte(iList) = 0
      lblDte(iList).Visible = False
      txtS1s(iList).Visible = False
      txtS1t(iList).Visible = False
      txtS2s(iList).Visible = False
      txtS2t(iList).Visible = False
      txtS3s(iList).Visible = False
      txtS3t(iList).Visible = False
      txtS4s(iList).Visible = False
      txtS4t(iList).Visible = False
   Next
   lblDte(iList) = 0
   lblDte(iList).Visible = False
   txtS1s(iList).Visible = False
   txtS1t(iList).Visible = False
   txtS2s(iList).Visible = False
   txtS2t(iList).Visible = False
   txtS3s(iList).Visible = False
   txtS3t(iList).Visible = False
   txtS4s(iList).Visible = False
   txtS4t(iList).Visible = False
   For iList = 28 To 35
      lblDte(iList) = 0
      lblDte(iList).Visible = False
      txtS1s(iList).Visible = False
      txtS1t(iList).Visible = False
      txtS2s(iList).Visible = False
      txtS2t(iList).Visible = False
      txtS3s(iList).Visible = False
      txtS3t(iList).Visible = False
      txtS4s(iList).Visible = False
      txtS4t(iList).Visible = False
   Next
   lblDte(iList) = 0
   lblDte(iList).Visible = False
   txtS1s(iList).Visible = False
   txtS1t(iList).Visible = False
   txtS2s(iList).Visible = False
   txtS2t(iList).Visible = False
   txtS3s(iList).Visible = False
   txtS3t(iList).Visible = False
   txtS4s(iList).Visible = False
   txtS4t(iList).Visible = False
   sSql = Format$(cmbMon) & " 1, " & Format$(cmbYer)
   ThisMonth = DateValue(sSql)
   iStartMonth = Format(ThisMonth, "w")
   iList = iStartMonth - 1
   lblDte(iList) = 1
   lblDte(iList).Visible = True
   txtS1s(iList).Visible = True
   txtS1t(iList).Visible = True
   txtS2s(iList).Visible = True
   txtS2t(iList).Visible = True
   txtS3s(iList).Visible = True
   txtS3t(iList).Visible = True
   txtS4s(iList).Visible = True
   txtS4t(iList).Visible = True
   A = 1
   n = 1
   Do Until Format(ThisMonth + A, "mmm") <> cmbMon
      A = A + 1
      iList = iList + 1
      n = n + 1
      lblDte(iList) = n
      lblDte(iList).Visible = True
      txtS1s(iList).Visible = True
      txtS1t(iList).Visible = True
      txtS2s(iList).Visible = True
      txtS2t(iList).Visible = True
      txtS3s(iList).Visible = True
      txtS3t(iList).Visible = True
      txtS4s(iList).Visible = True
      txtS4t(iList).Visible = True
   Loop
   bGoodCalendar = GetTheCalendar(False)
   
End Sub

Private Function GetCalTemplate() As Byte
   Dim rdoTmp As ADODB.Recordset
   Dim sMsg As String
   On Error GoTo DiaErr1
   GetCalTemplate = False
   sSql = "SELECT * FROM CctmTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTmp)
   If bSqlRows Then
      With rdoTmp
         If !CALTOTHRS = 0 Then
            GetCalTemplate = False
         Else
            GetCalTemplate = True
            vShifts(1, 1) = "" & Trim(!CALSUNST1)
            vShifts(1, 2) = "" & Trim(!CALSUNST2)
            vShifts(1, 3) = "" & Trim(!CALSUNST3)
            vShifts(1, 4) = "" & Trim(!CALSUNST4)
            vShifts(1, 5) = 0 + !CALSUNHR1
            vShifts(1, 6) = 0 + !CALSUNHR2
            vShifts(1, 7) = 0 + !CALSUNHR3
            vShifts(1, 8) = 0 + !CALSUNHR4
            
            vShifts(2, 1) = "" & Trim(!CALMONST1)
            vShifts(2, 2) = "" & Trim(!CALMONST2)
            vShifts(2, 3) = "" & Trim(!CALMONST3)
            vShifts(2, 4) = "" & Trim(!CALMONST4)
            vShifts(2, 5) = 0 + !CALMONHR1
            vShifts(2, 6) = 0 + !CALMONHR2
            vShifts(2, 7) = 0 + !CALMONHR3
            vShifts(2, 8) = 0 + !CALMONHR4
            
            vShifts(3, 1) = "" & Trim(!CALTUEST1)
            vShifts(3, 2) = "" & Trim(!CALTUEST2)
            vShifts(3, 3) = "" & Trim(!CALTUEST3)
            vShifts(3, 4) = "" & Trim(!CALTUEST4)
            vShifts(3, 5) = 0 + !CALTUEHR1
            vShifts(3, 6) = 0 + !CALTUEHR2
            vShifts(3, 7) = 0 + !CALTUEHR3
            vShifts(3, 8) = 0 + !CALTUEHR4
            
            vShifts(4, 1) = "" & Trim(!CALWEDST1)
            vShifts(4, 2) = "" & Trim(!CALWEDST2)
            vShifts(4, 3) = "" & Trim(!CALWEDST3)
            vShifts(4, 4) = "" & Trim(!CALWEDST4)
            vShifts(4, 5) = 0 + !CALWEDHR1
            vShifts(4, 6) = 0 + !CALWEDHR2
            vShifts(4, 7) = 0 + !CALWEDHR3
            vShifts(4, 8) = 0 + !CALWEDHR4
            
            vShifts(5, 1) = "" & Trim(!CALTHUST1)
            vShifts(5, 2) = "" & Trim(!CALTHUST2)
            vShifts(5, 3) = "" & Trim(!CALTHUST3)
            vShifts(5, 4) = "" & Trim(!CALTHUST4)
            vShifts(5, 5) = 0 + !CALTHUHR1
            vShifts(5, 6) = 0 + !CALTHUHR2
            vShifts(5, 7) = 0 + !CALTHUHR3
            vShifts(5, 8) = 0 + !CALTHUHR4
            
            vShifts(6, 1) = "" & Trim(!CALFRIST1)
            vShifts(6, 2) = "" & Trim(!CALFRIST2)
            vShifts(6, 3) = "" & Trim(!CALFRIST3)
            vShifts(6, 4) = "" & Trim(!CALFRIST4)
            vShifts(6, 5) = 0 + !CALFRIHR1
            vShifts(6, 6) = 0 + !CALFRIHR2
            vShifts(6, 7) = 0 + !CALFRIHR3
            vShifts(6, 8) = 0 + !CALFRIHR4
            vShifts(7, 1) = "" & Trim(!CALSATST1)
            vShifts(7, 2) = "" & Trim(!CALSATST2)
            vShifts(7, 3) = "" & Trim(!CALSATST3)
            vShifts(7, 4) = "" & Trim(!CALSATST4)
            vShifts(7, 5) = 0 + !CALSATHR1
            vShifts(7, 6) = 0 + !CALSATHR2
            vShifts(7, 7) = 0 + !CALSATHR3
            vShifts(7, 8) = 0 + !CALSATHR4
         End If
      End With
   Else
      GetCalTemplate = False
   End If
   If Not GetCalTemplate Then
      MouseCursor 0
      sMsg = "Your Calendar Template Is Not Setup Or Has No Hours," & vbCr _
             & "Changes Will Not Be Recorded."
      MsgBox sMsg, vbExclamation, Caption
      bGoodCalendar = False
      cmdSave.Enabled = False
   End If
   cmbYer.Enabled = True
   Set rdoTmp = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcaltemp"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetTheCalendar(bReloadTemplate) As Byte
   Dim RdoCal As ADODB.Recordset
   Dim A As Integer
   Dim iList As Integer
   Dim K As Integer
   MouseCursor 13
   
   'close some boxes to avoid recursion
   On Error GoTo DiaErr1
   cmbYer.Enabled = False
   cmdSave.Enabled = False
   GetTheCalendar = False
   For iList = 0 To 6
      If lblDte(iList).Visible Then Exit For
   Next
   iStartDay = iList + 1
   If Not bReloadTemplate Then
      AdoQry.Parameters(0).Value = cmbMon & "-" & cmbYer
      bSqlRows = clsADOCon.GetQuerySet(RdoCal, AdoQry)
      If bSqlRows Then
         K = iStartDay - 2
         With RdoCal
            Do Until .EOF
               K = K + 1
               txtS1s(K) = "" & Trim(!COCSHS1)
               txtS2s(K) = "" & Trim(!COCSHS2)
               txtS3s(K) = "" & Trim(!COCSHS3)
               txtS4s(K) = "" & Trim(!COCSHS4)
               txtS1t(K) = Format(!COCSHT1, "#0.0")
               txtS2t(K) = Format(!COCSHT2, "#0.0")
               txtS3t(K) = Format(!COCSHT3, "#0.0")
               txtS4t(K) = Format(!COCSHT4, "#0.0")
               .MoveNext
            Loop
            ClearResultSet RdoCal
         End With
         GetTheCalendar = True
         lblFrom = "Saved Calendar"
         cmdSave.Enabled = True
         MouseCursor 0
         cmbYer.Enabled = True
         Set RdoCal = Nothing
         Exit Function
      End If
   End If
   lblFrom = "Calendar Template"
   For A = iStartDay To 6
      txtS1s(A - 1) = vShifts(A, 1)
      txtS2s(A - 1) = vShifts(A, 2)
      txtS3s(A - 1) = vShifts(A, 3)
      txtS4s(A - 1) = vShifts(A, 4)
      
      txtS1t(A - 1) = Format(vShifts(A, 5), "#0.0")
      txtS2t(A - 1) = Format(vShifts(A, 6), "#0.0")
      txtS3t(A - 1) = Format(vShifts(A, 7), "#0.0")
      txtS4t(A - 1) = Format(vShifts(A, 8), "#0.0")
   Next
   txtS1s(A - 1) = vShifts(A, 1)
   txtS2s(A - 1) = vShifts(A, 2)
   txtS3s(A - 1) = vShifts(A, 3)
   txtS4s(A - 1) = vShifts(A, 4)
   
   txtS1t(A - 1) = Format(vShifts(A, 5), "#0.0")
   txtS2t(A - 1) = Format(vShifts(A, 6), "#0.0")
   txtS3t(A - 1) = Format(vShifts(A, 7), "#0.0")
   txtS4t(A - 1) = Format(vShifts(A, 8), "#0.0")
   
   iList = A - 1
   K = iList
   For A = iList To 36
      K = K + 1
      If K > 7 Then K = 1
      If Not lblDte(A).Visible Then Exit For
      txtS1s(A) = vShifts(K, 1)
      txtS2s(A) = vShifts(K, 2)
      txtS3s(A) = vShifts(K, 3)
      txtS4s(A) = vShifts(K, 4)
      
      txtS1t(A) = Format(vShifts(K, 5), "#0.0")
      txtS2t(A) = Format(vShifts(K, 6), "#0.0")
      txtS3t(A) = Format(vShifts(K, 7), "#0.0")
      txtS4t(A) = Format(vShifts(K, 8), "#0.0")
   Next
   GetTheCalendar = True
   cmdSave.Enabled = True
   cmbYer.Enabled = True
   MouseCursor 0
   Set RdoCal = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getthecal"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   cmbYer.Enabled = True
   cmdSave.Enabled = False
   GetTheCalendar = False
   DoModuleErrors Me
   
End Function

Private Sub LoadCombos()
   Dim iList As Integer
   Dim A As Integer
   Dim vMonth As Variant
   
   On Error Resume Next
   MouseCursor 13
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM CoclTable WHERE COCREF= ? "
   
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 8
   
   AdoQry.Parameters.Append AdoParameter
   
   cmbMon.AddItem "Jan"
   cmbMon.AddItem "Feb"
   cmbMon.AddItem "Mar"
   cmbMon.AddItem "Apr"
   cmbMon.AddItem "May"
   cmbMon.AddItem "Jun"
   cmbMon.AddItem "Jul"
   cmbMon.AddItem "Aug"
   cmbMon.AddItem "Sep"
   cmbMon.AddItem "Oct"
   cmbMon.AddItem "Nov"
   cmbMon.AddItem "Dec"
   cmbMon = Format(Now, "mmm")
   A = Format(Now, "yyyy")
   For iList = A - 2 To A + 25
      AddComboStr cmbYer.hwnd, Format$(iList)
   Next
   cmbYer = Format$(Now, "yyyy")
   For iList = 0 To 35
      txtS1s(iList).ToolTipText = "Shift Start-Enter As 8.00a"
      txtS1t(iList).ToolTipText = "Shift Hours Enter As 2.5"
      
      txtS2s(iList).ToolTipText = "Shift Start-Enter As 8.00a"
      txtS2t(iList).ToolTipText = "Shift Hours Enter As 2.5"
      
      txtS3s(iList).ToolTipText = "Shift Start-Enter As 8.00a"
      txtS3t(iList).ToolTipText = "Shift Hours Enter As 2.5"
      
      txtS4s(iList).ToolTipText = "Shift Start-Enter As 8.00a"
      txtS4t(iList).ToolTipText = "Shift Hours Enter As 2.5"
   Next
   txtS1s(iList).ToolTipText = "Shift Start-Enter As 8.00a"
   txtS1t(iList).ToolTipText = "Shift Hours Enter As 2.5"
   
   txtS2s(iList).ToolTipText = "Shift Start-Enter As 8.00a"
   txtS2t(iList).ToolTipText = "Shift Hours Enter As 2.5"
   
   txtS3s(iList).ToolTipText = "Shift Start-Enter As 8.00a"
   txtS3t(iList).ToolTipText = "Shift Hours Enter As 2.5"
   
   txtS4s(iList).ToolTipText = "Shift Start-Enter As 8.00a"
   txtS4t(iList).ToolTipText = "Shift Hours Enter As 2.5"
   GetThisMonth
   bGoodTemplate = GetCalTemplate
   If bGoodTemplate Then bGoodCalendar = GetTheCalendar(False)
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "loadcombos"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
