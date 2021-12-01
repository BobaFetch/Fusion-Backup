VERSION 5.00
Begin VB.Form CapaCPe03b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work Center Calendar"
   ClientHeight    =   7380
   ClientLeft      =   1950
   ClientTop       =   645
   ClientWidth     =   8160
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   4201
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox fraThur 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   6
      Left            =   4608
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   393
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
   Begin VB.PictureBox fraSun 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   1
      Left            =   120
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   391
      Top             =   840
      Width           =   1092
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   580
         TabIndex        =   9
         Top             =   760
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   760
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   580
         TabIndex        =   7
         Top             =   560
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   560
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   580
         TabIndex        =   5
         Top             =   340
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   340
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   580
         TabIndex        =   3
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   2
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
         TabIndex        =   392
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
      TabIndex        =   381
      Top             =   1896
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   0
         TabIndex        =   389
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   580
         TabIndex        =   388
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   0
         TabIndex        =   387
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   580
         TabIndex        =   386
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   0
         TabIndex        =   385
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   580
         TabIndex        =   384
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   580
         TabIndex        =   383
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   0
         TabIndex        =   382
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
         TabIndex        =   390
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
      TabIndex        =   371
      Top             =   2964
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   0
         TabIndex        =   379
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   580
         TabIndex        =   378
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   0
         TabIndex        =   377
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   580
         TabIndex        =   376
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   0
         TabIndex        =   375
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   580
         TabIndex        =   374
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   580
         TabIndex        =   373
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   0
         TabIndex        =   372
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
         TabIndex        =   380
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
      TabIndex        =   361
      Top             =   4032
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   0
         TabIndex        =   369
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   580
         TabIndex        =   368
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   0
         TabIndex        =   367
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   580
         TabIndex        =   366
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   0
         TabIndex        =   365
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   580
         TabIndex        =   364
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   580
         TabIndex        =   363
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   0
         TabIndex        =   362
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
         TabIndex        =   370
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
      TabIndex        =   351
      Top             =   5100
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   0
         TabIndex        =   359
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   580
         TabIndex        =   358
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   0
         TabIndex        =   357
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   580
         TabIndex        =   356
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   0
         TabIndex        =   355
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   580
         TabIndex        =   354
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   580
         TabIndex        =   353
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   28
         Left            =   0
         TabIndex        =   352
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
         TabIndex        =   360
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
      TabIndex        =   349
      Top             =   840
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   580
         TabIndex        =   11
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   580
         TabIndex        =   13
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   14
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   580
         TabIndex        =   15
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   580
         TabIndex        =   17
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   16
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
         TabIndex        =   350
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
      TabIndex        =   339
      Top             =   1896
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   0
         TabIndex        =   347
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   580
         TabIndex        =   346
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   0
         TabIndex        =   345
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   580
         TabIndex        =   344
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   0
         TabIndex        =   343
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   580
         TabIndex        =   342
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   580
         TabIndex        =   341
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   0
         TabIndex        =   340
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
         TabIndex        =   348
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
      TabIndex        =   329
      Top             =   2964
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   0
         TabIndex        =   337
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   580
         TabIndex        =   336
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   0
         TabIndex        =   335
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   580
         TabIndex        =   334
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   0
         TabIndex        =   333
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   580
         TabIndex        =   332
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   580
         TabIndex        =   331
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   15
         Left            =   0
         TabIndex        =   330
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
         TabIndex        =   338
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
      TabIndex        =   319
      Top             =   4032
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   0
         TabIndex        =   327
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   580
         TabIndex        =   326
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   0
         TabIndex        =   325
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   580
         TabIndex        =   324
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   0
         TabIndex        =   323
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   580
         TabIndex        =   322
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   580
         TabIndex        =   321
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   22
         Left            =   0
         TabIndex        =   320
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
         TabIndex        =   328
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
      TabIndex        =   309
      Top             =   5100
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   0
         TabIndex        =   317
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   580
         TabIndex        =   316
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   0
         TabIndex        =   315
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   580
         TabIndex        =   314
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   0
         TabIndex        =   313
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   580
         TabIndex        =   312
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   580
         TabIndex        =   311
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   29
         Left            =   0
         TabIndex        =   310
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
         TabIndex        =   318
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
      TabIndex        =   299
      Top             =   840
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   307
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   580
         TabIndex        =   306
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   305
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   580
         TabIndex        =   304
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   303
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   580
         TabIndex        =   302
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   580
         TabIndex        =   301
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   300
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
         TabIndex        =   308
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
      TabIndex        =   289
      Top             =   1896
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   0
         TabIndex        =   297
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   580
         TabIndex        =   296
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   0
         TabIndex        =   295
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   580
         TabIndex        =   294
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   0
         TabIndex        =   293
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   580
         TabIndex        =   292
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   580
         TabIndex        =   291
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   0
         TabIndex        =   290
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
         TabIndex        =   298
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
      TabIndex        =   279
      Top             =   2964
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   0
         TabIndex        =   287
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   580
         TabIndex        =   286
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   0
         TabIndex        =   285
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   580
         TabIndex        =   284
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   0
         TabIndex        =   283
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   580
         TabIndex        =   282
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   580
         TabIndex        =   281
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   16
         Left            =   0
         TabIndex        =   280
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
         TabIndex        =   288
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
      TabIndex        =   269
      Top             =   4032
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   0
         TabIndex        =   277
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   576
         TabIndex        =   276
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   0
         TabIndex        =   275
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   576
         TabIndex        =   274
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   0
         TabIndex        =   273
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   576
         TabIndex        =   272
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   576
         TabIndex        =   271
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   23
         Left            =   0
         TabIndex        =   270
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
         TabIndex        =   278
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
      TabIndex        =   259
      Top             =   5100
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   0
         TabIndex        =   267
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   580
         TabIndex        =   266
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   0
         TabIndex        =   265
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   580
         TabIndex        =   264
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   0
         TabIndex        =   263
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   580
         TabIndex        =   262
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   580
         TabIndex        =   261
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   30
         Left            =   0
         TabIndex        =   260
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
         TabIndex        =   268
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraTue 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   6
      Left            =   2352
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   258
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
      Index           =   1
      Left            =   3480
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   248
      Top             =   840
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   256
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   580
         TabIndex        =   255
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   254
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   580
         TabIndex        =   253
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   252
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   580
         TabIndex        =   251
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   580
         TabIndex        =   250
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   249
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
         TabIndex        =   257
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
      TabIndex        =   238
      Top             =   1896
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   0
         TabIndex        =   246
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   580
         TabIndex        =   245
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   0
         TabIndex        =   244
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   580
         TabIndex        =   243
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   0
         TabIndex        =   242
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   580
         TabIndex        =   241
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   580
         TabIndex        =   240
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   0
         TabIndex        =   239
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
         TabIndex        =   247
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
      TabIndex        =   228
      Top             =   2964
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   0
         TabIndex        =   236
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   580
         TabIndex        =   235
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   0
         TabIndex        =   234
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   580
         TabIndex        =   233
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   0
         TabIndex        =   232
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   580
         TabIndex        =   231
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   580
         TabIndex        =   230
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   17
         Left            =   0
         TabIndex        =   229
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
         TabIndex        =   237
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
      TabIndex        =   218
      Top             =   4032
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   0
         TabIndex        =   226
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   580
         TabIndex        =   225
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   0
         TabIndex        =   224
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   580
         TabIndex        =   223
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   0
         TabIndex        =   222
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   580
         TabIndex        =   221
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   580
         TabIndex        =   220
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   24
         Left            =   0
         TabIndex        =   219
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
         TabIndex        =   227
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
      TabIndex        =   208
      Top             =   5100
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   0
         TabIndex        =   216
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   580
         TabIndex        =   215
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   0
         TabIndex        =   214
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   580
         TabIndex        =   213
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   0
         TabIndex        =   212
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   580
         TabIndex        =   211
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   580
         TabIndex        =   210
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   31
         Left            =   0
         TabIndex        =   209
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
         TabIndex        =   217
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraWed 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   6
      Left            =   3480
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   207
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
   Begin VB.PictureBox fraSun 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   6
      Left            =   120
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   197
      Top             =   6168
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   0
         TabIndex        =   205
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   580
         TabIndex        =   204
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   0
         TabIndex        =   203
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   580
         TabIndex        =   202
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   0
         TabIndex        =   201
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   580
         TabIndex        =   200
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   580
         TabIndex        =   199
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   35
         Left            =   0
         TabIndex        =   198
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
         TabIndex        =   206
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
      TabIndex        =   187
      Top             =   6168
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   36
         Left            =   0
         TabIndex        =   195
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
         TabIndex        =   194
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   36
         Left            =   0
         TabIndex        =   193
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
         TabIndex        =   192
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   36
         Left            =   0
         TabIndex        =   191
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
         TabIndex        =   190
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
         TabIndex        =   189
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   36
         Left            =   0
         TabIndex        =   188
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
         TabIndex        =   196
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
      TabIndex        =   177
      Top             =   840
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   185
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   580
         TabIndex        =   184
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   183
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   580
         TabIndex        =   182
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   181
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   580
         TabIndex        =   180
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   580
         TabIndex        =   179
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   178
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
         TabIndex        =   186
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
      TabIndex        =   167
      Top             =   1896
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   0
         TabIndex        =   175
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   580
         TabIndex        =   174
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   0
         TabIndex        =   173
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   580
         TabIndex        =   172
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   0
         TabIndex        =   171
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   580
         TabIndex        =   170
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   580
         TabIndex        =   169
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   0
         TabIndex        =   168
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
         TabIndex        =   176
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
      TabIndex        =   157
      Top             =   2964
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   0
         TabIndex        =   165
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   580
         TabIndex        =   164
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   0
         TabIndex        =   163
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   580
         TabIndex        =   162
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   0
         TabIndex        =   161
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   580
         TabIndex        =   160
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   580
         TabIndex        =   159
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   18
         Left            =   0
         TabIndex        =   158
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
         TabIndex        =   166
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
      TabIndex        =   147
      Top             =   4032
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   0
         TabIndex        =   155
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   580
         TabIndex        =   154
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   0
         TabIndex        =   153
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   580
         TabIndex        =   152
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   0
         TabIndex        =   151
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   580
         TabIndex        =   150
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   580
         TabIndex        =   149
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   25
         Left            =   0
         TabIndex        =   148
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
         TabIndex        =   156
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
      TabIndex        =   137
      Top             =   5100
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   0
         TabIndex        =   145
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   580
         TabIndex        =   144
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   0
         TabIndex        =   143
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   580
         TabIndex        =   142
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   0
         TabIndex        =   141
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   580
         TabIndex        =   140
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   580
         TabIndex        =   139
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   32
         Left            =   0
         TabIndex        =   138
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
         TabIndex        =   146
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
      TabIndex        =   127
      Top             =   1896
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   0
         TabIndex        =   135
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   580
         TabIndex        =   134
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   0
         TabIndex        =   133
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   580
         TabIndex        =   132
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   0
         TabIndex        =   131
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   580
         TabIndex        =   130
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   580
         TabIndex        =   129
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   12
         Left            =   0
         TabIndex        =   128
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
         TabIndex        =   136
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
      TabIndex        =   117
      Top             =   2964
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   0
         TabIndex        =   125
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   576
         TabIndex        =   124
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   0
         TabIndex        =   123
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   576
         TabIndex        =   122
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   0
         TabIndex        =   121
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   576
         TabIndex        =   120
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   576
         TabIndex        =   119
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   0
         TabIndex        =   118
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
         TabIndex        =   126
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
      TabIndex        =   107
      Top             =   4032
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   0
         TabIndex        =   115
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   580
         TabIndex        =   114
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   0
         TabIndex        =   113
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   580
         TabIndex        =   112
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   0
         TabIndex        =   111
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   580
         TabIndex        =   110
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   580
         TabIndex        =   109
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   26
         Left            =   0
         TabIndex        =   108
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
         TabIndex        =   116
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
      TabIndex        =   97
      Top             =   5100
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   0
         TabIndex        =   105
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   580
         TabIndex        =   104
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   0
         TabIndex        =   103
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   580
         TabIndex        =   102
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   0
         TabIndex        =   101
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   580
         TabIndex        =   100
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   580
         TabIndex        =   99
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   33
         Left            =   0
         TabIndex        =   98
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
         TabIndex        =   106
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraFri 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   6
      Left            =   5736
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   96
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
      Index           =   1
      Left            =   6876
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   86
      Top             =   840
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   94
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   580
         TabIndex        =   93
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   92
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   580
         TabIndex        =   91
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   90
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   580
         TabIndex        =   89
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   580
         TabIndex        =   88
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   87
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
         TabIndex        =   95
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
      TabIndex        =   76
      Top             =   1896
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   0
         TabIndex        =   84
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   580
         TabIndex        =   83
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   0
         TabIndex        =   82
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   580
         TabIndex        =   81
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   0
         TabIndex        =   80
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   580
         TabIndex        =   79
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   580
         TabIndex        =   78
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   0
         TabIndex        =   77
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
         TabIndex        =   85
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
      TabIndex        =   66
      Top             =   2952
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   0
         TabIndex        =   74
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   580
         TabIndex        =   73
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   0
         TabIndex        =   72
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   580
         TabIndex        =   71
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   0
         TabIndex        =   70
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   580
         TabIndex        =   69
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   580
         TabIndex        =   68
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   0
         TabIndex        =   67
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
         TabIndex        =   75
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
      TabIndex        =   56
      Top             =   4032
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   0
         TabIndex        =   64
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   580
         TabIndex        =   63
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   0
         TabIndex        =   62
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   580
         TabIndex        =   61
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   0
         TabIndex        =   60
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   580
         TabIndex        =   59
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   580
         TabIndex        =   58
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   27
         Left            =   0
         TabIndex        =   57
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
         TabIndex        =   65
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
      TabIndex        =   46
      Top             =   5100
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   0
         TabIndex        =   54
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   580
         TabIndex        =   53
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   0
         TabIndex        =   52
         Top             =   345
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   580
         TabIndex        =   51
         Top             =   345
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   0
         TabIndex        =   50
         Top             =   555
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   580
         TabIndex        =   49
         Top             =   555
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   580
         TabIndex        =   48
         Top             =   765
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   34
         Left            =   0
         TabIndex        =   47
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
         TabIndex        =   55
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox fraSat 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   6
      Left            =   6876
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   45
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
   Begin VB.PictureBox fraFri 
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   1
      Left            =   5736
      ScaleHeight     =   1020
      ScaleWidth      =   1095
      TabIndex        =   35
      Top             =   840
      Width           =   1092
      Begin VB.TextBox txtS1s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   43
         Top             =   120
         Width           =   580
      End
      Begin VB.TextBox txtS1t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   580
         TabIndex        =   42
         Top             =   120
         Width           =   460
      End
      Begin VB.TextBox txtS2s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   41
         Top             =   348
         Width           =   580
      End
      Begin VB.TextBox txtS2t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   580
         TabIndex        =   40
         Top             =   348
         Width           =   460
      End
      Begin VB.TextBox txtS3s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   39
         Top             =   552
         Width           =   580
      End
      Begin VB.TextBox txtS3t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   580
         TabIndex        =   38
         Top             =   552
         Width           =   460
      End
      Begin VB.TextBox txtS4t 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   580
         TabIndex        =   37
         Top             =   768
         Width           =   460
      End
      Begin VB.TextBox txtS4s 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   36
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
         TabIndex        =   44
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CheckBox optAll 
      Caption         =   "All Centers"
      Height          =   255
      Left            =   2040
      TabIndex        =   34
      Top             =   7440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPe03b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optWcn 
      Caption         =   "Select Calendar"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   7320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6240
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Press To Save or Update Calendar"
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7160
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   800
   End
   Begin VB.ComboBox cmbYer 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "Year"
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cmbMon 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Month"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   255
      Index           =   9
      Left            =   4560
      TabIndex        =   31
      Top             =   120
      Width           =   90
   End
   Begin VB.Label lblFrom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   4875
      TabIndex        =   30
      ToolTipText     =   "Work Center"
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "For:"
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   28
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblFrom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   27
      ToolTipText     =   "Shop"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saturday"
      Height          =   252
      Index           =   7
      Left            =   6840
      TabIndex        =   26
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Friday"
      Height          =   252
      Index           =   6
      Left            =   5736
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Month/Year"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   18
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "CapaCPe03b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'9/22/03 Added the Calendar options (WC or company)
'6/12/06 Revised ToolTipText
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim ADOParameter2 As ADODB.Parameter
Dim AdoParameter3 As ADODB.Parameter

Dim bGoodCalendar As Byte
Dim iStartDay As Integer
Dim vShifts(8, 9)
Dim sMsg  As String

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
      OpenHelpContext 4201
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSave_Click()
   Dim bResponse As Byte
   Dim iList As Integer
   Dim K As Integer
   Dim dDate As Date
   Dim sCenter As String
   Dim sShop As String
   Dim sMsg As String
   If bGoodCalendar = 0 Then
      MsgBox "The Work Center Calendar Can't Be Saved.", vbInformation, Caption
      Exit Sub
   End If
   If optAll.Value = vbChecked Then
      sMsg = "Do You Want To Update ALL Work Center " _
             & vbCr & "Calendars Based On This Calendar?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         DoAllCenters
         Exit Sub
      End If
   End If
   MouseCursor 11
   cmdSave.Enabled = False
   sShop = Compress(lblFrom(0))
   sCenter = Compress(lblFrom(1))
   'On Error Resume Next
   sSql = "DELETE FROM WcclTable WHERE WCCREF='" & cmbMon & "-" & cmbYer & "' " _
          & "AND WCCCENTER='" & sCenter & "' AND WCCSHOP='" & sShop & "'"
   clsADOCon.ExecuteSQL sSql
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
      dDate = cmbMon & "-" & K & "-" & cmbYer
      sSql = "INSERT INTO WcclTable (WCCREF,WCCCENTER,WCCSHOP,WCCDAY," _
             & "WCCSHH1,WCCSHH2,WCCSHH3,WCCSHH4," _
             & "WCCSHR1,WCCSHR2,WCCSHR3,WCCSHR4,WCCDATE) " _
             & "VALUES('" & cmbMon & "-" & cmbYer & "','" _
             & sCenter & "','" _
             & sShop & "'," _
             & str(K) & "," _
             & txtS1s(iList) & "," _
             & txtS2s(iList) & "," _
             & txtS3s(iList) & "," _
             & txtS4s(iList) & "," _
             & Val(txtS1t(iList)) & "," _
             & Val(txtS2t(iList)) & "," _
             & Val(txtS3t(iList)) & "," _
             & Val(txtS4t(iList)) & ",'" _
             & Format(dDate, "mm/dd/yy") & "')"
      clsADOCon.ExecuteSQL sSql
   Next
   MouseCursor 0
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      SysMsg "The Calendar Was Saved.", True, Me
   Else
      clsADOCon.RollbackTrans
      MsgBox "Couldn't Save The Calendar.", _
         vbInformation, Caption
   End If
   cmdSave.Enabled = True
   Exit Sub
   
CwcalSv1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume CwcalSv2
CwcalSv2:
   MouseCursor 0
   On Error Resume Next
   clsADOCon.RollbackTrans
   MsgBox Trim(str(CurrError.Number)) & vbCr & CurrError.Description, vbExclamation, Caption
   
End Sub


Private Sub Form_Activate()
   MouseCursor 0
   GetThisMonth
   MDISect.lblBotPanel = Caption
   
End Sub

Private Sub Form_Load()
   Dim A As Integer
   Dim iList As Integer
   Dim vMonth As Variant
   FormLoad Me, ES_DONTLIST, ES_DONTRESIZE
   'Don't resize it
   'SetFormSize Me
   Move 0, 0
   On Error Resume Next
   lblFrom(0) = CapaCPe03a.cmbShp
   lblFrom(1) = CapaCPe03a.cmbWcn
   sSql = "SELECT * FROM WcclTable WHERE WCCREF= ? AND WCCCENTER= ? AND WCCSHOP = ?"
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.SIZE = 8
   AdoParameter1.Type = adChar
   
   Set ADOParameter2 = New ADODB.Parameter
   ADOParameter2.SIZE = 12
   ADOParameter2.Type = adChar
   
   Set AdoParameter3 = New ADODB.Parameter
   AdoParameter3.SIZE = 12
   AdoParameter3.Type = adChar
  
   
   AdoQry.Parameters.Append AdoParameter1
   AdoQry.Parameters.Append ADOParameter2
   AdoQry.Parameters.Append AdoParameter3
   
   
   cmbMon.Clear
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
   iList = Format(Now, "m")
   cmbMon = cmbMon.List(iList - 1)
   A = Format(Now, "yyyy")
   For iList = A - 2 To A + 25
      AddComboStr cmbYer.hwnd, Format$(iList)
   Next
   cmbYer = Format(ES_SYSDATE, "yyyy")
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
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Hide
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set AdoParameter1 = Nothing
   Set ADOParameter2 = Nothing
   Set AdoParameter3 = Nothing
   Set AdoQry = Nothing
   CapaCPe03a.Show
   Set CapaCPe03b = Nothing
   
End Sub




Private Sub optAll_Click()
   'never visible-flag for All Centers
   
End Sub

Private Sub optWcn_Click()
   'never visible - times from Work Center (1), Company Calendar (0)
   
End Sub

Private Sub txtS1s_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtS1s_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtS1s_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub


Private Sub txtS1s_LostFocus(Index As Integer)
   txtS1s(Index) = CheckLen(txtS1s(Index), 4)
   txtS1s(Index) = Format(Abs(Val(txtS1s(Index))), "#0.0")
   If Val(txtS1s(Index)) = 0 Then txtS1t(Index) = "0.0"
   
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
   KeyValue KeyAscii
   
End Sub


Private Sub txtS2s_LostFocus(Index As Integer)
   txtS2s(Index) = CheckLen(txtS2s(Index), 4)
   txtS2s(Index) = Format(Abs(Val(txtS2s(Index))), "#0.0")
   If Val(txtS2s(Index)) = 0 Then txtS2t(Index) = "0.0"
   
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
   KeyValue KeyAscii
   
End Sub

Private Sub txtS3s_LostFocus(Index As Integer)
   txtS3s(Index) = CheckLen(txtS3s(Index), 4)
   txtS3s(Index) = Format(Abs(Val(txtS3s(Index))), "#0.0")
   If Val(txtS3s(Index)) = 0 Then txtS3t(Index) = "0.0"
   
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
   KeyValue KeyAscii
   
End Sub

Private Sub txtS4s_LostFocus(Index As Integer)
   txtS4s(Index) = CheckLen(txtS4s(Index), 4)
   txtS4s(Index) = Format(Abs(Val(txtS4s(Index))), "#0.0")
   If Val(txtS4s(Index)) = 0 Then txtS4t(Index) = "0.0"
   
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
   
   On Error GoTo DiaErr1
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
   bGoodCalendar = GetTheCalendar()
   Exit Sub
   
DiaErr1:
   sProcName = "getthismo"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Function GetTheCalendar() As Byte
   Dim RdoCal As ADODB.Recordset
   Dim A As Integer
   Dim iList As Integer
   Dim sCalendar As String
   Dim sCenter As String
   Dim sShop As String
   Dim sMsg As String
   Dim bResponse  As Byte
   'close some boxes to avoid recursion
   'cmbMon.Enabled = False
   cmbYer.Enabled = False
   cmdSave.Enabled = False
   
   GetTheCalendar = 0
   sShop = Compress(lblFrom(0))
   sCenter = Compress(lblFrom(1))
   sCalendar = cmbMon & "-" & cmbYer
   On Error GoTo CwcalGc1
   AdoQry.Parameters(0).Value = sCalendar
   AdoQry.Parameters(1).Value = sCenter
   AdoQry.Parameters(2).Value = sShop
   bSqlRows = clsADOCon.GetQuerySet(RdoCal, AdoQry, ES_KEYSET, False, 1)
   
   clsADOCon.ADOErrNum = 0

   If bSqlRows Then
      sMsg = "This Monthly Calendar Already Exists." _
      & vbCr & "Do You Want To Update With Weekly Calendar?" _
      & vbCr & "(Need To Apply The Update)"
      
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   
      If bResponse = vbNo Then
         
         For iList = 0 To 6
            If lblDte(iList).Visible Then Exit For
         Next
         With RdoCal
            Do Until .EOF
               txtS1s(iList) = Format$(0 + !WCCSHH1, "#0.0")
               txtS2s(iList) = Format$(0 + !WCCSHH2, "#0.0")
               txtS3s(iList) = Format$(0 + !WCCSHH3, "#0.0")
               txtS4s(iList) = Format$(0 + !WCCSHH4, "#0.0")
               txtS1t(iList) = Format$(0 + !WCCSHR1, "#0.0")
               txtS2t(iList) = Format$(0 + !WCCSHR2, "#0.0")
               txtS3t(iList) = Format$(0 + !WCCSHR3, "#0.0")
               txtS4t(iList) = Format$(0 + !WCCSHR4, "#0.0")
               iList = iList + 1
               .MoveNext
            Loop
            ClearResultSet RdoCal
         End With
         GetTheCalendar = 1
      Else
         If optWcn.Value = vbUnchecked Then
            FillFromCompCalender (sCalendar)
         Else
            ' Fill From Work Center
            FillFromWCCalender sCenter, sShop
         End If
      End If
   Else
      If optWcn.Value = vbUnchecked Then
         FillFromCompCalender (sCalendar)
      Else
         ' Fill From Work Center
         FillFromWCCalender sCenter, sShop
      End If
   End If
   
   If clsADOCon.ADOErrNum = 0 Then GetTheCalendar = 1
   
   On Error Resume Next
   Set RdoCal = Nothing
   cmbYer.Enabled = True
   cmdSave.Enabled = True
   MouseCursor 0
   Exit Function
   
CwcalGc1:
   sProcName = "getthecalendar"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume CwcalGc2
CwcalGc2:
   On Error Resume Next
   'cmbMon.Enabled = True
   cmbYer.Enabled = True
   cmdSave.Enabled = False
   GetTheCalendar = 0
   DoModuleErrors Me
   
End Function

Private Function FillFromCompCalender(sCalen As String)
   Dim RdoCal As ADODB.Recordset
   Dim iList As Integer
   Dim A As Integer

   'Company Calendar
   sSql = "SELECT * FROM CoclTable WHERE COCREF='" & sCalen & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCal, ES_FORWARD)
   If bSqlRows Then
      For iList = 0 To 6
         If lblDte(iList).Visible Then Exit For
      Next
      With RdoCal
         Do Until .EOF
            txtS1s(iList) = Format(!COCSHT1, "#0.0")
            txtS2s(iList) = Format(!COCSHT2, "#0.0")
            txtS3s(iList) = Format(!COCSHT3, "#0.0")
            txtS4s(iList) = Format(!COCSHT4, "#0.0")
            txtS1t(iList) = "0.0"
            txtS2t(iList) = "0.0"
            txtS3t(iList) = "0.0"
            txtS4t(iList) = "0.0"
            iList = iList + 1
            .MoveNext
         Loop
         ClearResultSet RdoCal
      End With
   End If

   Set RdoCal = Nothing
   
End Function

Private Function FillFromWCCalender(sCenter As String, sShop As String)
   Dim RdoCal As ADODB.Recordset
   Dim iList As Integer
   Dim A As Integer
   'Work Center

   sSql = "SELECT * FROM WcntTable WHERE WCNREF='" & sCenter _
          & "' AND WCNSHOP='" & sShop & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCal, ES_FORWARD)
   
   If bSqlRows Then
      With RdoCal
         Do Until .EOF
            For iList = 0 To 6
               Select Case iList
                  Case 0
                     For A = 0 To 35 Step 7
                        txtS1s(A) = Format$(0 + !WCNSUNHR1, "#0.0")
                        txtS1t(A) = Format$(0 + !WCNSUNMU1, "#0.0")
                        txtS2s(A) = Format$(0 + !WCNSUNHR2, "#0.0")
                        txtS2t(A) = Format$(0 + !WCNSUNMU2, "#0.0")
                        txtS3s(A) = Format$(0 + !WCNSUNHR3, "#0.0")
                        txtS3t(A) = Format$(0 + !WCNSUNMU3, "#0.0")
                        txtS4s(A) = Format$(0 + !WCNSUNHR4, "#0.0")
                        txtS4t(A) = Format$(0 + !WCNSUNMU4, "#0.0")
                     Next
                  Case 1
                     For A = 1 To 36 Step 7
                        txtS1s(A) = Format$(0 + !WCNMONHR1, "#0.0")
                        txtS1t(A) = Format$(0 + !WCNMONMU1, "#0.0")
                        txtS2s(A) = Format$(0 + !WCNMONHR2, "#0.0")
                        txtS2t(A) = Format$(0 + !WCNMONMU2, "#0.0")
                        txtS3s(A) = Format$(0 + !WCNMONHR3, "#0.0")
                        txtS3t(A) = Format$(0 + !WCNMONMU3, "#0.0")
                        txtS4s(A) = Format$(0 + !WCNMONHR4, "#0.0")
                        txtS4t(A) = Format$(0 + !WCNMONMU4, "#0.0")
                     Next
                  Case 2
                     For A = 2 To 30 Step 7
                        txtS1s(A) = Format$(0 + !WCNTUEHR1, "#0.0")
                        txtS1t(A) = Format$(0 + !WCNTUEMU1, "#0.0")
                        txtS2s(A) = Format$(0 + !WCNTUEHR2, "#0.0")
                        txtS2t(A) = Format$(0 + !WCNTUEMU2, "#0.0")
                        txtS3s(A) = Format$(0 + !WCNTUEHR3, "#0.0")
                        txtS3t(A) = Format$(0 + !WCNTUEMU3, "#0.0")
                        txtS4s(A) = Format$(0 + !WCNTUEHR4, "#0.0")
                        txtS4t(A) = Format$(0 + !WCNTUEMU4, "#0.0")
                     Next
                  Case 3
                     For A = 3 To 31 Step 7
                        txtS1s(A) = Format$(0 + !WCNWEDHR1, "#0.0")
                        txtS1t(A) = Format$(0 + !WCNWEDMU1, "#0.0")
                        txtS2s(A) = Format$(0 + !WCNWEDHR2, "#0.0")
                        txtS2t(A) = Format$(0 + !WCNWEDMU2, "#0.0")
                        txtS3s(A) = Format$(0 + !WCNWEDHR3, "#0.0")
                        txtS3t(A) = Format$(0 + !WCNWEDMU3, "#0.0")
                        txtS4s(A) = Format$(0 + !WCNWEDHR4, "#0.0")
                        txtS4t(A) = Format$(0 + !WCNWEDMU4, "#0.0")
                     Next
                  Case 4
                     For A = 4 To 32 Step 7
                        txtS1s(A) = Format$(0 + !WCNTHUHR1, "#0.0")
                        txtS1t(A) = Format$(0 + !WCNTHUMU1, "#0.0")
                        txtS2s(A) = Format$(0 + !WCNTHUHR2, "#0.0")
                        txtS2t(A) = Format$(0 + !WCNTHUMU2, "#0.0")
                        txtS3s(A) = Format$(0 + !WCNTHUHR3, "#0.0")
                        txtS3t(A) = Format$(0 + !WCNTHUMU3, "#0.0")
                        txtS4s(A) = Format$(0 + !WCNTHUHR4, "#0.0")
                        txtS4t(A) = Format$(0 + !WCNTHUMU4, "#0.0")
                     Next
                  Case 5
                     For A = 5 To 33 Step 7
                        txtS1s(A) = Format$(0 + !WCNFRIHR1, "#0.0")
                        txtS1t(A) = Format$(0 + !WCNFRIMU1, "#0.0")
                        txtS2s(A) = Format$(0 + !WCNFRIHR2, "#0.0")
                        txtS2t(A) = Format$(0 + !WCNFRIMU2, "#0.0")
                        txtS3s(A) = Format$(0 + !WCNFRIHR3, "#0.0")
                        txtS3t(A) = Format$(0 + !WCNFRIMU3, "#0.0")
                        txtS4s(A) = Format$(0 + !WCNFRIHR4, "#0.0")
                        txtS4t(A) = Format$(0 + !WCNFRIMU4, "#0.0")
                     Next
                  Case 6
                     For A = 6 To 34 Step 7
                        txtS1s(A) = Format$(0 + !WCNSATHR1, "#0.0")
                        txtS1t(A) = Format$(0 + !WCNSATMU1, "#0.0")
                        txtS2s(A) = Format$(0 + !WCNSATHR2, "#0.0")
                        txtS2t(A) = Format$(0 + !WCNSATMU2, "#0.0")
                        txtS3s(A) = Format$(0 + !WCNSATHR3, "#0.0")
                        txtS3t(A) = Format$(0 + !WCNSATMU3, "#0.0")
                        txtS4s(A) = Format$(0 + !WCNSATHR4, "#0.0")
                        txtS4t(A) = Format$(0 + !WCNSATMU4, "#0.0")
                     Next
               End Select
            Next
            .MoveNext
         Loop
         ClearResultSet RdoCal
      End With
   End If
   
   Set RdoCal = Nothing

End Function
Private Sub DoAllCenters()
   Dim RdoCnt As ADODB.Recordset
   
   Dim A As Integer
   Dim b As Integer
   Dim C As Integer
   Dim iList As Integer
   Dim K As Integer
   Dim iTotalCenters As Integer
   Dim dDate As String
   Dim sCenter As String
   Dim sShop As String
   Dim sCenters(500, 2) As String
   cmdSave.Enabled = False
   MouseCursor 11
   sSql = "SELECT WCNREF,WCNSHOP FROM WcntTable ORDER BY WCNREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCnt, ES_FORWARD)
   If bSqlRows Then
      With RdoCnt
         Do Until .EOF
            iList = iList + 1
            sCenters(iList, 0) = "" & Trim(!WCNREF)
            sCenters(iList, 1) = "" & Trim(!WCNSHOP)
            .MoveNext
         Loop
         ClearResultSet RdoCnt
      End With
   End If
   iTotalCenters = iList
   For iList = 0 To 6
      If lblDte(iList).Visible Then Exit For
   Next
   iStartDay = iList
   If iTotalCenters > 0 Then b = 100 / iTotalCenters
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   SysOpen.pnl = "Updating Calendars"
   SysOpen.Move CapaCPe03b.Left + (CapaCPe03b.Width / 2), CapaCPe03b.Top + (CapaCPe03b.Height / 2 - 1000)
   SysOpen.Show
   For A = 1 To iTotalCenters
      C = C + b
      If C > 95 Then C = 95
      SysOpen.prg1.Value = C
      SysOpen.Refresh
      K = 0
      sCenter = sCenters(A, 0)
      sShop = sCenters(A, 1)
      sSql = "DELETE FROM WcclTable WHERE WCCREF='" & cmbMon & "-" & cmbYer & "' " _
             & "AND WCCCENTER='" & sCenter & "' AND WCCSHOP='" & sShop & "'"
      clsADOCon.ExecuteSQL sSql
      For iList = iStartDay To 36
         If Not lblDte(iList).Visible Then Exit For
         K = K + 1
         dDate = cmbMon & "-" & K & "-" & cmbYer
         sSql = "INSERT INTO WcclTable (WCCREF,WCCCENTER,WCCSHOP,WCCDAY," _
                & "WCCSHH1,WCCSHH2,WCCSHH3,WCCSHH4," _
                & "WCCSHR1,WCCSHR2,WCCSHR3,WCCSHR4,WCCDATE) " _
                & "VALUES('" & cmbMon & "-" & cmbYer & "','" _
                & sCenter & "','" _
                & sShop & "'," _
                & str(K) & "," _
                & txtS1s(iList) & "," _
                & txtS2s(iList) & "," _
                & txtS3s(iList) & "," _
                & txtS4s(iList) & "," _
                & Val(txtS1t(iList)) & "," _
                & Val(txtS2t(iList)) & "," _
                & Val(txtS3t(iList)) & "," _
                & Val(txtS4t(iList)) & ",'" _
                & Format(dDate, "mm/dd/yy") & "')"
         clsADOCon.ExecuteSQL sSql
      Next
   Next
   MouseCursor 0
   SysOpen.prg1.Value = 100
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      SysMsg "The Calendars Were Saved.", True, Me
   Else
      clsADOCon.RollbackTrans
      MsgBox "Couldn't Save The Calendars.", _
         vbInformation, Caption
   End If
   Unload SysOpen
   cmdSave.Enabled = True
   Set RdoCnt = Nothing
End Sub
