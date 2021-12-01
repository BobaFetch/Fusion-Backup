VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDISect 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox SSPanel1 
      Align           =   1  'Align Top
      BackColor       =   &H00D8E9EC&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   4620
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.PictureBox SSPanel2 
         BackColor       =   &H00D8E9EC&
         Height          =   975
         Left            =   0
         ScaleHeight     =   915
         ScaleWidth      =   4620
         TabIndex        =   2
         Top             =   0
         Width           =   4680
         Begin VB.Label lblBotPanel 
            Height          =   336
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   2532
         End
         Begin VB.Label OvrPanel 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "OVER"
            Height          =   324
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   780
         End
         Begin VB.Label SystemMsg 
            Caption         =   "SystemMsg"
            Height          =   315
            Left            =   0
            TabIndex        =   4
            Top             =   660
            Width           =   3615
         End
         Begin VB.Label Label2 
            Caption         =   "Dummy form required by OpenSqlServer"
            Height          =   255
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   3735
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Dummy form required by OpenSqlServer"
         Height          =   435
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3735
      End
   End
   Begin Crystal.CrystalReport Crw 
      Left            =   2160
      Top             =   1860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   110
      WindowTop       =   35
      WindowWidth     =   460
      WindowHeight    =   410
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   7
      DiscardSavedData=   -1  'True
      WindowState     =   1
      PrintFileLinesPerPage=   60
      WindowShowProgressCtls=   0   'False
   End
End
Attribute VB_Name = "MDISect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
