VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.MDIForm MDISect 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   Enabled         =   0   'False
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin MSComDlg.CommonDialog Cdi 
      Left            =   1260
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "IIF"
      DialogTitle     =   "Export AR Activity To ."
      Filter          =   "IIF"
      FontSize        =   4.38642e-38
   End
   Begin VB.PictureBox SSPanel1 
      Align           =   1  'Align Top
      BackColor       =   &H00D8E9EC&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   7140
      TabIndex        =   0
      Top             =   0
      Width           =   7200
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
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   360
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "MDISectDummy.frx":0000
   End
End
Attribute VB_Name = "MDISect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bUnloading As Byte
Dim bOnLoad As Byte

Private Sub MDIForm_Load()
    MouseCursor 13
'    Dim sYear As String
'
'    bResize = GetSetting("Esi2000", "System", "ResizeForm", bResize)
'    SaveSetting "Esi2000", "AppTitle", "fina", "ESI Finance"
'
'    MouseCursor 13
'    sYear = Format$(Now, "yyyy")
'    GetRecentList "EsiFina"
'    '11/23/04
    On Error Resume Next
    If bUnloading = 0 Then
       bOnLoad = 1
        ' MM 9/5/2009
        ' Open the database connection
        If Not OpenDBServer(False) Then
            End
        End If
        ' Check the security
        'CheckSectionPermissions
        ' Show the MDI form
       'Show
    End If
   
End Sub
