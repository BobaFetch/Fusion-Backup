VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form AdmnCertUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Authorization Access Manager"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraUser 
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   6135
      Begin VB.PictureBox ImgPic 
         Height          =   945
         Left            =   240
         ScaleHeight     =   885
         ScaleWidth      =   4935
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Depending On Size, The Image May Appear Cropped. See The Report For Actual Image Representation"
         Top             =   1440
         Width           =   4995
      End
      Begin VB.TextBox txtNme 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Tag             =   "2"
         Top             =   180
         Width           =   3855
      End
      Begin VB.TextBox txtNik 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Tag             =   "2"
         Top             =   540
         Width           =   1935
      End
      Begin VB.TextBox txtInt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         TabIndex        =   7
         Tag             =   "3"
         Top             =   540
         Width           =   975
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nickname"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Initials"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   11
         Top             =   540
         Width           =   495
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Signature :"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnCertUser.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdChg 
      Caption         =   "C&hange"
      Height          =   300
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Change The Current User Password"
      Top             =   960
      Width           =   840
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   300
      Left            =   5280
      TabIndex        =   1
      ToolTipText     =   "Add A New User"
      Top             =   600
      Width           =   840
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double Click To Select A User (Or Select And Press Enter)"
      Top             =   480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4048
      _Version        =   393216
      Rows            =   20
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorSel    =   -2147483636
      AllowBigSelection=   0   'False
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6600
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   840
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   5280
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5685
      FormDesignWidth =   7515
   End
   Begin VB.Label lblUsers 
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "AdmnCertUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'9/9/05 Corrected opening and closing files
Dim bOnLoad As Byte
Dim iOldrec As Integer

'grid columns
Private Const COL_Name = 0
Private Const COL_Rpt = 1
Private Const COL_Active = 2
Private Const COL_RecNo = 3
Private Const COL_Count = 4

Dim Chunk() As Byte
Dim Chunks As Integer
Dim Fragment As Integer
Dim i As Integer
Dim DataFile As Integer
Dim Fl As Long
Dim sPicFile As String
Const ChunkSize As Integer = 16384


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub chkShowInactive_Click()
   FormatGrid
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdChg_Click()
   
   Dim strUserName As String
   Dim iCurrentRec As Integer
   Dim strRptName As String
   
   
   If Grid1.row = 0 Then
      If Grid1.Rows > 1 Then
         Grid1.row = 1
      Else
         Exit Sub
      End If
   End If
   
   
   iCurrentRec = Grid1.TextMatrix(Grid1.row, COL_RecNo)
   strUserName = Grid1.TextMatrix(Grid1.row, COL_Name)
   strRptName = Grid1.TextMatrix(Grid1.row, COL_Rpt)
   
   
   AdmnCerUpdUsr.txtUserName = strUserName
   AdmnCerUpdUsr.cmbRpt = strRptName
   AdmnCerUpdUsr.Show
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      OpenHelpContext 30
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub cmdNew_Click()
   AdmnCerNUser.Show
   'FormatGrid
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      'iOldrec = iUserIdx
      FormatGrid
      bOnLoad = 0
      
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   'iUserIdx = iOldrec
   On Error Resume Next
   FormUnload
   Set AdmnCertUser = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Public Sub FormatGrid()
   Dim row As Integer
   Dim iList As Integer
   'Dim RdoRpt As rdoResultset
   Dim RdoRpt As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   Grid1.Clear
   Grid1.Cols = COL_Count
   Grid1.ColWidth(COL_Name) = 1900
   Grid1.ColWidth(COL_Rpt) = 1550
   Grid1.ColWidth(COL_Active) = 900
   Grid1.ColWidth(COL_RecNo) = 0
   Grid1.row = 0
   Grid1.TextMatrix(0, COL_Name) = "User Id"
   Grid1.TextMatrix(0, COL_Rpt) = "Report Name"
   Grid1.TextMatrix(0, COL_Active) = "Active"
   Grid1.TextMatrix(0, COL_RecNo) = "Row"
   

   Grid1.Rows = 1
   
   sSql = "select USRNAME, ReportName , ISNULL(ACTIVE, 0) ACTIVE" _
         & " From dbo.CertUsrSec, CertReports" _
            & " Where CERTRPTID = REPORTID"
  ' bSqlRows = GetDataSet(RdoRpt, ES_FORWARD)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   
   If bSqlRows Then
      
      With RdoRpt
         Do Until .EOF
            row = row + 1
            If row >= Grid1.Rows Then Grid1.Rows = Grid1.Rows + 1
            Grid1.TextMatrix(row, COL_Name) = !USRNAME
            Grid1.TextMatrix(row, COL_Rpt) = !ReportName
            Grid1.TextMatrix(row, COL_Active) = IIf(!Active = 1, "Active", "Inactive")
            Grid1.TextMatrix(row, COL_RecNo) = iList
            
            .MoveNext
         Loop
         ClearResultSet RdoRpt
      End With
      
   End If
   Set RdoRpt = Nothing
   MouseCursor 0
   Exit Sub
   
   'lblUsers = row
   If Grid1.Rows > 1 Then
      Grid1.Sort = 1
      Grid1.Col = 0
      Grid1.row = 1
      Grid1_DblClick
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "formatgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Grid1_Click()
   GetSelectedUser
End Sub

Private Sub Grid1_DblClick()
   GetSelectedUser
   
End Sub


Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      
      GetSelectedUser
   End If
   
End Sub

Private Sub ResetControls(bOpen As Boolean)
   Dim iList As Integer
   txtNme.Enabled = bOpen
   txtNik.Enabled = bOpen
   txtInt.Enabled = bOpen
End Sub


Private Sub GetSelectedUser()

   Dim strUserName As String
   Dim iCurrentRec As Integer
   Dim strRptName As String
   
   If Grid1.row = 0 Then
      If Grid1.Rows > 1 Then
         Grid1.row = 1
      Else
         Exit Sub
      End If
   End If
   
   
   iCurrentRec = Grid1.TextMatrix(Grid1.row, COL_RecNo)
   strUserName = Grid1.TextMatrix(Grid1.row, COL_Name)
   strRptName = Grid1.TextMatrix(Grid1.row, COL_Rpt)
   
   GetUserInfo strUserName, strRptName
   
   ResetControls True
   
End Sub

Private Function GetUserInfo(strUserName As String, strRptName As String)

   Dim iRptID As Integer
   'Dim RdoRpt As rdoResultset
   Dim RdoRpt As ADODB.Recordset
   On Error GoTo DiaErr1
   
   iRptID = GetReportID(strRptName)
   
   If (iRptID <> 0) Then
      
      sSql = "SELECT USRFULLNAME, USRNICKNAME, USRINITIAL,USRSIG FROM CertUsrSec " _
            & " WHERE USRNAME = '" & strUserName & "' AND CERTRPTID = " & CStr(iRptID)
      
'      bSqlRows = GetDataSet(RdoRpt, ES_FORWARD)
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)

      If bSqlRows Then
         
         With RdoRpt
            txtNme = Trim(!USRFULLNAME)
            txtNik = Trim(!USRNICKNAME)
            txtInt = Trim(!USRINITIAL)
         End With
      
         ClearResultSet RdoRpt
         
         ' Set the Picture
         GetPicture strUserName, iRptID
      End If
      Set RdoRpt = Nothing
      
   End If

   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "FillReports"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
      
End Function


Private Sub GetPicture(strUserName As String, iRptID As Integer)
   
   Dim DiskFile As String
   'Dim RdoPic As rdoResultset
   Dim RdoPic As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "SELECT USRSIG FROM CertUsrSec WHERE " _
            & " USRNAME = '" & strUserName & "' AND CERTRPTID = " & CStr(iRptID)
   
  ' bSqlRows = GetDataSet(RdoPic, ES_STATIC)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPic, ES_STATIC)
   If Not bSqlRows Then
      MsgBox "There Is No Signature For This User and Report.", _
         vbInformation, Caption
   Else
      If IsNull(RdoPic!USRSIG) Then
         MsgBox "Signature Has Not Been Assigned.", _
            vbInformation, Caption
      Else
         DiskFile = "c:\Program Files\ES2000\Temp\picture1.jpg"
         If Len(Dir$(DiskFile)) > 0 Then Kill DiskFile
         
         DataFile = FreeFile
         Open DiskFile For Binary Access Write As DataFile
         'Fl = RdoPic!USRSIG.ColumnSize
         Fl = RdoPic.Fields("USRSIG").ActualSize
         Chunks = Fl \ ChunkSize
         Fragment = Fl Mod ChunkSize
         ReDim Chunk(Fragment)
         'Chunk() = RdoPic!USRSIG.GetChunk(Fragment)
         Chunk() = RdoPic.Fields("USRSIG").GetChunk(Fragment)
         Put DataFile, , Chunk()
         For i = 1 To Chunks
            ReDim Buffer(ChunkSize)
            'Chunk() = RdoPic!USRSIG.GetChunk(ChunkSize)
            Chunk() = RdoPic.Fields("USRSIG").GetChunk(ChunkSize)
            Put DataFile, , Chunk()
         Next
         Close DataFile
         ImgPic.Picture = LoadPicture(DiskFile)
      End If
   End If
   RdoPic.Close
   
   Exit Sub
   
DiaErr1:
   sProcName = "getpicture"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


