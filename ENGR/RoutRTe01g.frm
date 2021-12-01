VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTe01g 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routing Pictures"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRte 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   19
      Tag             =   "8"
      Text            =   " "
      ToolTipText     =   "(30) Char Maximun"
      Top             =   720
      Width           =   3075
   End
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   810
      Picture         =   "RoutRTe01g.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Print The Report"
      Top             =   7440
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   240
      Picture         =   "RoutRTe01g.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Display The Report"
      Top             =   7440
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Test"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   16
      ToolTipText     =   "Test The Report (Pictures By Routing Operation)"
      Top             =   7440
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox Image1 
      Height          =   4572
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   7995
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Depending On Size, The Image May Appear Cropped. See The Report For Actual Image Representation"
      Top             =   2280
      Width           =   8052
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTe01g.frx":0308
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "F&ile"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5640
      TabIndex        =   13
      ToolTipText     =   "Select A File To Be Saved"
      Top             =   1080
      Width           =   915
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   315
      Left            =   5640
      TabIndex        =   12
      ToolTipText     =   "Load A Current Picture"
      Top             =   720
      Width           =   915
   End
   Begin VB.TextBox txtPic 
      Height          =   285
      Left            =   1680
      MaxLength       =   80
      TabIndex        =   3
      Tag             =   "8"
      Text            =   " "
      ToolTipText     =   "(80) Char Maximun"
      Top             =   1920
      Width           =   6492
   End
   Begin VB.TextBox txtCmt 
      Height          =   948
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "8"
      Text            =   "RoutRTe01g.frx":0AB6
      ToolTipText     =   "Operation Comments"
      Top             =   7680
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
      Tag             =   "8"
      Text            =   " "
      ToolTipText     =   "(30) Char Maximun"
      Top             =   1080
      Width           =   3075
   End
   Begin VB.ComboBox cmbOpno 
      Height          =   315
      Left            =   7200
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Operation Number"
      Top             =   8160
      Visible         =   0   'False
      WhatsThisHelpID =   100
      Width           =   948
   End
   Begin VB.ComboBox cmbRte 
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Routing"
      Top             =   8160
      Visible         =   0   'False
      WhatsThisHelpID =   100
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   7560
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8040
      Top             =   7560
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7080
      FormDesignWidth =   9255
   End
   Begin VB.Label lblFile 
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   6240
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture Description"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1755
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation:"
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   9
      Tag             =   "4"
      Top             =   7800
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   288
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   912
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation"
      Height          =   375
      Index           =   1
      Left            =   7200
      TabIndex        =   6
      Top             =   8400
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Routing "
      Height          =   372
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1272
   End
End
Attribute VB_Name = "RoutRTe01g"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'7/20/06 New
Option Explicit
Dim RdoPic As ADODB.Recordset
Dim bOnLoad As Byte
Dim bGoodOp As Byte
Dim bGoodRout As Byte
Dim bGoodPic As Byte
Dim bNoHelp As Byte

'Chunks (Pictures)
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

Private Sub GetPicture()
   Dim DiskFile As String
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM RtpcTable WHERE (OPREF='" & Compress(txtRte) _
          & "' AND OPNO=0)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPic, ES_STATIC)
   If Not bSqlRows Then
      MsgBox "There Is No Picture Row For This Item.", _
         vbInformation, Caption
   Else
      If IsNull(RdoPic!OPPICTURE) Then
         MsgBox "A Picture Has Not Been Assigned.", _
            vbInformation, Caption
      Else
         
         If (RdoPic.Fields("OPPICTURE").ActualSize = 1) Then
            MsgBox "There was an error in capturing the picture. " & vbCrLf _
               & "Please Select A Picture From A File.", _
               vbInformation, Caption
            lblFile.Visible = True
            
         Else
            DiskFile = "c:\Program Files\ES2000\Temp\picture1.jpg"
            If Len(Dir$(DiskFile)) > 0 Then Kill DiskFile
            
            DataFile = FreeFile
            Open DiskFile For Binary Access Write As DataFile
            Fl = RdoPic.Fields("OPPICTURE").ActualSize
            Chunks = Fl \ ChunkSize
            Fragment = Fl Mod ChunkSize
            ReDim Chunk(Fragment)
            Chunk() = RdoPic.Fields("OPPICTURE").GetChunk(Fragment)
            Put DataFile, , Chunk()
            For i = 1 To Chunks
               ReDim Buffer(ChunkSize)
               Chunk() = RdoPic.Fields("OPPICTURE").GetChunk(ChunkSize)
               Put DataFile, , Chunk()
            Next
            Close DataFile
            Image1.Picture = LoadPicture(DiskFile)
            cmdReport.Enabled = True
            lblFile.Visible = False

         End If
         
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


Private Sub SavePicture()
   DataFile = FreeFile
   sPicFile = lblFile
   On Error GoTo DiaErr1
   Open sPicFile For Binary Access Read As DataFile
   Fl = LOF(DataFile)
   
   sSql = "SELECT * FROM RtpcTable WHERE (OPREF='" & Compress(txtRte) _
          & "' AND OPNO=0)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPic, ES_KEYSET)
   
   'RdoPic.Edit
   Chunks = Fl \ ChunkSize
   Fragment = Fl Mod ChunkSize
   RdoPic!OPPICTURE.AppendChunk Null
   ReDim Chunk(Fragment)
   Get DataFile, , Chunk()
   
   RdoPic!OPPICTURE.AppendChunk Chunk()
   ReDim Chunk(ChunkSize)
   For i = 1 To Chunks
      Get DataFile, , Chunk()
      RdoPic!OPPICTURE.AppendChunk Chunk()
   Next
   Close DataFile
   RdoPic.Update
   On Error Resume Next
   RdoPic.Close
   Exit Sub
   
DiaErr1:
   sProcName = "savepicture"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdFile_Click()
   With MDISect.Cdi
      .Filter = "Pictures(*.jpg;*.bmp)|*.jpg;*.bmp"
      .ShowOpen
   End With
   If Trim(MDISect.Cdi.FileName) <> "" Then
      On Error GoTo DiaErr1
      lblFile.Visible = True
      lblFile = MDISect.Cdi.FileName
      SavePicture
      GetPicture
   End If
   
   Exit Sub
DiaErr1:
   MsgBox "Couldn't Load Picture. Not A Valid File.", _
      vbInformation, Caption
   
End Sub


Private Sub cmdLoad_Click()
   SelectPicture
   'cmdFile.Enabled = True
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      txtPic = ""
      txtPic.Enabled = True
      cmdFile.Enabled = True
      SelectPicture
      'Image1.Picture = LoadPicture("")
      'lblFile = ""
      bGoodPic = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim DiskFile As String
   DiskFile = "c:\Program Files\ES2000\Temp\picture1.jpg"
   If Len(Dir$(DiskFile)) > 0 Then Kill DiskFile
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set RdoPic = Nothing
   Set RoutRTe01g = Nothing
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDsc.BackColor = BackColor
   txtCmt.BackColor = BackColor
   
End Sub

Private Sub txtDsc_Change()
   If Left(txtDsc, 6) = "*** Ro" Then _
           txtDsc.ForeColor = ES_RED Else txtDsc.ForeColor = vbBlack
   
End Sub




Private Function SelectPicture() As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM RtpcTable WHERE (OPREF='" & Compress(txtRte) _
          & "' AND OPNO=0)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPic, ES_STATIC)
   If Not bSqlRows Then
      sSql = "INSERT INTO RtpcTable (OPREF,OPNO) " _
             & "VALUES('" & Compress(txtRte) & "',0)"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      SelectPicture = 0
   Else
      txtPic = "" & RdoPic!OPDESC
      If IsNull(RdoPic!OPPICTURE) Then
         SelectPicture = 0
      Else
         SelectPicture = 1
      End If
   End If
   RdoPic.Close
   If SelectPicture = 0 Then
      MsgBox "A Picture Has Not Been Captured In The Table. " & vbCrLf _
         & "Please Select A Picture From A File.", _
         vbInformation, Caption
   Else
      GetPicture
   End If
   On Error Resume Next
   RdoPic.Close
   Exit Function
   
DiaErr1:
   sProcName = "selectpicture"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtPic_LostFocus()
   txtPic = CheckLen(txtPic, 80)
   txtPic = StrCase(txtPic)
   On Error Resume Next
   sSql = "UPDATE RtpcTable SET OPDESC='" & txtPic & "' WHERE " _
          & "(OPREF='" & Compress(txtRte) & "' AND OPNO=0)"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub
