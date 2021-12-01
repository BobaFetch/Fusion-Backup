VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form AdmnCerUpdUsr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update User Signature Authentication"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   420
      Left            =   7320
      TabIndex        =   20
      ToolTipText     =   "Add A New User"
      Top             =   720
      Width           =   840
   End
   Begin VB.Frame fraUser 
      Height          =   4935
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   6975
      Begin VB.TextBox txtUserName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   0
         Tag             =   "2"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtvPas 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         TabIndex        =   5
         ToolTipText     =   "Case Sensitive Max (15) Char"
         Top             =   2040
         Width           =   1965
      End
      Begin VB.ComboBox cmbRole 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2160
         TabIndex        =   7
         Tag             =   "8"
         ToolTipText     =   "Select User Class From List"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox txtPas 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         TabIndex        =   4
         ToolTipText     =   "Case Sensitive Max (15) Char"
         Top             =   1560
         Width           =   1965
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load Sig"
         Height          =   420
         Left            =   5520
         TabIndex        =   9
         ToolTipText     =   "Add A New User"
         Top             =   3720
         Width           =   1200
      End
      Begin VB.ComboBox cmbRpt 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Tag             =   "8"
         ToolTipText     =   "Select User Class From List"
         Top             =   2520
         Width           =   3015
      End
      Begin VB.PictureBox imgSig 
         Height          =   945
         Left            =   960
         ScaleHeight     =   885
         ScaleWidth      =   4215
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Depending On Size, The Image May Appear Cropped. See The Report For Actual Image Representation"
         Top             =   3780
         Width           =   4275
      End
      Begin VB.TextBox txtNme 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Tag             =   "2"
         Top             =   780
         Width           =   4455
      End
      Begin VB.TextBox txtNik 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Tag             =   "2"
         Top             =   1140
         Width           =   1935
      End
      Begin VB.TextBox txtInt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         TabIndex        =   3
         Tag             =   "3"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox optAct 
         Alignment       =   1  'Right Justify
         Caption         =   "Active User?"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5400
         TabIndex        =   8
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "UserName"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Verify Password"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Authorization Role"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Authorization Report"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Signature :"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label a 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nickname"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Initials"
         Height          =   255
         Index           =   2
         Left            =   4680
         TabIndex        =   14
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnCerUpdUsr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7320
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   840
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   5160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5445
      FormDesignWidth =   8445
   End
   Begin VB.Label lblUsers 
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "AdmnCerUpdUsr"
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
Private Const COL_Group = 1
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

Private iUserIdx As Integer
Private iCurrentRec As Integer
Private iFreeIdx As Integer
Private iFreeDbf As Integer
Private strFileName As String

Private strUserName As String
Private strPrevRptName As String

Private Sub cmdCan_Click()
   
   AdmnCertUser.SetFocus
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      OpenHelpContext 30
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdLoad_Click()
   With MdiSect.Cdi
      .Filter = "Pictures(*.jpg;*.bmp)|*.jpg;*.bmp"
      .ShowOpen
   End With
   If Trim(MdiSect.Cdi.FileName) <> "" Then
      On Error GoTo DiaErr1
      'lblFile.Visible = True
      strFileName = MdiSect.Cdi.FileName
      imgSig.Picture = LoadPicture(strFileName)
   End If
   
   Exit Sub

DiaErr1:
   MsgBox "Couldn't Load Picture. Not A Valid File.", _
      vbInformation, Caption

End Sub

Private Sub SavePicture(strPicFile As String, strUserName As String, iRptID As Integer)
   'Dim RdoPic As rdoResultset
   Dim RdoPic As ADODB.Recordset
   
   DataFile = FreeFile
   On Error GoTo DiaErr1
   Open strPicFile For Binary Access Read As DataFile
   Fl = LOF(DataFile)
   
   sSql = "SELECT * FROM CertUsrSec WHERE (USRNAME='" & strUserName _
          & "' AND CERTRPTID =" & CStr(iRptID) & ")"
  ' bSqlRows = GetDataSet(RdoPic, ES_KEYSET)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPic, ES_KEYSET)
   
   If bSqlRows Then
   
      If (Err.Number = 0) Then
         'RdoPic.Edit
         Chunks = Fl \ ChunkSize
         Fragment = Fl Mod ChunkSize
         RdoPic!USRSIG.AppendChunk Null
         ReDim Chunk(Fragment)
         Get DataFile, , Chunk()
         
         RdoPic!USRSIG.AppendChunk Chunk()
         ReDim Chunk(ChunkSize)
         For i = 1 To Chunks
            Get DataFile, , Chunk()
            RdoPic!USRSIG.AppendChunk Chunk()
         Next
         Close DataFile
         RdoPic.Update
         On Error Resume Next
         RdoPic.Close
      Else
      
         MsgBox "Couldn't add Signature.", _
            vbExclamation, Caption
      
      End If
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "savepicture"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdSave_Click()

   On Error GoTo DiaErr1
   
   Dim bResponse As Byte
   Dim strUserName As String
   Dim strFullname As String
   Dim strNickName As String
   Dim strInit As String
   
   Dim strPassWd As String
   Dim strPassWd1 As String
   Dim strCertRpt As String
   Dim strCertRole As String
   Dim Active As Integer
   Dim iRptID As Integer
   Dim bret As Boolean
   
   strUserName = txtUserName.Text
   strFullname = txtNme.Text
   strNickName = txtNik.Text
   strInit = txtInt.Text
   strPassWd = txtPas.Text
   strPassWd1 = txtvPas.Text
   strCertRpt = cmbRpt.Text
   strCertRole = cmbRole.Text
   Active = optAct.Value
   
   
   
   If ((strPassWd = "") Or (strPassWd = "")) Then
      MsgBox ("The Passwords can not be empty. Please enter the password")
      Exit Sub
   End If
   
   If (strPassWd <> strPassWd1) Then
      MsgBox ("The Passwords are not same. Please re-enter the password")
      Exit Sub
   End If
   
   If (strFileName = "") Then
      bResponse = MsgBox("Signature file is not attached. Do you wna to continue..", ES_YESQUESTION, Caption)
      If bResponse = vbNo Then
         Exit Sub
      End If
   End If
   
   iRptID = GetReportID(strCertRpt)
   
   If (CInt(iRptID) = 0) Then
      MsgBox ("Not Valid Certification Report. Please contact Administrator.")
      Exit Sub
   End If
   
   bret = CheckExistUser(strUserName, iRptID)
   
   If (bret = True) Then
      Err.Clear
      
    '  RdoCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
     
      sSql = "UPDATE CertUsrSec  SET USRPASS = '" & strPassWd & "'," _
            & "USRFULLNAME = '" & strFullname & "',USRNICKNAME = '" & strNickName & "'," _
            & "USRINITIAL = '" & strInit & "' WHERE USRNAME = '" & strUserName & "' " _
               & " AND CERTRPTID = '" & CStr(iRptID) & "'"
   
   '   sSql = "INSERT INTO CertUsrSec (USRNAME, USRPASS, USRFULLNAME, USRNICKNAME, " _
   '            & " USRINITIAL, CERTRPTID,Active) VALUES ('" & strUserName & "','" & strPassWd & "','" _
   '          & strFullname & "','" & strNickName & "','" & strInit & "','" _
   '          & CStr(iRptID) & "','" & CStr(Active) & "')"
      
   '   RdoCon.Execute sSql, rdExecDirect
      clsADOCon.ExecuteSQL sSql
   
   
      If (strFileName <> "") Then
         SavePicture strFileName, strUserName, iRptID
      End If
      
      If (clsADOCon.ADOErrNum = 0) Then
         
         clsADOCon.CommitTrans
         SysMsg "Added Authetication Signature", True
         
      Else
         clsADOCon.RollbackTrans
         SysMsg "Signature Information was not able to add.", True
         
      End If
   Else
      MsgBox ("The User is not already assinged to this Report.")
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "cmdSave"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me


End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      iOldrec = iUserIdx
      
      cmbRole.AddItem "Director"
      cmbRole.AddItem "Inspector"
      cmbRole.AddItem "Operator"
      
      
      strUserName = AdmnCerUpdUsr.txtUserName
      strPrevRptName = AdmnCerUpdUsr.cmbRpt

      GetUserInfo strUserName, strPrevRptName
      bOnLoad = 0
      
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   FillReports
   bOnLoad = 1
   
End Sub

Private Function GetUserInfo(strUserName As String, strRptName As String)

   Dim iRptID As Integer
'   Dim RdoRpt As rdoResultset
   Dim RdoRpt As ADODB.Recordset
   
   
   On Error GoTo DiaErr1
   
   iRptID = GetReportID(strRptName)
   
   If (iRptID <> 0) Then
      
      sSql = "SELECT USRFULLNAME, USRNICKNAME, USRINITIAL,USRPASS FROM CertUsrSec " _
            & " WHERE USRNAME = '" & strUserName & "' AND CERTRPTID = " & CStr(iRptID)
      
 '     bSqlRows = GetDataSet(RdoRpt, ES_FORWARD)
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
     
      If bSqlRows Then
         
         With RdoRpt
            txtNme = Trim(!USRFULLNAME)
            txtNik = Trim(!USRNICKNAME)
            txtInt = Trim(!USRINITIAL)
            txtPas = Trim(!USRPASS)
            txtvPas = Trim(!USRPASS)
            
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
  ' Dim RdoPic As rdoResultset
   Dim RdoPic As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT USRSIG FROM CertUsrSec WHERE " _
            & " USRNAME = '" & strUserName & "' AND CERTRPTID = " & CStr(iRptID)
   
  ' bSqlRows = GetDataSet(RdoPic, ES_STATIC)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPic, ES_KEYSET)
  
   
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
         imgSig.Picture = LoadPicture(DiskFile)
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



Private Function CheckExistUser(strUserName As String, iRptID As Integer) As Boolean
'   Dim RdoRpt As rdoResultset
   Dim RdoRpt As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   If (strUserName = "") Then
      MsgBox ("The User name can not be empty.")
      CheckExistUser = False
      Exit Function
   End If
   
   sSql = "SELECT * FROM CertUsrSec WHERE " _
            & " USRNAME = '" & strUserName & "' AND CERTRPTID = " & CStr(iRptID)
   'bSqlRows = GetDataSet(RdoRpt, ES_FORWARD)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)

   If bSqlRows Then
      CheckExistUser = True
      ClearResultSet RdoRpt
   Else
      CheckExistUser = False
   End If
   Set RdoRpt = Nothing
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "FillReports"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Private Sub FillReports()

   'Dim RdoRpt As rdoResultset
   Dim RdoRpt As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "select REPORTNAME, REPORTID from CertReports"
   'bSqlRows = GetDataSet(RdoRpt, ES_FORWARD)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   
   If bSqlRows Then
      
      With RdoRpt
         Do Until .EOF
            AddComboStr cmbRpt.hwnd, "" & Trim(!ReportName)
            .MoveNext
         Loop
         ClearResultSet RdoRpt
      End With
      
   End If
   Set RdoRpt = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "FillReports"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set AdmnCerUpdUsr = Nothing
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub ResetControls(bOpen As Boolean)
   Dim iList As Integer
   txtNme.Enabled = bOpen
   txtNik.Enabled = bOpen
   txtInt.Enabled = bOpen
   optAct.Enabled = bOpen
   cmbRpt.Enabled = bOpen
End Sub

