VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form AdmnCerNUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Authorization User"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   420
      Left            =   7320
      TabIndex        =   22
      ToolTipText     =   "Add A New User"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Frame fraUser 
      Height          =   4935
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   7095
      Begin VB.TextBox txtUserName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Tag             =   "2"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtvPas 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         TabIndex        =   6
         ToolTipText     =   "Case Sensitive Max (15) Char"
         Top             =   2040
         Width           =   1965
      End
      Begin VB.ComboBox cmbRole 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Tag             =   "8"
         ToolTipText     =   "Select User Class From List"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox txtPas 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         TabIndex        =   5
         ToolTipText     =   "Case Sensitive Max (15) Char"
         Top             =   1560
         Width           =   1965
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load Sig"
         Height          =   420
         Left            =   4800
         TabIndex        =   10
         ToolTipText     =   "Add A New User"
         Top             =   3720
         Width           =   1200
      End
      Begin VB.ComboBox cmbRpt 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Tag             =   "8"
         ToolTipText     =   "Select User Class From List"
         Top             =   2520
         Width           =   3015
      End
      Begin VB.PictureBox imgSig 
         Height          =   945
         Left            =   240
         ScaleHeight     =   885
         ScaleWidth      =   4215
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Depending On Size, The Image May Appear Cropped. See The Report For Actual Image Representation"
         Top             =   3780
         Width           =   4275
      End
      Begin VB.TextBox txtNme 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Tag             =   "2"
         Top             =   780
         Width           =   4455
      End
      Begin VB.TextBox txtNik 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Tag             =   "2"
         Top             =   1140
         Width           =   1935
      End
      Begin VB.TextBox txtInt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   4
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
         Left            =   5160
         TabIndex        =   9
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "UserName"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Verify Password"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Authorization Role"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   24
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Authorization Report"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Signature :"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label a 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nickname"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Initials"
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   15
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnCerNUser.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double Click To Select A User (Or Select And Press Enter)"
      Top             =   720
      Width           =   5655
      _ExtentX        =   9975
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
      Left            =   7560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   840
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   6120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8265
      FormDesignWidth =   8685
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Users :"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblUsers 
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "AdmnCerNUser"
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
'   bSqlRows = GetDataSet(RdoPic, ES_KEYSET)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPic, ES_KEYSET)
 
   If bSqlRows Then
   
      If (clsADOCon.ADOErrNum = 0) Then
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
      MsgBox ("The User is already assinged to this Report.")
      Exit Sub
   End If
   
   Err.Clear
   
   clsADOCon.BeginTrans
   
   sSql = "INSERT INTO CertUsrSec (USRNAME, USRPASS, USRFULLNAME, USRNICKNAME, " _
            & " USRINITIAL, CERTRPTID,Active) VALUES ('" & strUserName & "','" & strPassWd & "','" _
          & strFullname & "','" & strNickName & "','" & strInit & "','" _
          & CStr(iRptID) & "','" & CStr(Active) & "')"
   
   'RdoCon.Execute sSql, rdExecDirect
   clsADOCon.ExecuteSQL sSql


   If (strFileName <> "") Then
      SavePicture strFileName, strUserName, iRptID
   End If
   
   If (Err.Number = 0) Then
      
      clsADOCon.CommitTrans
      SysMsg "Added Authetication Signature", True
      
   Else
      clsADOCon.RollbackTrans
      SysMsg "Signature Information was not able to add.", True
      
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
      FormatGrid
           
           
      cmbRole.AddItem "Director"
      cmbRole.AddItem "Inspector"
      cmbRole.AddItem "Operator"
      bOnLoad = 0
      
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   OpenDbfFiles
   FillReports
   bOnLoad = 1
   
End Sub

Private Function CheckExistUser(strUserName As String, iRptID As Integer) As Boolean
   'Dim RdoRpt As rdoResultset
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
   Dim RdoRpt As ADODB.Recordset
   'Dim RdoRpt As rdoResultset
   On Error GoTo DiaErr1
   sSql = "select REPORTNAME, REPORTID from CertReports"
'   bSqlRows = GetDataSet(RdoRpt, ES_FORWARD)
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
   Set AdmnCerNUser = Nothing
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Public Sub FormatGrid()
   Dim row As Integer
   Dim iList As Integer
   
   Grid1.Clear
   Grid1.Cols = COL_Count
   Grid1.ColWidth(COL_Name) = 1900
   Grid1.ColWidth(COL_Group) = 1550
   Grid1.ColWidth(COL_Active) = 900
   Grid1.ColWidth(COL_RecNo) = 0
   Grid1.row = 0
   Grid1.TextMatrix(0, COL_Name) = "User Id"
   Grid1.TextMatrix(0, COL_Group) = "Group"
   Grid1.TextMatrix(0, COL_Active) = "Active"
   Grid1.TextMatrix(0, COL_RecNo) = "Row"
   

   Grid1.Rows = 1
   For iList = FIRSTUSERRECORDNO To LOF(iFreeDbf) \ Len(Secure)
      Get #iFreeIdx, iList, SecPw
      Get #iFreeDbf, iList, Secure
      row = row + 1
      If row >= Grid1.Rows Then Grid1.Rows = Grid1.Rows + 1
      Grid1.TextMatrix(row, COL_Name) = SecPw.UserLcName
      Grid1.TextMatrix(row, COL_Group) = IIf(SecPw.UserAdmn = 1, "Administrator", "User")
      Grid1.TextMatrix(row, COL_Active) = IIf(Secure.UserActive = 1, "Active", "Inactive")
      Grid1.TextMatrix(row, COL_RecNo) = iList
   Next

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

Private Sub optAct_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub ResetControls(bOpen As Boolean)
   Dim iList As Integer
   txtNme.Enabled = bOpen
   txtNik.Enabled = bOpen
   txtInt.Enabled = bOpen
   optAct.Enabled = bOpen
   cmbRpt.Enabled = bOpen
End Sub

Private Sub OpenDbfFiles()
   On Error Resume Next
   Close iFreeIdx
   Close iFreeDbf
   iFreeIdx = FreeFile
   Open sFilePath & "rstval.eid" For Random Shared As iFreeIdx Len = Len(SecPw)
   iFreeDbf = FreeFile
   Open sFilePath & "rstval.edd" For Random Shared As iFreeDbf Len = Len(Secure)
   
End Sub

Private Sub GetSelectedUser()
   If Grid1.row = 0 Then
      If Grid1.Rows > 1 Then
         Grid1.row = 1
      Else
         Exit Sub
      End If
   End If
   
   iCurrentRec = Grid1.TextMatrix(Grid1.row, COL_RecNo)
   Get #iFreeIdx, iCurrentRec, SecPw
   Get #iFreeDbf, iCurrentRec, Secure
   txtUserName = Trim(SecPw.UserLcName)
   txtNme = Trim(Secure.UserName)
   txtNik = Trim(Secure.UserNickName)
   txtInt = Trim(Secure.UserInitials)
   optAct.Value = Secure.UserActive
   txtPas = Trim("")
   txtvPas = Trim("")
   
   ' Hide module button options
   'chkHideModule.Value = Secure.UserZHideModule
   
   ResetControls True
   
End Sub

