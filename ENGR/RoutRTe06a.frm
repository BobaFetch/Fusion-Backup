VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form RoutRTe06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routing Operation Pictures"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdunload 
      Caption         =   "&UnLoad"
      Height          =   315
      Left            =   7440
      TabIndex        =   20
      ToolTipText     =   "Load A Current Picture"
      Top             =   1800
      Width           =   915
   End
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   684
      Picture         =   "RoutRTe06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Print The Report"
      Top             =   7800
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   120
      Picture         =   "RoutRTe06a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Display The Report"
      Top             =   7800
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Test"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7440
      TabIndex        =   17
      ToolTipText     =   "Test The Report (Pictures By Routing Operation)"
      Top             =   2160
      Width           =   915
   End
   Begin VB.PictureBox Image1 
      Height          =   4572
      Left            =   240
      ScaleHeight     =   4515
      ScaleWidth      =   7995
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Depending On Size, The Image May Appear Cropped. See The Report For Actual Image Representation"
      Top             =   2880
      Width           =   8052
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTe06a.frx":0308
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
      Left            =   7440
      TabIndex        =   13
      ToolTipText     =   "Select A File To Be Saved"
      Top             =   1440
      Width           =   915
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   315
      Left            =   7440
      TabIndex        =   12
      ToolTipText     =   "Load A Current Picture"
      Top             =   1080
      Width           =   915
   End
   Begin VB.TextBox txtPic 
      Height          =   285
      Left            =   1800
      MaxLength       =   80
      TabIndex        =   3
      Tag             =   "8"
      Text            =   " "
      ToolTipText     =   "(80) Char Maximun"
      Top             =   2520
      Width           =   6492
   End
   Begin VB.TextBox txtCmt 
      Height          =   948
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "8"
      Text            =   "RoutRTe06a.frx":0AB6
      ToolTipText     =   "Operation Comments"
      Top             =   1440
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
      Height          =   288
      Left            =   6360
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Operation Number"
      Top             =   720
      WhatsThisHelpID =   100
      Width           =   948
   End
   Begin VB.ComboBox cmbRte 
      Height          =   288
      Left            =   1800
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Routing"
      Top             =   720
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
      FormDesignHeight=   7785
      FormDesignWidth =   8505
   End
   Begin VB.Label lblHelp 
      Caption         =   "Please Read The Subject Help For This Procedure"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   240
      TabIndex        =   15
      Top             =   300
      Visible         =   0   'False
      Width           =   4092
   End
   Begin VB.Label lblFile 
      Height          =   252
      Left            =   720
      TabIndex        =   11
      Top             =   6840
      Visible         =   0   'False
      Width           =   5532
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture Description"
      Height          =   288
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   1752
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation:"
      Height          =   288
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Tag             =   "4"
      Top             =   1440
      Width           =   912
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
      Height          =   372
      Index           =   1
      Left            =   5400
      TabIndex        =   6
      Top             =   720
      Width           =   912
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
Attribute VB_Name = "RoutRTe06a"
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
Const picturePath As String = "c:\Program Files\ES2000\Temp\"


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub PrintReport()
   MouseCursor 13
   On Error GoTo DiaErr1
   
   
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim cCRViewer As EsCrystalRptViewer
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("engrt07")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "OPNO"
   aFormulaName.Add "RequestBy"
   'aFormulaName.Add "ShowDetails"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & cmbOpno & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   'aFormulaValue.Add optDet.value
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{RtpcTable.OPREF} = '" & Compress(cmbRte) & "' AND {RtpcTable.OPNO} =" _
          & Val(cmbOpno) & " "
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Private Sub PrintReport()
'   MouseCursor 13
'   On Error GoTo DiaErr1
'   SetMdiReportsize MDISect
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "OPNO='" & cmbOpno & "'"
'   MDISect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
'   sCustomReport = GetCustomReport("engrt07")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   sSql = "{RtpcTable.OPREF} = '" & Compress(cmbRte) & "' AND {RtpcTable.OPNO} =" _
'          & Val(cmbOpno) & " "
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
'   MouseCursor 0
'   Exit Sub
'
'DiaErr1:
'   sProcName = "printreport"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub

Private Sub GetPicture()
   Dim DiskFile As String
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM RtpcTable WHERE (OPREF='" & Compress(cmbRte) _
          & "' AND OPNO=" & Val(cmbOpno) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPic, ES_STATIC)
   If Not bSqlRows Then
      MsgBox "There Is No Picture Row For This Item.", _
         vbInformation, Caption
   Else
      If IsNull(RdoPic!OPPICTURE) Then
         MsgBox "A Picture Has Not Been Assigned.", _
            vbInformation, Caption
      Else
         
         'DiskFile = "c:\Program Files\ES2000\Temp\picture1.jpg"
         CreateDirectoryPath (picturePath)
         DiskFile = picturePath + "picture1.jpg"
         If Len(Dir$(DiskFile)) > 0 Then Kill DiskFile
         
         DataFile = FreeFile
         Open DiskFile For Binary Access Write As DataFile
         'Fl = RdoPic!OPPICTURE.ColumnSize
         ' TODO - Test it
         Fl = RdoPic.Fields("OPPICTURE").ActualSize
         Chunks = Fl \ ChunkSize
         Fragment = Fl Mod ChunkSize
         ReDim Chunk(Fragment)
         'Chunk() = RdoPic!OPPICTURE.GetChunk(Fragment)
         Chunk() = RdoPic.Fields("OPPICTURE").GetChunk(Fragment)
         
         Put DataFile, , Chunk()
         For i = 1 To Chunks
            ReDim Buffer(ChunkSize)
            'Chunk() = RdoPic!OPPICTURE.GetChunk(ChunkSize)
            Chunk() = RdoPic.Fields("OPPICTURE").GetChunk(ChunkSize)
            Put DataFile, , Chunk()
         Next
         Close DataFile
         Image1.Picture = LoadPicture(DiskFile)
         cmdReport.Enabled = True
      End If
   End If
   RdoPic.Close
   lblFile.Visible = False
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
   
   CheckRoutingPic
    
   sSql = "SELECT * FROM RtpcTable WHERE (OPREF='" & Compress(cmbRte) _
          & "' AND OPNO=" & Val(cmbOpno) & ")"
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

Private Sub cmbOpno_Change()
   cmdReport.Enabled = False
   
End Sub

Private Sub cmbOpno_Click()
   GetOperation
   
End Sub


Private Sub cmbOpno_LostFocus()
   cmbOpno = Format(Abs(Val(cmbOpno)), "000")
   Dim bByte As Byte
   Dim iList As Integer
   cmbOpno = Format(Abs(Val(cmbOpno)), "000")
   For iList = 0 To cmbOpno.ListCount - 1
      If cmbOpno.List(iList) = cmbOpno Then bByte = 1
   Next
   If bByte = 0 Then
      Beep
      cmbOpno = cmbOpno.List(0)
   Else
      GetOperation
   End If
   
End Sub


Private Sub cmbRte_Click()
   bGoodRout = GetRouting()
   
End Sub


Private Sub cmbRte_LostFocus()
   bGoodRout = GetRouting()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdFile_Click()
   If bNoHelp = 0 Then
      MsgBox "Our Records Show That You Have Not Read The Subject Help.", _
         vbExclamation, Caption
   Else
      With MDISect.Cdi
         .filter = "Pictures(*.jpg;*.bmp)|*.jpg;*.bmp"
         .ShowOpen
      End With
      If Trim(MDISect.Cdi.FileName) <> "" Then
         On Error GoTo DiaErr1
         lblFile.Visible = True
         lblFile = MDISect.Cdi.FileName
         SavePicture
         GetPicture
      End If
   End If
   Exit Sub
DiaErr1:
   MsgBox "Couldn't Load Picture. Not A Valid File.", _
      vbInformation, Caption
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3106
      MouseCursor 0
      cmdHlp = False
   End If
   SaveSetting "Esi2000", "EsiEngr", "rte06a", 1
   bNoHelp = 1
   lblHelp.Visible = False
   
End Sub

Private Sub cmdLoad_Click()
   If bGoodRout = 0 Or bGoodOp = 0 Then Exit Sub
   If bNoHelp = 0 Then
      MsgBox "Our Records Show That You Have Not Read The Subject Help.", _
         vbExclamation, Caption
   Else
      SelectPicture
   End If
   
End Sub

Private Sub cmdReport_Click()
   PrintReport
   
End Sub

Private Sub cmdunload_Click()
   Dim strOpRef As String
   Dim strOpno As String
   
   strOpRef = Compress(cmbRte)
   strOpno = Trim(cmbOpno)
   
   If ((strOpRef <> "") And (strOpno <> "")) Then
      
      sSql = "DELETE FROM RtpcTable WHERE (OPREF='" & Compress(strOpRef) _
          & "' AND OPNO=" & Val(strOpno) & ")"

      clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      txtDsc = ""
      txtCmt = ""
      txtPic = ""
      txtPic.Enabled = False
      cmdFile.Enabled = False
      Image1.Picture = LoadPicture("")
      lblFile = ""
      bGoodPic = False
   End If
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bNoHelp = 1 ' TODO: MM GetSetting("Esi2000", "EsiEngr", "rte06a", bNoHelp)
      If bNoHelp = 0 Then lblHelp.Visible = True
      FillRoutings
      If cmbRte.ListCount > 0 Then bGoodRout = GetRouting()
      bOnLoad = 0
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
   FormUnload
   Set RoutRTe06a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDsc.BackColor = BackColor
   txtCmt.BackColor = BackColor
   
End Sub



Private Function GetRouting() As Byte
   Dim RdoRte As ADODB.Recordset
   txtDsc = ""
   txtCmt = ""
   txtPic = ""
   txtPic.Enabled = False
   cmdFile.Enabled = False
   Image1.Picture = LoadPicture("")
   lblFile = ""
   bGoodPic = False
   cmbOpno.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT RTREF,RTNUM,RTDESC FROM RthdTable WHERE " _
          & "RTREF='" & Compress(cmbRte) & " '"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      With RdoRte
         cmbRte = "" & Trim(!RTNUM)
         txtDsc = "" & Trim(!RTDESC)
         GetRouting = 1
         .Cancel
      End With
      ClearResultSet RdoRte
   Else
      cmdLoad.Enabled = False
      txtDsc = "*** Routing Wasn't Found ***"
      GetRouting = 0
   End If
   Set RdoRte = Nothing
   If GetRouting Then FillOperations
   Exit Function
   
DiaErr1:
   sProcName = "getrout"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtDsc_Change()
   If Left(txtDsc, 6) = "*** Ro" Then _
           txtDsc.ForeColor = ES_RED Else txtDsc.ForeColor = vbBlack
   
End Sub



Private Sub FillOperations()
   On Error GoTo DiaErr1
   sSql = "SELECT OPREF,OPNO FROM RtopTable WHERE " _
          & "OPREF='" & Compress(cmbRte) & "'"
   LoadNumComboBox cmbOpno, "000", 1
   If cmbOpno.ListCount > 0 Then
      bGoodOp = 1
      GetOperation
   Else
      bGoodOp = 0
      cmdLoad.Enabled = False
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "filloperations"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetOperation()
   Dim RdoOps As ADODB.Recordset
   lblFile = ""
   Image1.Picture = LoadPicture("")
   txtPic.Enabled = False
   bGoodPic = False
   On Error GoTo DiaErr1
   sSql = "SELECT OPREF,OPNO,OPCOMT FROM RtopTable WHERE OPREF='" _
          & Compress(cmbRte) & "' AND OPNO=" & Val(cmbOpno) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOps, ES_STATIC)
   If bSqlRows Then
      txtCmt = "" & Trim(RdoOps!OPCOMT)
      RdoOps.Cancel
      ClearResultSet RdoOps
      On Error Resume Next
      bGoodOp = 1
      txtPic.Enabled = True
      cmdLoad.Enabled = True
      cmdFile.Enabled = True
   Else
      bGoodOp = 0
      cmdLoad.Enabled = False
      cmdFile.Enabled = False
      txtCmt = ""
   End If
   Set RdoOps = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getoperation"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function SelectPicture() As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM RtpcTable WHERE (OPREF='" & Compress(cmbRte) _
          & "' AND OPNO=" & Val(cmbOpno) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPic, ES_STATIC)
   If Not bSqlRows Then
      sSql = "INSERT INTO RtpcTable (OPREF,OPNO) " _
             & "VALUES('" & Compress(cmbRte) & "'," _
             & Val(cmbOpno) & ")"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
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
          & "(OPREF='" & Compress(cmbRte) & "' AND OPNO=" & Val(cmbOpno) & ") "
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
End Sub


Private Function CheckRoutingPic()

   Dim RdoRec As ADODB.Recordset

   sSql = "SELECT * FROM RtpcTable WHERE (OPREF='" & Compress(cmbRte) _
          & "' AND OPNO=" & Val(cmbOpno) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRec)
   
   If Not bSqlRows Then
      
      sSql = "INSERT INTO RtpcTable (OPREF,OPNO) " _
          & "VALUES('" & Compress(cmbRte) & "'," _
          & Val(cmbOpno) & ")"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If
   
   ' close the recordset
   Set RdoRec = Nothing

End Function


