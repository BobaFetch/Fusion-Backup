VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form DocuDCf05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Global Document List Document Change"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCanSel 
      Caption         =   "&Clear All"
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      ToolTipText     =   "Delete This Revision"
      Top             =   2400
      Width           =   1275
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "Select &All"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Delete This Revision"
      Top             =   2400
      Width           =   1275
   End
   Begin VB.ComboBox cmbNewDoc 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      ToolTipText     =   "Select Part Number"
      Top             =   1320
      Width           =   5625
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   315
      Left            =   7920
      TabIndex        =   3
      ToolTipText     =   "Delete This Revision"
      Top             =   2760
      Width           =   915
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCf05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbOldDoc 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "Select Part Number"
      Top             =   480
      Width           =   5565
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7980
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   60
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   7200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6900
      FormDesignWidth =   9030
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Click The Row To Select A Partnumber to Re-Schedule MO"
      Top             =   2760
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   7011
      _Version        =   393216
      Rows            =   10
      Cols            =   3
      FixedCols       =   0
      BackColorSel    =   -2147483640
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
   End
   Begin VB.Image Chkno 
      Height          =   180
      Left            =   8040
      Picture         =   "DocuDCf05a.frx":07AE
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Chkyes 
      Height          =   180
      Left            =   8040
      Picture         =   "DocuDCf05a.frx":0805
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Doc Number"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblNewDoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   1680
      Width           =   5595
   End
   Begin VB.Label lblOldDoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doc Number to Replace"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "DocuDCf05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
'Dim RdoPrt As rdoQuery
Dim AdoCmdObj As ADODB.Command
'Dim RdoBmh As rdoResultset
Dim RdoBmh As ADODB.Recordset

Dim bGoodPart As Byte
Dim bGoodList As Byte
Dim bOnLoad As Byte
Dim sPartNumber As String
Dim sPartBomrev As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
   
End Sub

Private Sub cmbNewDoc_Click()
   lblNewDoc = GetDocDescription(cmbNewDoc)
End Sub

Private Sub cmbNewDoc_LostFocus()
   cmbOldDoc = CheckLen(cmbOldDoc, 30)
   lblOldDoc = GetDocDescription(cmbOldDoc)
End Sub

Private Sub cmbOldDoc_Change()
   lblOldDoc = GetDocDescription(cmbOldDoc)
   FillGrid
End Sub

Private Sub cmbOldDoc_Click()
   lblOldDoc = GetDocDescription(cmbOldDoc)
   FillGrid
End Sub

Private Sub cmbOldDoc_LostFocus()
   cmbOldDoc = CheckLen(cmbOldDoc, 30)
   lblOldDoc = GetDocDescription(cmbOldDoc)
   ' get the Partlist for the replace partnumner
'   Dim iPartType As Integer
'   iPartType = GetPartType(Compress(cmbOldDoc))
'
'   If (iPartType <> 0) Then
'
'      ClearGrid
'      'lblType = CStr(iPartType)
'      cmbNewDoc.Clear
'      FillPartListByType cmbNewDoc, iPartType
'   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub ClearGrid()
      
   grd.Clear
   grd.Rows = 2
   With grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .Rows = 1
      .row = 0
      .col = 0
      .Text = "Sel"
      .col = 1
      .Text = "Part Number"
      .col = 2
      .Text = "Part Description"
      .ColWidth(0) = 350
      .ColWidth(1) = 3000
      .ColWidth(2) = .Width - .ColWidth(0) - .ColWidth(1) - 400
      
   End With
      
End Sub

Private Sub cmdCanSel_Click()
   Dim iList As Integer
   For iList = 1 To grd.Rows - 1
       grd.col = 0
       grd.row = iList
       ' Only if the part is checked
       If grd.CellPicture = Chkyes.Picture Then
           Set grd.CellPicture = Chkno.Picture
       End If
   Next
End Sub


Private Sub cmdSelAll_Click()
   Dim iList As Integer
   For iList = 1 To grd.Rows - 1
       grd.col = 0
       grd.row = iList
       ' Only if the part is checked
       If grd.CellPicture = Chkno.Picture Then
           Set grd.CellPicture = Chkyes.Picture
       End If
   Next
End Sub

Private Sub cmdUpdate_Click()
   Dim bAssigned As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sNewRevision As String
   Dim iList As Integer
   Dim strDocList As String
   
   Dim oldNum As String, oldClass As String, oldSheet As String, oldRev As String
   ExtractDocInfo cmbOldDoc, oldNum, oldClass, oldSheet, oldRev
   Dim newNum As String, newClass As String, newSheet As String, newRev As String
   ExtractDocInfo cmbNewDoc, newNum, newClass, newSheet, newRev
   If oldNum = "" Or newNum = "" Then Exit Sub
   
   clsADOCon.BeginTrans
   
   ' Go through all the record in the grid and update the document references
   For iList = 1 To grd.Rows - 1
      grd.col = 0
      grd.row = iList
      ' Only if the part is checked
      If grd.CellPicture = Chkyes.Picture Then
         grd.col = 1
         strDocList = Compress(grd.Text)
        
         sSql = "UPDATE DlstTable set DLSDOCREF = '" & newNum & "'," & vbCrLf _
            & "DLSDOCCLASS = '" & newClass & "'," & vbCrLf _
            & "DLSDOCSHEET = '" & newSheet & "'," & vbCrLf _
            & "DLSDOCREV = '" & newRev & "'" & vbCrLf _
            & "WHERE DLSDOCREF = '" & oldNum & "'" & vbCrLf _
            & "AND DLSDOCCLASS = '" & oldClass & "'" & vbCrLf _
            & "AND DLSDOCSHEET = '" & oldSheet & "'" & vbCrLf _
            & "AND DLSDOCREV = '" & oldRev & "'"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
     
      End If
    Next
            
   If Err = 0 Then
   
      clsADOCon.CommitTrans
      SysMsg "Document has been replaced.", True, Me
   Else
      clsADOCon.RollbackTrans
      SysMsg "Document couldn't be replaced.", True, Me
   End If

   MouseCursor 0
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3251
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub FillGrid()

   Dim rs As ADODB.Recordset
   'grd.Clear
   
   'On Error Resume Next
   grd.Rows = 1
   
   Dim oldNum As String, oldClass As String, oldSheet As String, oldRev As String
   'Dim s As String
   'strOldDoc = Compress(cmbOldDoc)
   'If strOldDoc = "" Then Exit Sub
   ExtractDocInfo cmbOldDoc, oldNum, oldClass, oldSheet, oldRev
   If oldNum = "" Then Exit Sub
   
   'ClearGrid
   
   'If (strOldDoc <> strNewDoc) Then
      sSql = "select rtrim(pt.PARTNUM) as PARTNUM, rtrim(pt.PADESC) as PADESC" & vbCrLf
      sSql = sSql & "from DlstTable dl" & vbCrLf
      sSql = sSql & "join DlsthdTable dh on dl.DLSREF = dh.DLSTREF" & vbCrLf
      sSql = sSql & "join PartTable pt on pt.PARTREF = dh.DLSTREF" & vbCrLf
      sSql = sSql & "where dl.DLSDOCREF = '" & oldNum & "'" & vbCrLf
      sSql = sSql & "order by pt.PARTNUM"
      
      bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
      If bSqlRows Then
         With rs
            Do Until .EOF
               grd.Rows = grd.Rows + 1
               grd.row = grd.Rows - 1
               grd.col = 0
               Set grd.CellPicture = Chkno.Picture
               grd.col = 1
               grd.Text = "" & Trim(!PartNum)
               grd.col = 2
               grd.Text = "" & Trim(!PADESC)
               
               .MoveNext
            Loop
            ClearResultSet rs
         End With
'      Else
'         MsgBox "There are no document lists for document " & strOldDoc, _
'            vbInformation, Caption
      End If
      Set rs = Nothing
   'End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      MouseCursor 13
      FillDocs cmbOldDoc, True
      FillDocs cmbNewDoc, False
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub
'Private Function GetPartType(strPartNum As String) As Integer
'   'Dim RdoType As rdoResultset
'   Dim RdoType As ADODB.Recordset
'   On Error GoTo modErr1
'   sSql = "SELECT PALEVEL FROM PartTable WHERE PARTREF = '" & Compress(strPartNum) & "'"
'  ' bSqlRows = GetDataSet(RdoType, ES_FORWARD)
'   bSqlRows = clsADOCon.GetDataSet(sSql, RdoType, ES_FORWARD)
'
'    If bSqlRows Then
'      GetPartType = RdoType!PALEVEL
'      ClearResultSet RdoType
'   Else
'      GetPartType = -1
'   End If
'   Set RdoType = Nothing
'   Exit Function
'
'modErr1:
'   sProcName = "GetPartType"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors MDISect.ActiveForm
'
'End Function
'Private Sub FillPartListByType(Cntrl As Control, Ptypenum As Integer)
'   'Dim RdoFp4 As rdoResultset
'
'   Dim RdoFp4 As ADODB.Recordset
'   On Error GoTo modErr1
'   sSql = "SELECT PARTNUM FROM PartTable WHERE PALEVEL = " & Ptypenum
''   bSqlRows = GetDataSet(RdoFp4, ES_FORWARD)
'   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFp4, ES_FORWARD)
'
'   If bSqlRows Then
'      With RdoFp4
'         Do Until .EOF
'            AddComboStr Cntrl.hwnd, "" & Trim(!PartNum)
'            .MoveNext
'         Loop
'         ClearResultSet RdoFp4
'      End With
'   End If
'   If Cntrl.ListCount > 0 Then Cntrl = Cntrl.List(0)
'   Set RdoFp4 = Nothing
'   Exit Sub
'
'modErr1:
'   sProcName = "FillPartListByType"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors MDISect.ActiveForm
'
'End Sub


Private Sub Grd_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      grd.col = 0
      If grd.CellPicture = Chkyes.Picture Then
         Set grd.CellPicture = Chkno.Picture
      Else
         Set grd.CellPicture = Chkyes.Picture
      End If
   End If
   
End Sub

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   grd.col = 0
   If grd.CellPicture = Chkyes.Picture Then
      Set grd.CellPicture = Chkno.Picture
   Else
      Set grd.CellPicture = Chkyes.Picture
   End If
   
End Sub


Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   
'   With grd
'      .ColAlignment(0) = 0
'      .ColAlignment(1) = 0
'      .Rows = 1
'      .row = 0
'      .col = 0
'      .Text = "Sel"
'      .col = 1
'      .Text = "Part Number"
'      .col = 2
'      .Text = "Part Description"
'      .ColWidth(0) = 700
'      .ColWidth(1) = 1200
'      .ColWidth(2) = 4000
'
'   End With

   ClearGrid

   MouseCursor 0
   
   
   bOnLoad = 1
   
End Sub

'
Private Sub Form_Resize()
   Refresh

End Sub


Private Sub Form_Unload(Cancel As Integer)
   SaveCurrentSelections
   On Error Resume Next
   FormUnload
   Set BompBMf05a = Nothing
   
End Sub

Private Sub FillDocs(combo As ComboBox, onlyIfInDocList As Boolean)
   Dim Ado As ADODB.Recordset
   combo.Clear
   If onlyIfInDocList Then
      sSql = "select distinct (rtrim(DONUM) + '/'  +  rtrim(DOCLASS) + '/' + rtrim(DOSHEET) + '/' + rtrim(DOREV)) as DONUM" & vbCrLf _
         & "from DDocTable doc" & vbCrLf _
         & "join DlstTable dl on dl.DLSDOCREF = doc.DOREF and dl.DLSDOCCLASS = doc.DOCLASS" & vbCrLf _
         & "and dl.DLSDOCSHEET = doc.DOSHEET and dl.DLSDOCREV = doc.DOREV" & vbCrLf _
         & "join DlsthdTable dh on dl.DLSREF = dh.DLSTREF" & vbCrLf _
         & "order by DONUM"
   Else
      sSql = "select (rtrim(DONUM) + '/'  +  rtrim(DOCLASS) + '/' + rtrim(DOSHEET) + '/' + rtrim(DOREV)) as DONUM from DDocTable order by DONUM"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, Ado, ES_FORWARD)
   If bSqlRows Then
      With Ado
         Do Until .EOF
            Dim item As ListItem
            combo.AddItem !DONUM
            .MoveNext
         Loop
         ClearResultSet Ado
      End With
   End If
'   If combo.ListCount <> 0 Then combo.ListIndex = 0
   Set Ado = Nothing
End Sub

Private Function GetDocDescription(DocInfo As String) As String
   Dim rs As ADODB.Recordset
   Dim list() As String
   GetDocDescription = ""
   list = Split(Trim(DocInfo), "/")
   If UBound(list) < 0 Then Exit Function
   sSql = "select rtrim(DODESCR) as DODESCR from DDocTable where DOREF='" & Compress(list(0)) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
   If bSqlRows Then
      GetDocDescription = rs(0)
      ClearResultSet rs
   End If
   Set rs = Nothing
End Function

Private Sub ExtractDocInfo(info As String, ByRef num As String, ByRef Class As String, _
   ByRef sheet As String, ByRef rev As String)

   'example:   000-310-6259  SHTS 1  2  3 (class='DWG' sheet='' rev='NC')
   Dim list() As String
   num = ""
   Class = ""
   sheet = ""
   rev = ""
   Dim str As String, i As Integer
   str = Trim(info)
   list = Split(Trim(info), "/")
   If UBound(list) < 3 Then Exit Sub
   num = Compress(list(0))
   Class = list(1)
   sheet = list(2)
   rev = list(3)
End Sub
