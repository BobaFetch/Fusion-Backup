VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form BompBMf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy A Bill Of Material"
   ClientHeight    =   3270
   ClientLeft      =   1845
   ClientTop       =   1710
   ClientWidth     =   7050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   3250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "BompBMf01a.frx":07AE
      Height          =   320
      Left            =   5040
      Picture         =   "BompBMf01a.frx":0C4C
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Parts List for Part and Revision"
      Top             =   1080
      Width           =   350
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6120
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3270
      FormDesignWidth =   7050
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "C&opy"
      Height          =   315
      Left            =   6120
      TabIndex        =   5
      ToolTipText     =   "Copy The Old Bill To The New Bill"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtRev 
      Height          =   285
      Left            =   6000
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "To This Revision"
      Top             =   1800
      Width           =   735
   End
   Begin VB.CheckBox optDef 
      Alignment       =   1  'Right Justify
      Caption         =   "Set As Default "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1635
   End
   Begin VB.ComboBox cmbNpl 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Copy To This Part"
      Top             =   1800
      Width           =   3375
   End
   Begin VB.ComboBox cmbPls 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Part Number"
      Top             =   1080
      Width           =   3405
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   6000
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Revision-Select From List"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6135
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1560
      TabIndex        =   16
      Top             =   2880
      Width           =   3012
      _ExtentX        =   5318
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblDsc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblDsc1 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   12
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label lblDsc1 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   11
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Bill"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev:"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Bill"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev:"
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   6
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "BompBMf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/3/06 Created CopyPartsList and revamped entire Copy Routine
Option Explicit
'Dim RdoStm As rdoQuery
'Dim RdoPls As rdoQuery
Dim AdoCmdStm As ADODB.Command
Dim AdoCmdPls As ADODB.Command

Dim bGoodCur As Byte
Dim bGoodNew As Byte
Dim bHeaderIs As Byte
Dim bDiffPn As Byte
Dim bOnLoad As Byte

Dim sHeader As String
Dim sPartNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbNpl_Click()
   bGoodNew = GetNewPart()
   
End Sub

Private Sub cmbNpl_LostFocus()
   cmbNpl = CheckLen(cmbNpl, 30)
   bGoodNew = GetNewPart()
   
End Sub

Private Sub cmbPls_Click()
   bGoodCur = GetCurPart()
   
End Sub

Private Sub cmbPls_LostFocus()
   cmbPls = CheckLen(cmbPls, 30)
   bGoodCur = GetCurPart()
   
End Sub

Private Sub cmbRev_LostFocus()
   cmbRev = CheckLen(cmbRev, 4)
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3250
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdNew_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If bGoodNew = 0 Then
      cmdNew.Enabled = False
      Exit Sub
   End If
   
   If Compress(cmbPls) = Compress(cmbNpl) Then
      If Trim(cmbRev) = Trim(txtRev) Then
         cmdNew.Enabled = False
         Exit Sub
      End If
   End If
   
   If bHeaderIs Or Not bGoodCur Then Exit Sub
   On Error GoTo DiaErr1
   sMsg = "Create A Bill Of Material Rev " & txtRev & " For Part Number " & cmbNpl & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      CopyPartsList
   Else
      CancelTrans
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "cmdnew_click"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   Exit Sub
   
DiaErr3:
   On Error Resume Next
   clsADOCon.RollbackTrans
   MouseCursor 0
   MsgBox Err.Description & vbCrLf _
      & "Header Copied. Could Not Copy List.", vbExclamation, Caption
   
End Sub

Private Sub cmdVew_Click()
   If cmdVew Then
      lblDsc = lblDsc1(0)
      ViewBomTree.Show
      cmdVew = False
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillParts
      If cmbPls.ListCount > 0 Then cmbPls = cmbPls.List(0)
      'If cUR.CurrentPart <> "" Then cmbPls = cUR.CurrentPart
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PABOMREV FROM " _
          & "PartTable WHERE PARTREF= ? AND PALEVEL<4"
   
   Set AdoCmdStm = New ADODB.Command
   AdoCmdStm.CommandText = sSql
   
   Dim prmPrtRef As ADODB.Parameter
   Set prmPrtRef = New ADODB.Parameter
   prmPrtRef.Type = adChar
   prmPrtRef.Size = 30
   AdoCmdStm.Parameters.Append prmPrtRef
   
   'Set RdoStm = RdoCon.CreatePreparedStatement("", sSql)
   'TODO:
   'RdoStm.MaxRows = 1
   
   sSql = "SELECT BMHREV FROM BmhdTable WHERE BMHREF= ? ORDER BY BMHREV"
   
   Set AdoCmdPls = New ADODB.Command
   AdoCmdPls.CommandText = sSql
   
   Dim prmBMRef As ADODB.Parameter
   Set prmBMRef = New ADODB.Parameter
   prmBMRef.Type = adChar
   prmBMRef.Size = 30
   AdoCmdPls.Parameters.Append prmBMRef
   'Set RdoPls = RdoCon.CreatePreparedStatement("", sSql)
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   SaveCurrentSelections
   On Error Resume Next
   FormUnload
   'Set RdoStm = Nothing
   Set AdoCmdStm = Nothing
   Set AdoCmdPls = Nothing
   
   Set BompBMf01a = Nothing
   
End Sub

Private Sub FillParts()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_SortedPartTypesBelow4"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         cmbPls = "" & Trim(!PartNum)
         cmbNpl = "" & Trim(!PartNum)
         Do Until .EOF
            AddComboStr cmbPls.hwnd, "" & Trim(!PartNum)
            AddComboStr cmbNpl.hwnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   Set RdoCmb = Nothing
   bGoodCur = GetCurPart()
   bGoodNew = GetNewPart()
   Exit Sub
   
DiaErr1:
   sProcName = "fillparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Function GetCurPart() As Byte
   Dim RdoCur As ADODB.Recordset
   sPartNumber = Compress(cmbPls)
   cmbRev.Clear
   cmbRev = ""
   cmdNew.Enabled = False
   AdoCmdStm.Parameters(0) = sPartNumber
   On Error GoTo DiaErr1
   bSqlRows = clsADOCon.GetQuerySet(RdoCur, AdoCmdStm)
   If bSqlRows Then
      With RdoCur
         cmbPls = "" & Trim(!PartNum)
         cmbRev = "" & Trim(!PABOMREV)
         lblDsc1(0) = "" & Trim(!PADESC)
         ClearResultSet RdoCur
      End With
      cUR.CurrentPart = cmbPls
      GetCurPart = True
   Else
      lblDsc1(0) = ""
      If Not bOnLoad Then MsgBox "Part Wasn't Found or Wrong Type.", vbExclamation, Caption
      GetCurPart = False
   End If
   If GetCurPart Then
      'RdoPls(0) = sPartNumber
      AdoCmdPls.Parameters(0) = sPartNumber
      bSqlRows = clsADOCon.GetQuerySet(RdoCur, AdoCmdPls)
      If bSqlRows Then
         With RdoCur
            Do Until .EOF
               AddComboStr cmbRev.hwnd, "" & Trim(!BMHREV)
               .MoveNext
            Loop
            ClearResultSet RdoCur
         End With
      Else
         sPartNumber = ""
      End If
   End If
   Set RdoCur = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcurpar"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Function GetNewPart() As Byte
   Dim RdoNew As ADODB.Recordset
   txtRev = ""
   sPartNumber = Compress(cmbNpl)
   On Error GoTo DiaErr1
   AdoCmdStm.Parameters(0) = sPartNumber
   'RdoStm(0) = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoNew, AdoCmdStm)
   If bSqlRows Then
      With RdoNew
         cmbNpl = "" & Trim(!PartNum)
         txtRev = "" & Trim(!PABOMREV)
         lblDsc1(1) = "" & Trim(!PADESC)
         ClearResultSet RdoNew
      End With
      GetNewPart = 1
   Else
      cmdNew.Enabled = False
      txtRev = ""
      lblDsc1(1) = "*** Not A Valid Part Number ***"
      GetNewPart = 0
   End If
   Set RdoNew = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getnewpar"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblDsc1_Change(Index As Integer)
   If Left(lblDsc1(1), 7) = "*** Not" Then _
           lblDsc1(1).ForeColor = ES_RED Else lblDsc1(1).ForeColor = vbBlack
   
End Sub

Private Sub optDef_GotFocus()
   If bHeaderIs Then MsgBox "The Parts List Already Exists.", vbInformation, Caption
   
End Sub

Private Sub optDef_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtRev_LostFocus()
   txtRev = CheckLen(txtRev, 4)
   bDiffPn = 0
   bHeaderIs = GetRevision()
   If bHeaderIs Then
      cmdNew.Enabled = False
   Else
      If lblDsc1(1).ForeColor <> ES_RED Then cmdNew.Enabled = True
   End If
   
End Sub



Private Function GetRevision() As Byte
   Dim RdoRev As ADODB.Recordset
   sPartNumber = Compress(cmbNpl)
   txtRev = Compress(txtRev)
   sHeader = txtRev
   If Not bGoodNew Then Exit Function
   On Error GoTo DiaErr1
   If Trim(cmbPls) <> Trim(cmbNpl) Then
      bDiffPn = 1
      GetRevision = 0
      Exit Function
   End If
   sSql = "SELECT BMHREF,BMHREV FROM BmhdTable " _
          & "WHERE BMHREF='" & sPartNumber & "' AND BMHREV='" & sHeader & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRev)
   If bSqlRows Then
      With RdoRev
         txtRev = "" & Trim(!BMHREV)
         GetRevision = 1
         ClearResultSet RdoRev
      End With
   Else
      sHeader = ""
      GetRevision = 0
   End If
   Set RdoRev = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrevisi"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

'10/3/06 New

Private Sub CopyPartsList()
   Dim sOldPart As String
   Dim sNewRevision As String
   Dim sOldRevision As String
   
   MouseCursor 13
   sOldPart = Compress(cmbPls)
   sPartNumber = Compress(cmbNpl)
   sOldRevision = Compress(cmbRev)
   sNewRevision = Trim(txtRev)
   
   On Error Resume Next
   'In case the temp table remains
   sSql = "DROP TABLE ##Bmpl"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   prg1.Visible = True
   prg1.Value = 10
   
   'New Header
   Err.Clear
   clsADOCon.ADOErrNum = 0
   
'   sSql = "INSERT INTO BmhdTable (BMHREF,BMHPARTNO,BMHREV,BMHPART) " _
'          & "VALUES('" & sPartNumber & "','" & cmbNpl & "','" & sNewRevision & "','" & sPartNumber & "')"
'   clsADOCon.ExecuteSql sSql 'rdExecDirect
'   If clsADOCon.ADOErrNum <> 0 Then
'      clsADOCon.ADOErrNum = 0
'      sSql = "DELETE FROM BmplTable WHERE BMASSYPART='" _
'             & sPartNumber & "' AND BMREV='" & sNewRevision & "'"
'      clsADOCon.ExecuteSql sSql 'rdExecDirect
'   End If
'   Err.Clear
   
   ' if BOM exists, overwrite its part list
   Dim ado As ADODB.Recordset
   sSql = "select * from BmhdTable where BMHREF = '" & sPartNumber & "'" & vbCrLf _
      & "and BMHREV = '" & sNewRevision & "'"
   If clsADOCon.GetDataSet(sSql, ado) Then
      Set ado = Nothing
      sSql = "DELETE FROM BmplTable WHERE BMASSYPART='" _
             & sPartNumber & "' AND BMREV='" & sNewRevision & "'"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   ' BOM does not exist.  create it
   Else
      Set ado = Nothing
      sSql = "INSERT INTO BmhdTable (BMHREF,BMHPARTNO,BMHREV,BMHPART) " _
             & "VALUES('" & sPartNumber & "','" & cmbNpl & "','" & sNewRevision & "','" & sPartNumber & "')"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If
   If clsADOCon.ADOErrNum <> 0 Then
      clsADOCon.ADOErrNum = 0
      sSql = "DELETE FROM BmplTable WHERE BMASSYPART='" _
             & sPartNumber & "' AND BMREV='" & sNewRevision & "'"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If
   Err.Clear
   prg1.Value = 30
   Sleep 200
   
   sSql = "SELECT * INTO ##Bmpl from BmplTable where BMASSYPART='" & sOldPart & "' " _
          & "AND BMREV='" & sOldRevision & "'"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   sSql = "UPDATE ##Bmpl SET BMASSYPART='" & sPartNumber & "',BMREV='" & sNewRevision & "'"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   prg1.Value = 80
   Sleep 200
   
   sSql = "INSERT INTO BmplTable SELECT * FROM ##Bmpl"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   prg1.Value = 100
   MouseCursor 0
   If clsADOCon.ADOErrNum = 0 Then
      If optDef Then
         sSql = "UPDATE PartTable SET PABOMREV='" & sNewRevision & "' " _
                & "WHERE PARTREF='" & sPartNumber & "' "
         clsADOCon.ExecuteSql sSql 'rdExecDirect
      End If
      MsgBox "The Parts List Has Been Successfully Copied.", _
         vbInformation, Caption
   Else
      MsgBox "The Parts List Could Not Be Successfully Copied.", _
         vbExclamation, Caption
      MsgBox Err.Description
   End If
   prg1.Visible = False
   
End Sub
