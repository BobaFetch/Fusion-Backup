VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DocuDCe06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update MO Document Lists"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCe06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   6240
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Update The Current Manufacturing Order Document List"
      Top             =   1560
      Width           =   870
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Revision-Select From List"
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Open MO's Not Canceled, Complete Or Closed"
      Top             =   720
      Width           =   3345
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6120
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Contains Open MO Runs Not Canceled, Complete Or Closed"
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2775
      FormDesignWidth =   7200
   End
   Begin VB.Label lblRev 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6120
      TabIndex        =   12
      ToolTipText     =   "Current Revision For The Selected Run"
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   11
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblDoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   1920
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Document List Rev"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Replace The Current Manufacturing Order Document List"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Part Number"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   9
      Left            =   5280
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "DocuDCe06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'4/28/04 new
'3/1/05 Added revision
Option Explicit
Dim bCancel As Byte
Dim bOnLoad As Byte

'Passed document stuff
Dim iDocEco As Integer
Dim sDocName As String
Dim sDocClass As String
Dim sDocSheet As String
Dim sDocDesc As String
Dim sDocAdcn As String


Dim sOldPart As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Click()
   If sOldPart <> cmbPrt Then GetTheDesc
   
End Sub


Private Sub cmbPrt_LostFocus()
   If bCancel = 0 Then
      cmbPrt = CheckLen(cmbPrt, 30)
      
      If (Not ValidPartNumber(cmbPrt.Text)) Then
         MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
            vbInformation, Caption
         cmbPrt = ""
         Exit Sub
      End If
      
      If sOldPart <> cmbPrt Then GetTheDesc
   End If
   
End Sub


Private Sub cmbRev_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   If bCancel = 1 Then Exit Sub
   cmbRev = CheckLen(cmbRev, 6)
   If cmbRev.ListCount > 0 Then
      For iList = 0 To cmbRev.ListCount - 1
         If Trim(cmbRev) = Trim(cmbRev.List(iList)) Then bByte = 1
      Next
      If bByte = 0 Then
         Beep
         cmbRev = cmbRev.List(0)
      End If
   End If
   
End Sub


Private Sub cmbRun_Click()
   GetCurrentListRev
   
End Sub


Private Sub cmbRun_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   
   If cmbRun.ListCount > 0 Then
      For iList = 0 To cmbRun.ListCount - 1
         If Val(cmbRun) = Val(cmbRun.List(iList)) Then bByte = 1
      Next
      If bByte = 0 Then
         Beep
         cmbRun = cmbRun.List(0)
      End If
   End If
   GetCurrentListRev
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3306
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sRev As String
   If Trim(cmbRev) = "" Then sRev = "NC" Else sRev = Trim(cmbRev)
   If lblDoc.ForeColor = ES_RED Then
      MsgBox "There Is No Current Document List For This Part.", _
         vbInformation, Caption
   Else
      sMsg = "Update The Selected Run To The List Of Rev " & sRev & "?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then UpdateMOList Else CancelTrans
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set DocuDCe06a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblDsc.ForeColor = vbBlack
   lblDoc.ForeColor = vbBlack
   sOldPart = ""
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_RunsNotLikeC"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetTheDesc()
   Dim sDescription As String
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   sOldPart = cmbPrt
   FillRuns
   
End Sub

Private Sub lblDoc_Change()
   If Left(lblDoc, 5) = "*** N" Then _
           lblDoc.ForeColor = ES_RED Else lblDoc.ForeColor = vbBlack
   
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 5) = "*** P" Then lblDsc.ForeColor = ES_RED _
           Else lblDsc.ForeColor = vbBlack
   
End Sub


Private Sub FillRuns()
   Dim RdoRns As ADODB.Recordset
   cmbRun.Clear
   On Error GoTo DiaErr1
   If lblDsc.ForeColor <> ES_RED Then
      sSql = "SELECT RUNREF,RUNNO,RUNSTATUS FROM RunsTable " _
             & "WHERE RUNREF='" & Compress(cmbPrt) & "' AND " _
             & "RUNSTATUS NOT LIKE 'C%'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns, ES_FORWARD)
      If bSqlRows Then
         With RdoRns
            Do Until .EOF
               AddComboStr cmbRun.hwnd, str$(!Runno)
               .MoveNext
            Loop
            ClearResultSet RdoRns
         End With
      End If
      If cmbRun.ListCount > 0 Then
         cmbRun = cmbRun.List(0)
         GetDocRevisions
         GetCurrentListRev
      Else
         lblRev = ""
      End If
   End If
   Set RdoRns = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetDocRevisions()
   Dim RdoRev As ADODB.Recordset
   If bCancel = 1 Then Exit Sub
   cmbRev.Clear
   On Error GoTo DiaErr1
   If lblDsc.ForeColor <> ES_RED Then
      sSql = "SELECT DISTINCT DLSREV FROM DlstTable WHERE " _
             & "DLSREF='" & Compress(cmbPrt) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRev, ES_FORWARD)
      If bSqlRows Then
         With RdoRev
            Do Until .EOF
               cmbRev.AddItem !DLSREV
               .MoveNext
            Loop
            ClearResultSet RdoRev
         End With
      End If
   End If
   If cmbRev.ListCount > 0 Then
      cmbRev = cmbRev.List(0)
      lblDoc = "Document List(s) For Current Part Number"
   Else
      lblDoc = "*** No Current Document Lists ***"
   End If
   Set RdoRev = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getdocrevisions"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub UpdateMOList()
   Dim RdoList As ADODB.Recordset
   
   Dim iRow As Integer
   Dim sDocRef As String
   Dim sListRef As String
   Dim sListRev As String
   
   sListRef = Compress(cmbPrt)
   sListRev = Compress(cmbRev)
   cmdUpd.Enabled = False
   MouseCursor 13
   On Error GoTo DiaErr1
   sSql = "DELETE FROM RndlTable WHERE RUNDLSRUNREF='" & sListRef & " ' AND " _
          & "RUNDLSRUNNO=" & Val(cmbRun) & " "
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
   sSql = "SELECT * FROM DlstTable WHERE DLSREF='" & sListRef & "' " _
          & "AND DLSREV='" & sListRev & "' ORDER BY DLSDOCCLASS,DLSDOCREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoList, ES_FORWARD)
   If bSqlRows Then
      With RdoList
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         Do Until .EOF
            iRow = iRow + 1
            sDocRef = GetDocInformation("" & Trim(!DLSDOCREF), "" & Trim(!DLSDOCREV))
            sProcName = "updatemolist"
            sSql = "INSERT INTO RndlTable (RUNDLSNUM,RUNDLSRUNREF," _
                   & "RUNDLSRUNNO,RUNDLSREV,RUNDLSDOCREF,RUNDLSDOCREV," _
                   & "RUNDLSDOCREFLONG,RUNDLSDOCREFDESC,RUNDLSDOCREFSHEET," _
                   & "RUNDLSDOCREFCLASS,RUNDLSDOCREFADCN," _
                   & "RUNDLSDOCREFECO) VALUES(" & iRow & ",'" & sListRef & "'," _
                   & Val(cmbRun) & ",'" & sListRev & "','" & Trim(!DLSDOCREF) & "','" _
                   & Trim(!DLSDOCREV) & "','" & sDocName & "','" & sDocDesc & "','" _
                   & sDocSheet & "','" & sDocClass & "','" & sDocAdcn & "'," _
                   & iDocEco & ")"
            clsADOCon.ExecuteSQL sSql ' rdExecDirect
            .MoveNext
         Loop
         ClearResultSet RdoList
      End With
      MouseCursor 0
      If clsADOCon.ADOErrNum = 0 Then
         lblRev = cmbRev
         SysMsg "Manufacturing Order Updated.", True
      Else
         MsgBox "Could Not Successfully Update The MO.", _
            vbExclamation, Caption
      End If
   Else
      'Dummy Row for joins
      sSql = "INSERT INTO RndlTable (RUNDLSNUM,RUNDLSRUNREF, RUNDLSRUNNO) " _
             & "VALUES(1,'" & sListRef & "'," & Val(cmbRun) & ")"
      clsADOCon.ExecuteSQL sSql ' rdExecDirect
   End If
   cmdUpd.Enabled = True
   Set RdoList = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "updatemolust"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetDocInformation(DocumentRef As String, DocumentRev As String) As String
   Dim RdoDoc As ADODB.Recordset
   sProcName = "getdocinfo"
   sSql = "SELECT DOREF,DONUM,DOREV,DOCLASS,DOSHEET,DODESCR,DOECO," _
          & "DOADCN FROM DdocTable where (DOREF='" & DocumentRef & "' " _
          & "AND DOREV='" & DocumentRev & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_FORWARD)
   If bSqlRows Then
      With RdoDoc
         GetDocInformation = "" & Trim(!DOREF)
         sDocName = "" & Trim(!DONUM)
         sDocClass = "" & Trim(!DOCLASS)
         sDocSheet = "" & Trim(!DOSHEET)
         sDocDesc = "" & Trim(!DODESCR)
         iDocEco = !DOECO
         sDocAdcn = "" & Trim(!DOADCN)
         ClearResultSet RdoDoc
      End With
      sDocName = CheckStrings(sDocName)
      sDocAdcn = CheckStrings(sDocAdcn)
   Else
      sDocName = ""
      sDocClass = ""
      sDocSheet = ""
      sDocDesc = ""
      iDocEco = 0
      sDocAdcn = ""
   End If
   Set RdoDoc = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getdocinfo"
   
End Function

Private Function CheckStrings(TestString As String) As String
   Dim iLen As Integer
   Dim K As Integer
   Dim PartNo As String
   Dim NewPart As String
   
   On Error GoTo modErr1
   PartNo = Trim$(TestString)
   iLen = Len(PartNo)
   If iLen > 0 Then
      For K = 1 To iLen
         If Mid$(PartNo, K, 1) = Chr$(34) Or Mid$(PartNo, K, 1) = Chr$(39) _
                 Or Mid$(PartNo, K, 1) = Chr$(44) Then
            Mid$(PartNo, K, 1) = "-"
         End If
      Next
   End If
   CheckStrings = PartNo
   Exit Function
   
modErr1:
   Resume modErr2
modErr2:
   On Error Resume Next
   CheckStrings = ""
End Function

Private Function GetCurrentListRev() As Byte
   Dim RdoCur As ADODB.Recordset
   sSql = "SELECT RUNDLSRUNREF,RUNDLSRUNNO,RUNDLSREV FROM " _
          & "RndlTable WHERE RUNDLSRUNREF='" & Compress(cmbPrt) _
          & "' AND RUNDLSRUNNO=" & Val(cmbRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCur, ES_FORWARD)
   If bSqlRows Then lblRev = "" & Trim(RdoCur!RUNDLSREV) _
                             Else lblRev = ""
   Set RdoCur = Nothing
   
End Function
