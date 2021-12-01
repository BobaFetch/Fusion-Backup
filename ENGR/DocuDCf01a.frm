VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DocuDCf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy A Document List"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox cmbNrev 
      Height          =   288
      Left            =   5760
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "New Document Revision"
      Top             =   2160
      Width           =   972
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbTyp 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   1440
      TabIndex        =   17
      Tag             =   "8"
      ToolTipText     =   "Select Type From List"
      Top             =   720
      Width           =   2000
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   5760
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Document Revision"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox cmbNprt 
      Height          =   288
      Left            =   1440
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Copy To This Part"
      Top             =   2160
      Width           =   3435
   End
   Begin VB.CommandButton cmdCpy 
      Caption         =   "&Copy"
      Height          =   315
      Left            =   6120
      TabIndex        =   4
      ToolTipText     =   "Copy Document List"
      Top             =   720
      Width           =   870
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   288
      Left            =   1440
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Current Parts With A Document List"
      Top             =   1200
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3240
      FormDesignWidth =   7050
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   3
      Left            =   4920
      TabIndex        =   18
      ToolTipText     =   "New Document Revision"
      Top             =   2160
      Width           =   972
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   9
      Left            =   4920
      TabIndex        =   16
      ToolTipText     =   "Document Revision"
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label lblNtyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   5760
      TabIndex        =   15
      Top             =   2520
      Width           =   372
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   6
      Left            =   4920
      TabIndex        =   14
      Top             =   2520
      Width           =   612
   End
   Begin VB.Label lblNDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1440
      TabIndex        =   13
      Top             =   2520
      Width           =   3072
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy To:"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   2184
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Type"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   1332
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1440
      TabIndex        =   9
      Top             =   1548
      Width           =   3072
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1332
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   5760
      TabIndex        =   7
      Top             =   1560
      Width           =   372
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H00400000&
      Height          =   252
      Index           =   4
      Left            =   4920
      TabIndex        =   6
      Top             =   1560
      Width           =   612
   End
End
Attribute VB_Name = "DocuDCf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/3/06 Completely revised the CopyDocumentList procedure
Option Explicit
Dim bCanceled As Byte
Dim bOnLoad As Byte
Dim bGoodOld As Byte
Dim bGoodNew As Byte
Dim iType As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbNprt_Click()
   bGoodNew = GetNewPart()
   
End Sub


Private Sub cmbNprt_LostFocus()
   cmbNprt = CheckLen(cmbNprt, 30)
   bGoodNew = GetNewPart()
   If Not bGoodNew Then MsgBox "Part Number Wasn't Found.", vbExclamation, Caption
   
End Sub


Private Sub cmbNrev_LostFocus()
   cmbNrev = CheckLen(cmbNrev, 6)
   
End Sub


Private Sub cmbPrt_Click()
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   If Len(cmbPrt) > 0 Then bGoodOld = True
   GetOldDocRevisions
   
End Sub


Private Sub cmbPrt_LostFocus()
   If bCanceled Then Exit Sub
   cmbPrt = CheckLen(cmbPrt, 30)
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   FindPart
   If cmbPrt = "NONE" Or cmbPrt = "" Then
      bGoodOld = False
      MsgBox "Part Number Wasn't Found.", vbExclamation, Caption
   Else
      bGoodOld = True
   End If
   If bGoodOld Then FillDocumentRevisions Me Else cmbRev.Clear
   
End Sub


Private Sub cmbTyp_Click()
'   iType = cmbTyp.ListIndex
'   If iType < 1 Then iType = 1
   iType = Val(Left(cmbTyp, 1))
   If Val(Left(cmbTyp, 1)) = 1 Then
      cmbRev.Visible = True
      cmbNrev.Visible = True
      z1(3).Visible = True
      z1(9).Visible = True
   Else
      cmbRev.Visible = False
      cmbNrev.Visible = False
      z1(3).Visible = False
      z1(9).Visible = False
   End If
   FillCombo
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdCpy_Click()
   If bGoodOld And bGoodNew Then
      CopyDocumentList
   Else
      MsgBox "Requires Valid Document List And Part Numbers.", vbExclamation, Caption
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3350
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bCanceled = False
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()

    Sleep (500)
   FormLoad Me
   FormatControls
   
   ' cmbTyp.AddItem "Parts (Assign To A Part)"     'Type  Not Used
   cmbTyp.AddItem "1 - Parts                   " 'Type 1 Bom Revision
   cmbTyp.AddItem "2 - Service Parts           " 'Type 2"
   cmbTyp = cmbTyp.List(0)
   iType = 1
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set DocuDCf01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   cmbPrt.Clear
   cmbNprt.Clear
   cmbRev.Clear
   MouseCursor 13
   On Error GoTo DiaErr1
   iType = Val(Left(cmbTyp, 1))
   If iType < 1 Then iType = 1
   If iType = 1 Then
      sSql = "SELECT DISTINCT PARTREF,PARTNUM,DLSREF FROM PartTable" & vbCrLf _
         & "join DlstTable on PARTREF = DLSREF" & vbCrLf _
         & "ORDER BY PARTREF"
   Else
      sSql = "SELECT DISTINCT PARTREF,PARTNUM,DLSREF FROM PartTable" & vbCrLf _
         & "join DlstTable on PARTREF = DLSREF" & vbCrLf _
         & "where PALEVEL = 7" & vbCrLf _
         & "ORDER BY PARTREF"
   End If
   LoadComboBox cmbPrt
   If bSqlRows Then
      cmbPrt = cmbPrt.List(0)
   Else
      MouseCursor 0
      MsgBox "There Are No Current Document Assignments.", vbInformation, Caption
      Exit Sub
   End If
   If cmbPrt.ListCount > 0 Then cmbPrt_Click
   If iType = 1 Then
      sSql = "SELECT PARTREF,PARTNUM,PALEVEL,PABOMREV FROM PartTable WHERE PALEVEL<>6"
   Else
      sSql = "SELECT PARTREF,PARTNUM,PALEVEL,PABOMREV FROM PartTable WHERE PALEVEL=7"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         cmbNprt = "" & Trim(!PartNum)
         Do Until .EOF
            AddComboStr cmbNprt.hwnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   Set RdoCmb = Nothing
   If cmbNprt.ListCount > 0 Then cmbNprt_Click
   MouseCursor 0
   If cmbPrt.ListCount > 0 Then FillDocumentRevisions Me
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetNewPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   Dim sGetPart As String
   
   sGetPart = Compress(cmbNprt)
   On Error GoTo DiaErr1
   If Len(sGetPart) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL FROM PartTable WHERE PARTREF='" & sGetPart & "' " _
             & "AND PALEVEL<>6"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
      If bSqlRows Then
         With RdoPrt
            cmbNprt = "" & Trim(!PartNum)
            lblNDsc = "" & Trim(!PADESC)
            lblNtyp = Format(0 + !PALEVEL, "0")
         End With
         GetNewPart = True
      Else
         cmbNprt = ""
         lblNDsc = ""
         lblNtyp = ""
         GetNewPart = False
      End If
   End If
   On Error Resume Next
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getnewpar"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub CopyDocumentList()
   Dim RdoLst As ADODB.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sOldPart As String
   Dim sNewPart As String
   Dim sOldRev As String 'Not used yet
   Dim sNewRev As String
   
   'iType = Val(lblNtyp)
   iType = Val(Left(cmbTyp, 1))
   sOldPart = Compress(cmbPrt)
   sNewPart = Compress(cmbNprt)
   sOldRev = Compress(cmbRev)
   sNewRev = Compress(cmbNrev)
   
   If sOldPart = sNewPart Then
      If sOldRev = sNewRev Then
         MsgBox "You Are Trying To Copy A List To Itself.", vbExclamation, Caption
         Exit Sub
      End If
   End If
   
   'In case the temp table remains
   On Error Resume Next
   sSql = "DROP TABLE #Dlst"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   On Error GoTo DiaErr1
   
   sMsg = "This Function Will Overwrite An Existing Revison (If One). " & vbCrLf _
          & "Do You Want To Continue To Copy List " & Trim(cmbPrt) & vbCrLf _
          & "To A New List For " & Trim(cmbNprt) & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      clsADOCon.BeginTrans
      
      sSql = "Delete FROM DlstTable WHERE DLSREF='" & sNewPart & "'" & vbCrLf _
         & "AND DLSREV='" & sNewRev & "' and DLSTYPE = " & iType
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "SELECT * INTO #Dlst From DlstTable" & vbCrLf _
         & "WHERE DLSREF = '" & sOldPart & "'" & vbCrLf _
         & "AND DLSREV = '" & sOldRev & "' AND DLSTYPE = " & iType
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "UPDATE #Dlst SET DLSREF = '" & sNewPart & "'," & vbCrLf _
         & "DLSREV = '" & sNewRev & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "INSERT INTO DlstTable SELECT * FROM #Dlst "
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      MouseCursor 0
      clsADOCon.CommitTrans
      MsgBox "The list was successfully copied.  " & clsADOCon.RowsAffected & " items were copied.", _
         vbInformation, Caption
      cmbPrt.Clear
      cmbNprt.Clear
      FillCombo
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   sProcName = "CopyDocumentList"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   MsgBox CurrError.Description & vbCrLf _
      & "Couldn't Copy The List.", vbExclamation, Caption
   Resume DiaErr2
DiaErr2:
   DoModuleErrors Me
   
End Sub


Private Sub GetOldDocRevisions()
   cmbRev.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT DLSREV FROM DlstTable WHERE " _
          & "DLSREF='" & Compress(cmbPrt) & "' "
   LoadComboBox cmbRev, -1
   If cmbRev.ListCount > 0 Then cmbRev = cmbRev.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "getolddocr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub
