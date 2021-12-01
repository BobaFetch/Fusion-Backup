VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form BompBMf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change A Parts List Revision"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMf03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "BompBMf03a.frx":07AE
      Height          =   320
      Left            =   4680
      Picture         =   "BompBMf03a.frx":1120
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Parts List for Part and Revision"
      Top             =   1080
      Width           =   350
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3000
      FormDesignWidth =   6780
   End
   Begin VB.TextBox txtRev 
      Height          =   285
      Left            =   5760
      TabIndex        =   2
      ToolTipText     =   "New Revision"
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdChg 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5760
      TabIndex        =   6
      ToolTipText     =   "Change The Revision"
      Top             =   600
      Width           =   915
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5760
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Revision-Select From List"
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox cmbPls 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   1080
      Width           =   3405
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5760
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   13
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Effective"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Obsolete"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label txtObs 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   2520
      Width           =   950
   End
   Begin VB.Label txtEff 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   2160
      Width           =   950
   End
   Begin VB.Label txtRef 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev:"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parts List"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "BompBMf03a"
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
'Dim RdoBmh As ADODB.Recordset
Dim AdoCmdObj As ADODB.Command
Dim RdoBmh As ADODB.Recordset

Dim bGoodPart As Byte
Dim bGoodOList As Byte
Dim bGoodNList As Byte
Dim bOnLoad As Byte

Dim sPartNumber As String
Dim sPartBomrev As String

Private txtKeyPress() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbPls_Click()
   GetList
   
End Sub


Private Function GetPartsList() As Byte
   cmbRev = Compress(cmbRev)
   sPartNumber = Compress(cmbPls)
   On Error Resume Next
   RdoBmh.Close
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREF,BMHREV,BMHREFERENCE,BMHOBSOLETE,BMHREVDATE," _
          & "BMHEFFECTIVE FROM BmhdTable WHERE BMHREF='" & sPartNumber & "' " _
          & "AND BMHREV='" & Trim(cmbRev) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBmh)
   If bSqlRows Then
      With RdoBmh
         txtRef = "" & Trim(!BMHREFERENCE)
         txtEff = "" & Format(!BMHEFFECTIVE, "mm/dd/yyyy")
         txtObs = "" & Format(!BMHOBSOLETE, "mm/dd/yyyy")
      End With
      cmdChg.Enabled = True
      GetPartsList = True
   Else
      cmdChg.Enabled = False
      txtRef = ""
      txtEff = ""
      txtObs = ""
      GetPartsList = False
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getpartsl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbPls_LostFocus()
   cmbPls = CheckLen(cmbPls, 30)
   GetList
   
End Sub


Private Sub cmbRev_Click()
   bGoodOList = GetPartsList()
   
End Sub


Private Sub cmbRev_LostFocus()
   cmbRev = CheckLen(cmbRev, 4)
   bGoodOList = GetPartsList()
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdChg_Click()
   Dim bAssigned As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sOldRevision As String
   Dim sNewRevision As String
   
   If Not bGoodPart And Not bGoodOList Then Exit Sub
   cmbRev = Compress(cmbRev)
   sOldRevision = cmbRev
   sNewRevision = Compress(txtRev)
   sMsg = "Change Parts List " & sOldRevision & vbCrLf _
          & "Part Number " & cmbPls & " To " & sNewRevision & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      clsADOCon.BeginTrans
      sSql = "UPDATE BmhdTable SET BMHREV='" & sNewRevision & "' " _
             & "WHERE BMHREF='" & sPartNumber & "' AND BMHREV='" _
             & sOldRevision & "' "
      clsADOCon.ExecuteSQL sSql ' rdExecDirect
      
      sSql = "UPDATE BmplTable SET BMREV='" & sNewRevision & "' " _
             & "WHERE BMASSYPART='" & sPartNumber & "' AND BMREV='" _
             & sOldRevision & "' "
      clsADOCon.ExecuteSQL sSql ' rdExecDirect
      clsADOCon.CommitTrans
      If sPartBomrev = cmbRev Then
         sSql = "UPDATE PartTable SET PABOMREV='" & sNewRevision & "' " _
                & "WHERE PARTREF='" & sPartNumber & "' "
         clsADOCon.ExecuteSQL sSql ' rdExecDirect
      End If
      
   Else
      CancelTrans
      Exit Sub
   End If
   On Error Resume Next
   bGoodOList = GetPartsList()
   SysMsg "Parts List Revision Changed.", True, Me
   txtRev = ""
   MouseCursor 0
   Exit Sub
   
BcrevDn1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume BcrevDn2
BcrevDn2:
   On Error Resume Next
   clsADOCon.RollbackTrans
   MouseCursor 0
   MsgBox Err.Description & vbCrLf _
      & "Could Not Change The Parts List Revision.", vbExclamation, Caption
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3252
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdVew_Click()
   If cmdVew Then
      ViewBomTree.Show
      cmdVew = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillPartsBelow4 cmbPls
      If cmbPls.ListCount > 0 Then cmbPls = cmbPls.List(0)
      GetList
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub GetList()
   Dim RdoGet As ADODB.Recordset
   cmbRev.Clear
   cmbRev = ""
   sPartNumber = Compress(cmbPls)
   On Error GoTo DiaErr1
   AdoCmdObj.Parameters(0) = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoGet, AdoCmdObj)
   If bSqlRows Then
      With RdoGet
         lblDsc = "" & Trim(!PADESC)
         cmbRev = "" & Trim(!PABOMREV)
         sPartBomrev = "" & Trim(!PABOMREV)
         ClearResultSet RdoGet
      End With
      cUR.CurrentPart = cmbPls
      bGoodPart = True
   Else
      lblDsc = ""
      txtRef = ""
      txtEff = ""
      txtObs = ""
      sPartBomrev = ""
      MsgBox "Part Wasn't Found or Is The Wrong Type.", vbExclamation, Caption
      bGoodPart = False
   End If
   If bGoodPart Then
      FillBomhRev sPartNumber
      bGoodOList = GetPartsList()
   End If
   Set RdoGet = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getlist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PABOMREV FROM " _
          & "PartTable WHERE PARTREF= ? AND PALEVEL<4"
          
   Set AdoCmdObj = New ADODB.Command
   AdoCmdObj.CommandText = sSql
   
   Dim prmPrtRef As ADODB.Parameter
   Set prmPrtRef = New ADODB.Parameter
   prmPrtRef.Type = adChar
   prmPrtRef.Size = 30
   AdoCmdObj.Parameters.Append prmPrtRef
   
   'Set RdoPrt = RdoCon.CreatePreparedStatement("", sSql)
   ' TODO: RdoPrt.MaxRows = 1
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   SaveCurrentSelections
   FormUnload
   Set AdoCmdObj = Nothing
   Set RdoBmh = Nothing
   Set BompBMf03a = Nothing
   
End Sub


Private Sub txtRev_LostFocus()
   txtRev = CheckLen(txtRev, 4)
   If txtRev = "" Then
      cmdChg.Enabled = False
      Exit Sub
   Else
      bGoodNList = GetNewList()
      If Not bGoodNList Then MsgBox "That Revision Is Already Listed.", vbExclamation, Caption
   End If
   
End Sub



Private Function GetNewList() As Byte
   txtRev = Compress(txtRev)
   sPartNumber = Compress(cmbPls)
   On Error Resume Next
   RdoBmh.Close
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREF,BMHREV,BMHREFERENCE,BMHOBSOLETE,BMHREVDATE," _
          & "BMHEFFECTIVE FROM BmhdTable WHERE BMHREF='" & sPartNumber & "' " _
          & "AND BMHREV='" & Trim(txtRev) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBmh)
   If bSqlRows Then
      txtRev = ""
      cmdChg.Enabled = False
      GetNewList = False
   Else
      cmdChg.Enabled = True
      txtRef = ""
      txtEff = ""
      txtObs = ""
      GetNewList = True
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getnewl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
