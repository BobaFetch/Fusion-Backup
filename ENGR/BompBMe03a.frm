VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form BompBMe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign A Parts List To A Part"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   3203
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMe03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "BompBMe03a.frx":07AE
      Height          =   320
      Left            =   4680
      Picture         =   "BompBMe03a.frx":1120
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Parts List for Part and Revision"
      Top             =   1080
      Width           =   350
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5400
      Top             =   2520
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2985
      FormDesignWidth =   5895
   End
   Begin VB.CommandButton cmdAsn 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4920
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Set This Parts List as Default for Part"
      Top             =   600
      Width           =   915
   End
   Begin VB.ComboBox cmbPls 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   1080
      Width           =   3345
   End
   Begin VB.ComboBox cmbRev 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Revision-Select From List"
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Set As Default Routing"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   2412
   End
   Begin VB.Label txtObs 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   2520
      Width           =   1000
   End
   Begin VB.Label txtEff 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   2160
      Width           =   1000
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Effective"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Obsolete"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblLvl 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5040
      TabIndex        =   7
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Index           =   6
      Left            =   4440
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1820
      Width           =   1335
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   3135
   End
End
Attribute VB_Name = "BompBMe03a"
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
Dim RdoBmh As ADODB.Recordset

Dim bGoodPart As Byte
Dim bGoodList As Byte
Dim bOnLoad As Byte

Dim sPartNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbPls_Click()
   GetList
   
End Sub


Private Sub cmbPls_LostFocus()
   cmbPls = CheckLen(cmbPls, 30)
   
   If (Not ValidPartNumber(cmbPls.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPls = ""
      Exit Sub
   End If
   
   If Len(cmbPls) Then GetList
   
End Sub


Private Sub cmbRev_Click()
   bGoodList = GetPartsList()
   
End Sub


Private Sub cmbRev_LostFocus()
   cmbRev = CheckLen(cmbRev, 4)
   bGoodList = GetPartsList()
   If bGoodList = 0 Then MsgBox "That PL Revision Wasn't Found.", _
                  vbInformation, Caption
   
End Sub


Private Sub cmdAsn_Click()
   Dim bResponse As Byte
   Dim sCurrPart As String
   Dim sMsg As String
   sMsg = "Assign " & cmbRev & " As Default Parts List" & vbCrLf _
          & "For " & cmbPls & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      sCurrPart = Compress(cmbPls)
      sSql = "UPDATE PartTable SET PABOMREV='" & cmbRev & "' " _
             & "WHERE PARTREF='" & sCurrPart & "' "
      clsADOCon.ExecuteSql sSql ' rdExecDirect
      sSql = "UPDATE BmhdTable SET BMHPART='" & sCurrPart & "' " _
             & "WHERE BMHREF='" & sCurrPart & "' AND BMHREV='" _
             & cmbRev & "' "
      clsADOCon.ExecuteSql sSql ' rdExecDirect
      If clsADOCon.RowsAffected > 0 Then
         SysMsg "Default Set.", True, Me
      Else
         MsgBox "Couldn't Assign.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   On Error Resume Next
   cmbPls.SetFocus
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbPls = ""
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3203
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
      FillParts
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub GetList()
   Dim RdoPls As ADODB.Recordset
   cmbRev.Clear
   cmbRev = ""
   sPartNumber = Compress(cmbPls)
   On Error GoTo DiaErr1
   'RdoPrt(0) = sPartNumber
   AdoCmdObj.Parameters(0) = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoPls, AdoCmdObj)
   If bSqlRows Then
      With RdoPls
         cmbPls = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblLvl = "" & Trim(!PALEVEL)
         cmbRev = "" & Trim(!PABOMREV)
         ClearResultSet RdoPls
      End With
      cUR.CurrentPart = cmbPls
      cmdAsn.Enabled = True
      bGoodPart = True
   Else
      lblDsc = ""
      lblLvl = ""
      txtEff = ""
      txtObs = ""
      cmdAsn.Enabled = False
      If Not bOnLoad Then
         MsgBox "Part Wasn't Found or Is The Wrong Type.", vbExclamation, Caption
      Else
         If cmbPls.ListCount > 0 Then cmbPls = cmbPls.List(0)
      End If
      bGoodPart = False
   End If
   If bGoodPart Then
      FillBomhRev sPartNumber
      bGoodList = GetPartsList()
   End If
   Set RdoPls = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getlist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Function GetPartsList() As Byte
   cmbRev = Compress(cmbRev)
   sPartNumber = Compress(cmbPls)
   On Error Resume Next
   RdoBmh.Close
   If Trim(cmbRev) = "" Then
      GetPartsList = 1
   Else
      On Error GoTo DiaErr1
      sSql = "SELECT BMHREF,BMHREV,BMHREFERENCE,BMHOBSOLETE,BMHREVDATE," _
             & "BMHEFFECTIVE FROM BmhdTable WHERE BMHREF='" & sPartNumber & "' " _
             & "AND BMHREV='" & Trim(cmbRev) & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBmh, ES_FORWARD)
      If bSqlRows Then
         With RdoBmh
            txtEff = "" & Format(!BMHEFFECTIVE, "mm/dd/yyyy")
            txtObs = "" & Format(!BMHOBSOLETE, "mm/dd/yyyy")
            GetPartsList = 1
            ClearResultSet RdoBmh
         End With
      Else
         txtEff = ""
         txtObs = ""
         GetPartsList = 0
      End If
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getpartsl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PABOMREV FROM " _
          & "PartTable WHERE PARTREF= ? AND PALEVEL<4"
          
   Set AdoCmdObj = New ADODB.Command
   AdoCmdObj.CommandText = sSql
   
   Dim pPrtRef As ADODB.Parameter
   Set pPrtRef = New ADODB.Parameter
   pPrtRef.Type = adChar
   pPrtRef.Size = 30
   AdoCmdObj.Parameters.Append pPrtRef

'   Set RdoPrt = RdoCon.CreatePreparedStatement("", sSql)
' TODO: use RS
'   RdoPrt.MaxRows = 1
   bOnLoad = 1
   
End Sub


Private Sub FillParts()
   FillPartsBelow4 cmbPls
   On Error GoTo DiaErr1
   'If cUR.CurrentPart <> "" Then cmbPls = cUR.CurrentPart
   If cmbPls.ListCount > 0 Then cmbPls = cmbPls.List(0)
   GetList
   Exit Sub
   
DiaErr1:
   sProcName = "fillparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   SaveCurrentSelections
   On Error Resume Next
   Set AdoCmdObj = Nothing
   Set RdoBmh = Nothing
   FormUnload
   Set BompBMe03a = Nothing
   
End Sub

