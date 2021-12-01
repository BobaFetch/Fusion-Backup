VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form BompBMe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parts List"
   ClientHeight    =   3255
   ClientLeft      =   1620
   ClientTop       =   750
   ClientWidth     =   6990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z2 
      Height          =   30
      Left            =   180
      TabIndex        =   28
      Top             =   1320
      Width           =   6732
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optRefresh 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   2160
      TabIndex        =   26
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "BompBMe02a.frx":07AE
      Height          =   320
      Left            =   5040
      Picture         =   "BompBMe02a.frx":0C4C
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Parts List for Part and Revision"
      Top             =   600
      Width           =   350
   End
   Begin VB.CommandButton cmdRel 
      Caption         =   "&Release"
      Height          =   315
      Left            =   6000
      TabIndex        =   18
      ToolTipText     =   "Release (Unrelease) To Production"
      Top             =   2110
      Width           =   925
   End
   Begin VB.ComboBox txtObs 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Tag             =   "4"
      Top             =   2520
      Width           =   1250
   End
   Begin VB.ComboBox txtEff 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Tag             =   "4"
      Top             =   2160
      Width           =   1250
   End
   Begin VB.CommandButton cmdDte 
      Caption         =   "E&ffectivity"
      Height          =   315
      Left            =   6000
      TabIndex        =   20
      ToolTipText     =   "Set Effectivity Dates For This Revision"
      Top             =   2450
      Visible         =   0   'False
      Width           =   925
   End
   Begin VB.CommandButton cmdPhn 
      Caption         =   "&Assign"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Set This Parts List as Default for Part"
      Top             =   1770
      Width           =   925
   End
   Begin VB.CheckBox optPls 
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdPls 
      Caption         =   "&Parts List"
      Height          =   285
      Left            =   6000
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Show Parts List"
      Top             =   600
      Width           =   925
   End
   Begin VB.TextBox txtRef 
      Height          =   285
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Parts List Reference"
      Top             =   1800
      Width           =   1275
   End
   Begin VB.ComboBox cmbRev 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Revision (Blank For Default)"
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox cmbPls 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   600
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   925
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4680
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3255
      FormDesignWidth =   6990
   End
   Begin VB.Label txtRev 
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label txtPls 
      Height          =   255
      Left            =   5520
      TabIndex        =   24
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblRel 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3840
      TabIndex        =   22
      ToolTipText     =   "Released To Production Or Not Released"
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Released"
      Height          =   255
      Index           =   7
      Left            =   3000
      TabIndex        =   21
      ToolTipText     =   "Released To Production Or Not Released"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "No Current Parts List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Index           =   6
      Left            =   4920
      TabIndex        =   16
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5640
      TabIndex        =   15
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   12
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Obsolete"
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Effective"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Description"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "BompBMe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/7/05 Added updates to BMHREVDATE column (all changes)
'9/4/06 Removed Cur.CurrPart
Option Explicit
'Dim RdoPrt As rdoQuery
Dim RdoBmh As ADODB.Recordset
Dim AdoCmdObj As ADODB.Command

Dim bCancel As Byte
Dim bGoodPart As Byte
Dim bGoodList As Byte
Dim bNewHeader As Byte
Dim bOnLoad As Byte

Dim sOldPart As String
Dim sPartNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd

Private Function GetPartType(sPartNumber As String) As Integer
   Dim RdoType As ADODB.Recordset
   sSql = "SELECT PARTREF,PALEVEL FROM PartTable WHERE " _
          & "PARTREF='" & Compress(sPartNumber) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoType, ES_FORWARD)
   If bSqlRows Then
      With RdoType
         If Not IsNull(.Fields(1)) Then
            GetPartType = .Fields(1)
         Else
            GetPartType = 0
         End If
         ClearResultSet RdoType
      End With
   End If
   Set RdoType = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getparttype"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Function GetPartsList(onload As Byte) As Byte
   cmbRev = Compress(cmbRev)
   sPartNumber = Compress(cmbPls)
   On Error Resume Next
   RdoBmh.Close
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREF,BMHREV,BMHREFERENCE,BMHOBSOLETE,BMHREVDATE," _
          & "BMHEFFECTIVE,BMHRELEASED FROM BmhdTable WHERE BMHREF='" & sPartNumber & "' " _
          & "AND BMHREV='" & Trim(cmbRev) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBmh, ES_KEYSET)
   If bSqlRows Then
      With RdoBmh
         GetPartsList = 1
         txtRef = "" & Trim(!BMHREFERENCE)
         txtEff = "" & Format(!BMHEFFECTIVE, "mm/dd/yyyy")
         txtObs = "" & Format(!BMHOBSOLETE, "mm/dd/yyyy")
         If Not onload Then cmdPhn.Enabled = True
         If txtEff = "" Then txtEff = Format(ES_SYSDATE, "mm/dd/yyyy")
         If txtObs = "" Then txtObs = "00/00/0000"
         cmdPls.Enabled = True
         If txtRef = "" Then txtRef = "NONE"
         If !BMHRELEASED Then lblRel = "Y" Else lblRel = "N"
      End With
      lblEdit = "Editing Parts List."
   Else
      txtRef = ""
      txtEff = ""
      txtObs = ""
      cmdPls.Enabled = False
      cmdPhn.Enabled = False
      GetPartsList = 0
      lblEdit = "No Current Parts List."
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getpartsl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbPls_Click()
   If sOldPart <> cmbPls Then GetList
   
End Sub


Private Sub cmbPls_LostFocus()
   cmbPls = CheckLen(cmbPls, 30)
   
   If (Trim(cmbPls.Text) <> "" And Not ValidPartNumber(cmbPls.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPls = ""
      Exit Sub
   End If

   If bCancel = 0 Then
      If sOldPart <> cmbPls Then GetList
      sOldPart = cmbPls
   End If
   
End Sub


Private Sub cmbRev_Change()
   If Len(cmbRev) > 4 Then cmbRev = Left(cmbRev, 4)
   
End Sub

Private Sub cmbRev_Click()
   bGoodList = GetPartsList(False)
   
End Sub


Private Sub cmbRev_LostFocus()
   cmbRev = CheckLen(cmbRev, 4)
   bGoodList = GetPartsList(False)
   If Len(Trim(cmbRev)) > 0 Then _
          If bGoodList = 0 Then AddPartsList
   
End Sub


Private Sub cmdCan_Click()
   Dim b As Byte
   'did they forget something?
   For b = 0 To Forms.Count - 1
      If Forms(b).Name = "BompBM02b" Then Unload Forms(b)
   Next
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdDte_Click()
   'This button will only Be visible for testing
   BompBMe02c.Show vbModal
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3202
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdPhn_Click()
   Dim bResponse As Byte
   Dim sCurrPart As String
   Dim sMsg As String
   
   sMsg = "Assign As Default Parts List For" & vbCrLf _
          & cmbPls & "?"
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
   txtRef.SetFocus
   
End Sub

Private Sub cmdPls_Click()
   bGoodList = GetPartsList(False)
   If bGoodList Then
      On Error Resume Next
      MouseCursor 13
      RdoBmh!BMHREVDATE = Format(ES_SYSDATE, "mm/dd/yy")
      txtPls = cmbPls
      txtRev = cmbRev
      cmdPls.Enabled = True
      optPls.Value = vbChecked
      BompBMe02b.Show
   End If
   
End Sub



Private Sub cmdRel_Click()
   Dim iList As Integer
   Dim sPartNumber As String
   Dim sPartRev As String
   
   sPartNumber = Compress(cmbPls)
   sPartRev = Compress(cmbRev)
   On Error GoTo DiaErr1
   If lblRel = "N" Then iList = 1 Else iList = 0
   If iList = 1 Then
      sSql = "UPDATE BmhdTable SET BMHRELEASED=1,BMHRELEASEDATE='" _
             & Format(Now, "mm/dd/yy") & "' WHERE BMHREF='" _
             & sPartNumber & "' AND BMHREV='" & sPartRev & "' "
   Else
      sSql = "UPDATE BmhdTable SET BMHRELEASED=2,BMHRELEASEDATE=" _
             & "Null WHERE BMHREF='" _
             & sPartNumber & "' AND BMHREV='" & sPartRev & "' "
   End If
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   Sleep 500
   If clsADOCon.RowsAffected Then
      If lblRel = "N" Then
         lblRel = "Y"
         SysMsg "Parts Listed Was Released.", True, Me
      Else
         lblRel = "N"
         SysMsg "Parts Listed Was Unreleased.", True, Me
      End If
   Else
      MsgBox "Couldn't Find The Parts List Revision Record.", vbExclamation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "cmdrel_click"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdVew_Click()
   If cmdVew Then
      ViewBomTree.Show
      'Form1.Show
      cmdVew = False
   End If
   
End Sub

Private Sub Form_Activate()
   Dim iList As Integer
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      sOldPart = ""
      FillPartsBelow4 cmbPls
      If cmbPls.ListCount > 0 Then cmbPls = cmbPls.List(0)
      GetList
      bOnLoad = 0
   End If
   If optPls.Value = vbChecked Then
      Unload BompBMe02b
      optPls.Value = vbUnchecked
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PABOMREV FROM " _
          & "PartTable WHERE PARTREF= ? AND PALEVEL<4"
   
   
   Set AdoCmdObj = New ADODB.Command
   AdoCmdObj.CommandText = sSql
   
   Dim prmPtrRef As ADODB.Parameter
   Set prmPtrRef = New ADODB.Parameter
   prmPtrRef.Type = adChar
   prmPtrRef.Size = 30
   AdoCmdObj.Parameters.Append prmPtrRef
   
   'Set RdoPrt = RdoCon.CreatePreparedStatement("", sSql)
   ' TODO: MM not sure if we need that...
   ' TODO: We can set the MAX record in Record set
   'RdoPrt.MaxRows = 1
   
   
   'cUR.CurrentPart = GetSetting("Esi2000", "Current", "Part", cUR.CurrentPart)
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If optPls.Value = vbChecked Then Unload BompBMe02b
   SaveCurrentSelections
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   On Error Resume Next
   'RdoPrt.Close
   Set AdoCmdObj = Nothing
   Set RdoBmh = Nothing
   Set BompBMe02a = Nothing
End Sub




Private Sub GetList()
   Dim RdoPls As ADODB.Recordset
   cmbRev.Clear
   cmbRev = ""
   sPartNumber = Compress(cmbPls)
   cmdPhn.Enabled = False
   On Error GoTo DiaErr1
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
      cmdPhn.Enabled = False
      cmdPls.Enabled = True
      cUR.CurrentPart = cmbPls
      bGoodPart = 1
   Else
      cmdPls.Enabled = False
      lblDsc = ""
      lblLvl = ""
      txtRef = ""
      txtEff = ""
      txtObs = ""
      If sPartNumber <> "" Then
          MsgBox "Part Wasn't Found or Is The Wrong Type.", vbExclamation, Caption
      End If
      bGoodPart = 0
   End If
   If bGoodPart Then
      On Error Resume Next
      FillBomhRev sPartNumber
      bGoodList = GetPartsList(False)
   End If
   Set RdoPls = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getlist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub





Private Sub lblRel_Change()
   If lblRel = "Y" Then cmdRel.Caption = "Un&release" Else _
               cmdRel.Caption = "&Release"
   
End Sub

Private Sub optPls_Click()
   'checked only if loaded items
   'Never visible
   
End Sub

Private Sub txtEff_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtEff_LostFocus()
   On Error Resume Next
   If Trim(txtEff) = "" Then
      If bGoodList Then
         'RdoBmh.Edit
         RdoBmh!BMHREVDATE = Format(ES_SYSDATE, "mm/dd/yyyy")
         RdoBmh!BMHEFFECTIVE = Null
         RdoBmh.Update
         If Err > 0 Then ValidateEdit
      End If
   Else
      txtEff = CheckDateEx(txtEff)
      If bGoodList Then
         'RdoBmh.Edit
         RdoBmh!BMHREVDATE = Format(ES_SYSDATE, "mm/dd/yyyy")
         RdoBmh!BMHEFFECTIVE = Format(txtEff, "mm/dd/yyyy")
         RdoBmh.Update
         If Err > 0 Then ValidateEdit
      End If
   End If
   
End Sub

Private Sub txtObs_Click()
   If txtObs = "00/00/00" Then txtObs = ""
   
End Sub

Private Sub txtObs_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtObs_LostFocus()
   txtObs = Trim(txtObs)
   If txtObs = "" Then txtObs = "00/00/0000"
   If txtObs = "00/00/0000" Then Exit Sub
   If Len(txtObs) > 0 Then
      txtObs = CheckDateEx(txtObs)
      If bGoodList Then
         'RdoBmh.Edit
         RdoBmh!BMHOBSOLETE = txtObs
         RdoBmh.Update
         If Err > 0 Then ValidateEdit
      End If
   Else
      If bGoodList Then
         'RdoBmh.Edit
         RdoBmh!BMHOBSOLETE = Null
         RdoBmh.Update
         If Err > 0 Then ValidateEdit
      End If
   End If
   
End Sub

Private Sub txtRef_GotFocus()
   SelectFormat Me
   If bNewHeader Then
      bGoodList = GetPartsList(False)
      bNewHeader = False
   End If
   If bGoodList Then
      cmdPls.Enabled = True
      cmdPhn.Enabled = True
   End If
   
End Sub


Private Sub txtRef_LostFocus()
   txtRef = CheckLen(txtRef, 10)
   On Error Resume Next
   If txtRef = "NONE" Then Exit Sub
   If bGoodList Then
      On Error Resume Next
      'RdoBmh.Edit
      RdoBmh!BMHREFERENCE = txtRef
      RdoBmh!BMHREVDATE = Format(ES_SYSDATE, "mm/dd/yyyy")
      RdoBmh.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub AddPartsList()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sRevision As String
   
   On Error GoTo DiaErr1
   If sPartNumber = "" Then Exit Sub
   sMsg = "Parts List Revision " & Trim(cmbRev) & " Wasn't Found." & vbCrLf _
          & "Add Parts List?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      cmbRev = Compress(cmbRev)
      sRevision = cmbRev
      sSql = "INSERT INTO BmhdTable (BMHREF,BMHPARTNO,BMHREV) " _
             & "VALUES('" & sPartNumber & "','" & cmbPls & "','" & sRevision & "')"
      clsADOCon.ExecuteSql sSql ' rdExecDirect
      If clsADOCon.RowsAffected > 0 Then
         SysMsg "Revision Added.", True, Me
         AddComboStr cmbRev.hwnd, sRevision
         bGoodList = GetPartsList(False)
         On Error Resume Next
         txtRef.SetFocus
         BompBMe02c.Show vbModal
         bNewHeader = True
      Else
         MsgBox "Couldn't Add Parts List.", vbExclamation, Caption
      End If
   Else
      CancelTrans
      cmbRev = ""
      On Error Resume Next
      cmbRev.SetFocus
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addpartslist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
