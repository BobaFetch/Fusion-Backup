VERSION 5.00
Begin VB.Form frmKeyInJob 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log on to Job Using Keyboard"
   ClientHeight    =   4920
   ClientLeft      =   4410
   ClientTop       =   4980
   ClientWidth     =   12060
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboFillPn 
      Height          =   420
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "dont auto select"
      Text            =   "cboFillPn"
      Top             =   3120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.ComboBox cboLot 
      Height          =   420
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "dont auto select"
      Text            =   "cboLotNum"
      Top             =   4200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton cmdSetupTime 
      BackColor       =   &H000000FF&
      Caption         =   "SetupTime"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CheckBox chkBarCodeReader 
      Caption         =   "Use Bar Code Reader"
      Height          =   300
      Left            =   5640
      TabIndex        =   11
      Top             =   3000
      Width           =   2775
   End
   Begin VB.ComboBox cboOp 
      Height          =   420
      Left            =   7020
      TabIndex        =   2
      Text            =   "cboOp"
      Top             =   900
      Width           =   4800
   End
   Begin VB.ComboBox cboMO 
      Height          =   420
      Left            =   5460
      TabIndex        =   1
      Text            =   "cboMO"
      Top             =   900
      Width           =   1500
   End
   Begin VB.ComboBox cboPart 
      Height          =   420
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "dont auto select"
      Text            =   "cboPart"
      Top             =   900
      Width           =   5295
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   2580
      Picture         =   "KeyInJob.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1500
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   4500
      Picture         =   "KeyInJob.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label lblFillPn 
      Caption         =   "Fill Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLotNum 
      Caption         =   "Lot Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Operation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7020
      TabIndex        =   10
      Top             =   540
      Width           =   1395
   End
   Begin VB.Label Label3 
      Caption         =   "MO Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5460
      TabIndex        =   9
      Top             =   540
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Part Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   540
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Log on to what job?"
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   6135
   End
End
Attribute VB_Name = "frmKeyInJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private keysTyped As Integer    ' number of keys typed in combo
'Private keysSoFar As String     ' the keys that were typed

Private bSetupTime As Boolean
Private bSetupEnabled As Boolean
Dim bAllowSC As Boolean
Dim denyLoginIfPriorOpOpen As Boolean

Private Sub SetupBarCodeReader()
    If chkBarCodeReader.Value = 1 Then
        cmdProceed.Default = False
    Else
        cmdProceed.Default = True
    End If
End Sub

Private Sub cboFillPn_Click()
   cboLot.Clear
   FillSelLot
End Sub

Private Sub cboFillPn_LostFocus()
   FillSelLot
   cboLot.SetFocus
End Sub


Private Sub cboLot_LostFocus()
   Dim i As Integer
   Dim strCurLot As String
   Dim bFound As Boolean
   bFound = False
   strCurLot = cboLot
   
   For i = 0 To cboLot.ListCount - 1
      If cboLot.List(i) = strCurLot Then
          bFound = True
          Exit For
      End If
   Next i
   
   If (bFound = False) Then
      MsgBox "Lot Number Not Found in the List"
      cboLot = ""
   End If
   

End Sub

Private Sub cboMO_Click()
    If (cboMO.Text <> "") Then
        LoadPomComboWithOpenOpsForRun cboOp, cboPart.Text, CLng(cboMO.Text)
    End If
End Sub

Private Sub cboMO_GotFocus()
   'keysTyped = 0
   'keysSoFar = ""
   If chkBarCodeReader.Value = 0 Then ComboGotFocus cboMO
End Sub

Private Sub cboMO_KeyPress(KeyAscii As Integer)
    If chkBarCodeReader.Value = 1 And KeyAscii = 13 Then
        cboOp.SetFocus
    End If

   If chkBarCodeReader.Value = 0 Then KeyAscii = 0 'avoid auto-selection of first character
End Sub

Private Sub cboMO_KeyUp(KeyCode As Integer, Shift As Integer)
   If chkBarCodeReader.Value = 0 Then ComboKeyUp cboMO, KeyCode
End Sub

Private Sub cboMO_LostFocus()
    Dim i As Integer
    Dim bFound As Boolean
    bFound = False
    If (IsNumeric(cboMO.Text)) Then
      cboMO.Text = Val(cboMO.Text)
    End If
    For i = 0 To cboMO.ListCount - 1
        If cboMO.Text = cboMO.List(i) Then
            bFound = True
            Exit For
        End If
    Next i
    If Not bFound Then cboPart.SetFocus
    If (cboMO.Text <> "") Then
        LoadPomComboWithOpenOpsForRun cboOp, cboPart.Text, CLng(cboMO.Text)
    End If

End Sub

Private Sub cboOp_GotFocus()
   'keysTyped = 0
   If chkBarCodeReader.Value = 0 Then ComboGotFocus cboOp
End Sub

Private Sub cboOp_KeyPress(KeyAscii As Integer)
    If chkBarCodeReader.Value = 1 And KeyAscii = 13 Then
      If (EnabledSelLot = 1) Then
         cboLot.SetFocus
      Else
         cmdProceed.SetFocus
      End If
    End If

   If chkBarCodeReader.Value = 0 Then KeyAscii = 0 'avoid auto-selection of first character
End Sub

Private Sub cboLot_KeyPress(KeyAscii As Integer)
    If chkBarCodeReader.Value = 1 And KeyAscii = 13 Then
      cmdProceed.SetFocus
    End If
   If chkBarCodeReader.Value = 0 Then KeyAscii = 0 'avoid auto-selection of first character
End Sub

Private Sub cboOp_KeyUp(KeyCode As Integer, Shift As Integer)
   If chkBarCodeReader.Value = 0 Then ComboKeyUp cboOp, KeyCode
End Sub

Private Sub cboOp_LostFocus()
    Dim i As Integer
    Dim bFound As Boolean
    bFound = False
    'If cboOp.ListIndex = -1 Then Exit Sub
    
    'bar code reader can exit combobox without selecting valid field when entering op followed by carriage return
    If chkBarCodeReader.Value = 1 And cboOp.ListIndex = -1 Then
         For i = 0 To cboOp.ListCount - 1
             If cboOp.Text = cboOp.List(i) Then
                 cboOp.ListIndex = i
                 bFound = True
                 Exit For
             End If
         Next i
    ElseIf cboOp.ListIndex <> -1 Then
         bFound = True
    End If
    
    If cboOp.ListIndex = -1 Then Exit Sub

    If EnabledSelLot = 1 Then SelectPartByRtOp
    If Not bFound Then cboPart.SetFocus
    
End Sub

Private Sub GetSelLot()

    Dim strRunPart As String
    Dim strRunNo As String
    Dim strRunOp As String
    strRunPart = Compress(cboPart)
    strRunNo = cboMO
    strRunOp = cboOp.ItemData(cboOp.ListIndex)
   
    Dim RdoPart As ADODB.Recordset
     
    sSql = "SELECT lotpartref, ISNULL(b.lotuserlotid, '') as lotuserlotid FROM RnopTable a,lohdTable b WHERE OPREF = '" & strRunPart & "'" _
            & " AND OPRUN = " & strRunNo & " AND OPNO = " & strRunOp _
            & " AND a.lotuserlotid = b.lotuserlotid"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoPart)
    If bSqlRows Then
        With RdoPart
          cboLot = Trim(!lotuserlotid)
          cboFillPn = Trim(!lotpartref)
      End With
    Else
        cboLot = ""
    End If
    ClearResultSet RdoPart
    Set RdoPart = Nothing
End Sub

Private Sub SelectPartByRtOp()

   Dim strMOPart As String
   Dim strOpNo As String
   
   strMOPart = Compress(cboPart)
   'strOpNo = Compress(cboOp)
   If cboOp.ListIndex = -1 Then
      strOpNo = ""
   Else
      strOpNo = Compress(cboOp.ItemData(cboOp.ListIndex))
   End If
      
   If (strMOPart <> "" And strOpNo <> "") Then
   
      sSql = "select OPFILLREF from rtopTable where opref = '" & strMOPart & "' and opno = " & strOpNo

      'cboLot.Clear
      'cboFillPn.Clear
      
      Dim rs As ADODB.Recordset
      Set rs = clsADOCon.GetRecordSet(sSql, ES_STATIC)
      If Not rs.BOF And Not rs.EOF Then
         With rs
            cboFillPn = "" & Trim(.Fields(0))
            ' Allow
            cboFillPn.Enabled = False
            ClearResultSet rs
         End With
         FillSelLot
      Else
         cboFillPn.Enabled = True
         cboFillPn = ""
      End If
      Set rs = Nothing
      ' Fill lots only with the selected parts
      
   End If
   
End Sub

Private Sub SelectPartByBmpl()

   Dim strMOPart As String
   
   strMOPart = Compress(cboPart)
   
   If (strMOPart <> "") Then
      sSql = "SELECT DISTINCT LOTPARTREF FROM lohdTable, LtTrkTable, bmplTable " _
               & " WHERE lohdTable.Lotnumber = LtTrkTable.LOTNUMBER AND " _
            & " BMassyPart = '" & strMOPart & "' AND BMPARTREF = LOTPARTREF"

      cboLot.Clear
      cboFillPn.Clear
      
      Dim rs As ADODB.Recordset
      Set rs = clsADOCon.GetRecordSet(sSql, ES_STATIC)
      If Not rs.BOF And Not rs.EOF Then
         With rs
            Do Until .EOF
               AddComboStr cboFillPn.hwnd, "" & Trim(.Fields(0))
               .MoveNext
            Loop
            ' Allow
            'cboFillPn.Enabled = False
            ClearResultSet rs
         End With
         FillSelLot
      Else
         cboFillPn.Enabled = True
         cboFillPn = ""
      End If
      Set rs = Nothing
      ' Fill lots only with the selected parts
      
   End If
   
   
End Sub

      

Private Sub FillPart()
   sSql = "SELECT DISTINCT LOTPARTREF FROM lohdTable, LtTrkTable " _
         & " WHERE lohdTable.Lotnumber = LtTrkTable.LOTNUMBER"
   Dim rs As ADODB.Recordset
   cboFillPn.Clear
   cboLot.Clear
   AddComboStr cboFillPn.hwnd, ""
   Set rs = clsADOCon.GetRecordSet(sSql, ES_STATIC)
   If Not rs.BOF And Not rs.EOF Then
      With rs
         Do Until .EOF
            AddComboStr cboFillPn.hwnd, "" & Trim(.Fields(0))
            .MoveNext
         Loop
         ClearResultSet rs
      End With
   End If
   If cboFillPn.ListCount <> 0 Then
      bSqlRows = 1
      cboFillPn.ListIndex = 0
   Else
      bSqlRows = 0
   End If
   sSql = ""
   Set rs = Nothing
   
   
End Sub


Private Sub FillSelLot()

   Dim strFillPn As String
   
   strFillPn = cboFillPn
   If (strFillPn = "") Then
      sSql = "SELECT DISTINCT LtTrkTable.lotuserlotid FROM LohdTable, LtTrkTable " _
               & " WHERE LohdTable.Lotnumber = LtTrkTable.Lotnumber"
   Else
      
      sSql = "SELECT DISTINCT LtTrkTable.lotuserlotid FROM LohdTable, LtTrkTable " _
               & " WHERE LohdTable.Lotnumber = LtTrkTable.Lotnumber AND LOTPARTREF = '" _
               & strFillPn & "'"
   End If
   
   Dim rs As ADODB.Recordset
   cboLot.Clear
   AddComboStr cboLot.hwnd, ""
   Set rs = clsADOCon.GetRecordSet(sSql, ES_STATIC)
   
   If (Not rs Is Nothing) Then
      If Not rs.BOF And Not rs.EOF Then
         With rs
            Do Until .EOF
               AddComboStr cboLot.hwnd, "" & Trim(.Fields(0))
               .MoveNext
            Loop
            ClearResultSet rs
         End With
      End If
   End If
   
   If cboLot.ListCount <> 0 Then
      bSqlRows = 1
      cboLot.ListIndex = 0
   Else
      bSqlRows = 0
   End If
   sSql = ""
   Set rs = Nothing
   
   
End Sub

Private Sub cboPart_Click()
   cboOp.Clear
   LoadComboWithOpenRunsForPart cboMO, cboPart.Text, bAllowSC
End Sub

Private Sub cboPart_GotFocus()
   'keysTyped = 0
   If chkBarCodeReader.Value = 0 Then ComboGotFocus cboPart
End Sub

Private Sub cboPart_KeyPress(KeyAscii As Integer)
    If chkBarCodeReader.Value = 1 And KeyAscii = 13 Then
        cboMO.SetFocus
    End If

   If chkBarCodeReader.Value = 0 Then KeyAscii = 0 'avoid auto-selection of first character
End Sub

Private Sub cboPart_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If chkBarCodeReader.Value = 0 Then ComboKeyUp cboPart, KeyCode
End Sub

Private Sub cboPart_LostFocus()
    Dim i As Integer
    Dim bFound As Boolean
    bFound = False
    For i = 0 To cboPart.ListCount - 1
        If cboPart.Text = cboPart.List(i) Then
            bFound = True
            Exit For
        End If
    Next i
    If Not bFound Then cboPart.SetFocus
    LoadComboWithOpenRunsForPart cboMO, cboPart.Text, bAllowSC
    
    If (EnabledSelLot = 1) Then
      SelectPartByBmpl
      'SelectSelLotByPart
    End If
End Sub

Private Sub chkBarCodeReader_Click()
    SaveSetting "Esi2000", "ESIPOM", "UseBarCodeReader", CStr(chkBarCodeReader.Value)
    SetupBarCodeReader
End Sub

Private Sub cmdBack_Click()
   Unload Me
End Sub

Public Function UseOpDescription() As Boolean
   Dim rs As ADODB.Recordset
   sSql = "select CoUseNamesInPOM from ComnTable where CoUseNamesInPOM = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rs)
   UseOpDescription = bSqlRows
End Function


Private Sub cmdProceed_Click()
   'log user into operation
   If (cboMO.Text = "") Then
      MsgBox ("Please select MO number")
      Exit Sub
   End If
   
   
   'first make sure we're not already logged in to that operation
'   Dim opNo As String
'   opNo = cboOp.ItemData(cboOp.ListIndex)
'   If LoggedInToThisJob(cboPart.Text, cboMO.Text, opNo) Then
'      Exit Sub
'   End If
   
   'check for invalid op entered by scanner
   Dim opNo As Integer
   If cboOp.ListIndex = -1 Then
      If cboOp.Text <> "" Then
         If GetShowOpenOpOnly() Then
            MsgBox "Op " & cboOp.Text & " is not a valid open operation."
         Else
            MsgBox "Op " & cboOp.Text & " is not a valid operation."
         End If
      End If
      cboOp.SetFocus
      Exit Sub
   End If
   
   opNo = cboOp.ItemData(cboOp.ListIndex)
   'If LoggedInToThisJob(cboPart.Text, cboMO.Text, cboOp.Text) Then
   If Not LoggedInToThisJob(cboPart.Text, cboMO.Text, opNo) Then
   
'   Dim rdo As ADODB.Recordset
'   sSql = "select WCNNUM, WCNSHOP" & vbCrLf _
'          & "from WcntTable wc join RnopTable op on op.OPCENTER = wc.WCNNUM" & vbCrLf _
'          & "where OPREF='" & cboPart.Text & "'" & vbCrLf _
'          & "and OPRUN=" & cboMO.Text & vbCrLf _
'          & "and OPNO=" & cboOp.Text
'   If clsADOCon.GetDataSet(sSql, rdo) Then
'      mempCurrentEmployee.strCurWC = Trim(rdo!WCNNUM)
'      mempCurrentEmployee.strCurShop = Trim(rdo!WCNSHOP)
'   End If
'   Set rdo = Nothing
'

      Dim rdo As ADODB.Recordset
      sSql = "select WCNNUM, WCNSHOP" & vbCrLf
      
      If denyLoginIfPriorOpOpen Then
         sSql = sSql & ",ISNULL((select OPCOMPLETE from RnopTable prev where prev.OPREF = op.OPREF and prev.OPRUN = op.OPRUN" & vbCrLf _
            & "and prev.OPNO = (select max(opno) from RnopTable ops3 where ops3.OPREF = op.OPREF" & vbCrLf _
            & "and ops3.OPRUN = op.OPRUN and ops3.OPNO < op.OPNO)),1) as OKTOLOGIN" & vbCrLf
      Else
         sSql = sSql & ",1 as OKTOLOGIN" & vbCrLf
      End If
         
      sSql = sSql & "from WcntTable wc join RnopTable op on op.OPCENTER = wc.WCNNUM" & vbCrLf _
             & "where OPREF='" & Compress(cboPart.Text) & "'" & vbCrLf _
             & "and OPRUN=" & cboMO.Text & vbCrLf _
             & "and OPNO=" & opNo   ' cboOp.Text
      If clsADOCon.GetDataSet(sSql, rdo) Then
         If rdo!OKTOLOGIN = 1 Then
            mempCurrentEmployee.strCurWC = Trim(rdo!WCNNUM)
            mempCurrentEmployee.strCurShop = Trim(rdo!WCNSHOP)
         Else
            MsgBox "You cannot log into this operation until the prior operation is complete"
            Exit Sub
         End If
      End If
      Set rdo = Nothing
   End If   '@@@???

   If (EnabledSelLot = 1) Then
      If (Trim(cboLot) = "" Or Trim(cboFillPn) = "") Then
         MsgBox ("Please select Fill Number and Lot number")
         Exit Sub
      End If
   End If

   'log in to the operation specified
   'LogInToJob Compress(cboPart.Text), cboMO.Text, cboOp.Text, bSetupTime
   LogInToJob Compress(cboPart.Text), cboMO.Text, opNo, bSetupTime
   ' Add the Selected Lotnumber
   If (EnabledSelLot = 1) Then UpdateLotNumber
   Unload Me
End Sub

Private Sub UpdateLotNumber()

    Dim strRunPart As String
    Dim strRunNo As String
    Dim strRunOp As String
    Dim strLot As String
    
    strRunPart = Compress(cboPart)
    strRunNo = cboMO
    strRunOp = cboOp
    strLot = cboLot
   
    sSql = "UPDATE RnopTable SET lotuserlotid = '" & strLot & "' " _
            & " WHERE OPREF = '" & strRunPart & "'" _
            & " AND OPRUN = " & strRunNo & " AND OPNO = " & strRunOp
    clsADOCon.ExecuteSql sSql

End Sub
Private Sub cmdSetupTime_Click()
   If (bSetupTime = True) Then
      bSetupTime = False
      cmdSetupTime.BackColor = &HFF&
   Else
      bSetupTime = True
      cmdSetupTime.BackColor = &HFF00&
   End If
   
End Sub

Public Function EnabledAllowSC()
   Dim rdo As ADODB.Recordset
   Dim companyAccount As String
   
   sSql = "select ISNULL(COALLOWSCPOM, 0) as COALLOWSCPOM from ComnTable"
   If clsADOCon.GetDataSet(sSql, rdo) Then
      EnabledAllowSC = rdo!COALLOWSCPOM
      rdo.Close
   Else
      EnabledAllowSC = 0
   End If
   
End Function

Public Function EnabledSelLot()
   Dim rdo As ADODB.Recordset
   Dim companyAccount As String
   
   sSql = "select ISNULL(COLOTATPOM, 0) as COLOTATPOM from ComnTable"
   If clsADOCon.GetDataSet(sSql, rdo) Then
      EnabledSelLot = rdo!COLOTATPOM
      rdo.Close
   Else
      EnabledSelLot = 0
   End If
   
End Function


Private Sub Form_Load()
   CenterForm Me
   
   bAllowSC = IIf((EnabledAllowSC = 1), True, False)
   LoadComboWithPartsForOpenRuns cboPart, bAllowSC
   chkBarCodeReader = GetSetting("Esi2000", "ESIPOM", "UseBarCodeReader", "0")

   ' get the flag from system setting
   cmdSetupTime.Visible = IIf((GetSetupTimeEnabled = 1), True, False)
   bSetupTime = False
   cmdSetupTime.Default = False
   
   If (EnabledSelLot = 1) Then
    lblLotNum.Visible = True
    cboLot.Visible = True
    lblFillPn.Visible = True
    cboFillPn.Visible = True
    
    FillPart
    FillSelLot
   End If
   
   SetupBarCodeReader
   
   'get deny login flag
    Dim rs As ADODB.Recordset
    Dim deny As Integer
    
    sSql = "SELECT isnull(DenyLoginIfPriorOpOpen,0) as DenyLoginIfPriorOpOpen FROM ComnTable WHERE COREF=1"
    bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
    If bSqlRows Then
        denyLoginIfPriorOpOpen = IIf(rs!denyLoginIfPriorOpOpen = 1, True, False)
    End If
    Set rs = Nothing

End Sub

Public Sub LoadPomComboWithOpenOpsForRun(cbo As ComboBox, sPart As String, nRun As Long)
   'load the open ops for a run into a combobox
   
   Dim rdo As ADODB.Recordset
   Dim bTCSerOp As Boolean
   Dim bOnlyOpenOp As Boolean
   cbo.Clear
   
   Dim UseDescription As String
   UseDescription = UseOpDescription()
   
   bOnlyOpenOp = GetShowOpenOpOnly()
   bTCSerOp = GetTCServiceOp()
   sSql = "Select OPNO, rtrim(WCNDESC) as WCNDESC from RnopTable " & vbCrLf _
      & "join WcntTable on OPCENTER = WCNREF" & vbCrLf _
      & "WHERE OPREF = '" & Compress(sPart) & "'" & vbCrLf _
      & "AND OPRUN = " & nRun & vbCrLf
   If (bTCSerOp = True) Then
        If (bOnlyOpenOp = True) Then
            sSql = sSql & "AND OPCOMPLETE = 0 " & vbCrLf
        End If
    Else
         ' MM 6/1/2010
        If (bOnlyOpenOp = True) Then
            sSql = sSql & "AND (LTRIM(RTRIM(OPSERVPART)) = '' OR OPSERVPART IS NULL) " & vbCrLf _
               & "AND OPCOMPLETE = 0 " & vbCrLf
        Else
            sSql = sSql & "AND (LTRIM(RTRIM(OPSERVPART)) = '' OR OPSERVPART IS NULL) " & vbCrLf
        End If
    End If
    
    If UseDescription Then
      sSql = sSql & "ORDER BY WCNDESC"
   Else
      sSql = sSql & "ORDER BY OPNO"
   End If
   
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
         While Not .EOF
            If UseDescription Then
               cbo.AddItem !WCNDESC
            Else
               cbo.AddItem CStr(!opNo)
            End If
            
            cbo.ItemData(cbo.NewIndex) = CStr(!opNo)
            .MoveNext
         Wend
      End With
      If cbo.ListCount > 0 Then
         cbo.ListIndex = 0
      End If
      Set rdo = Nothing
      
   Else
      MsgBox "No open operations for this MO", vbExclamation ', sSysCaption
   End If
   
End Sub



