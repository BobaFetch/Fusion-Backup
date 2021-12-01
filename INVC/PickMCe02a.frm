VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form PickMCe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pick Substitution"
   ClientHeight    =   3765
   ClientLeft      =   1620
   ClientTop       =   855
   ClientWidth     =   6795
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z2 
      Height          =   30
      Left            =   120
      TabIndex        =   24
      Top             =   1480
      Width           =   6612
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PickMCe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optLot 
      Caption         =   "Lot Tracked Part"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1320
      TabIndex        =   21
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5820
      TabIndex        =   5
      ToolTipText     =   "Apply This Substitution To The Pick List"
      Top             =   3360
      Width           =   915
   End
   Begin VB.TextBox txtNqt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ComboBox cmbNew 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Susbstitute Part"
      Top             =   2640
      Width           =   3495
   End
   Begin VB.ComboBox cmbOld 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Select Existing Part On PL From List"
      Top             =   1680
      Width           =   3495
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Contains Manufacturing Orders With Unpicked Items"
      Top             =   720
      Width           =   3545
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5640
      TabIndex        =   1
      ToolTipText     =   "Contains Manufacturing Orders With Unpicked Items"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5800
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3765
      FormDesignWidth =   6795
   End
   Begin VB.Label lblQOH 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5040
      TabIndex        =   25
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label PickRecord 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   600
      TabIndex        =   22
      ToolTipText     =   "Pick Row"
      Top             =   2040
      Width           =   372
   End
   Begin VB.Label lblLvl 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   20
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   19
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblNum 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   18
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblOum 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   17
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblQty 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5040
      TabIndex        =   16
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblNds 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label lblOds 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Substitute Part"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Part"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   10
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status (PL,PP)"
      Height          =   255
      Index           =   15
      Left            =   4920
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "PickMCe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Corrected cmbNew fill
'9/1/04 omit tools
'2/16/05 Changed Combos to reflect only unpicked
'4/26/06 Added PickRow and Pick warnings for Update (Garret)
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim AdoPrt As ADODB.Command
Dim ADOParameter2 As ADODB.Parameter

Dim bOnLoad As Byte
Dim bGoodRuns As Byte
Dim bGoodMo As Byte
Dim bGoodPart As Byte
Dim bGoodPick As Boolean
'Dim bGoodQty  As Byte

Dim iOldLevel As Integer
Dim vItems(1000, 5) As Variant
Const PICKITEM_PartDescription = 0
Const PICKITEM_PartUnits = 1
Const PICKITEM_PickQuantity = 2
Const PICKITEM_PartType = 3
Const PICKITEM_RecordNo = 4

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtNqt = "0.000"
   
End Sub

Private Sub cmbNew_Click()
   GetNewPart
   
End Sub

Private Sub cmbNew_GotFocus()
   txtNqt.Enabled = True
   cmbNew_Click
   
End Sub


Private Sub cmbNew_LostFocus()
   cmbNew = CheckLen(cmbNew, 30)
   
   If (Not ValidPartNumber(cmbNew.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbNew = ""
      Exit Sub
   End If
   
   GetNewPart
   
End Sub


Private Sub cmbOld_Click()
   On Error Resume Next
   lblOds = vItems(cmbOld.ListIndex, PICKITEM_PartDescription)
   lblOum = vItems(cmbOld.ListIndex, PICKITEM_PartUnits)
   lblQty = vItems(cmbOld.ListIndex, PICKITEM_PickQuantity)
   lblLvl = vItems(cmbOld.ListIndex, PICKITEM_PartType)
   PickRecord = vItems(cmbOld.ListIndex, PICKITEM_RecordNo)
   FillNewParts
   
End Sub


Private Sub cmbOld_GotFocus()
   cmbOld_Click
   
End Sub


Private Sub cmbOld_LostFocus()
   On Error Resume Next
   lblOds = vItems(cmbOld.ListIndex, PICKITEM_PartDescription)
   lblOum = vItems(cmbOld.ListIndex, PICKITEM_PartUnits)
   lblQty = vItems(cmbOld.ListIndex, PICKITEM_PickQuantity)
   lblLvl = vItems(cmbOld.ListIndex, PICKITEM_PartType)
   PickRecord = vItems(cmbOld.ListIndex, PICKITEM_RecordNo)
   txtNqt = lblQty
   
End Sub


Private Sub cmbPrt_Click()
   bGoodRuns = GetRuns()
   
End Sub


Private Sub cmbPrt_GotFocus()
   cmbPrt_Click
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   bGoodRuns = GetRuns()
   
End Sub


Private Sub cmbRun_Click()
   bGoodPick = GetPicks(False)
   
End Sub

Private Sub cmbRun_GotFocus()
   cmbRun_Click
   
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   bGoodPick = GetPicks(True)
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5202"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSub_Click()
   Dim bResponse As Byte
   Dim bLotItem As Byte
   Dim iList As Integer
   
   Dim sMsg As String
   Dim sNewPart As String
   Dim sOldPart As String
   
   If Trim(PickRecord) = "" Then
      MsgBox "Invalid Pick Row. Please Re-Select.", _
         vbInformation, Caption
      Exit Sub
   End If
   If cmbNew = cmbOld Then
      MsgBox "The Old And New Parts Are The Same.", _
         vbInformation, Caption
      On Error Resume Next
      cmbNew.SetFocus
      Exit Sub
   End If
   'check it out
   If Not IsNumeric(txtNqt.Text) Then
      On Error Resume Next
      txtNqt.SetFocus
      MsgBox "There Must Be A Valid Quantity.", vbInformation, Caption
      Exit Sub
   End If
   
   If Val(txtNqt) <= 0 Then
      On Error Resume Next
      txtNqt.SetFocus
      MsgBox "There Must Be A Valid Quantity.", vbInformation, Caption
      Exit Sub
   End If
   
   If Val(txtNqt) <> Val(lblQty) Then
      sMsg = "The Old And New Quantities Differ. " & vbCr _
             & "Please Verify the Substitution."
      bResponse = MsgBox(sMsg, 1 + 32 + 256, Caption)
      If bResponse <> 1 Then
         On Error Resume Next
         cmbNew.SetFocus
         Exit Sub
      End If
   End If
   'must be ok..
'   sMsg = "Please Confirm The Substitution Of The " & vbCr _
'          & "New Part Number For The Existing Item." & vbCr _
'          & "Note: The Part Number Will Not Be Picked."
'   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
'   If bResponse = vbNo Then
'      On Error Resume Next
'      cmbNew.SetFocus
'      Exit Sub
'   End If
   cmdSub.Enabled = False
   MouseCursor 13
   iList = cmbOld.ListIndex
   If iList < 0 Then iList = 0
'   MsgBox "Please Note: This Function Replaces The Original," & vbCr _
'      & "But Does Not Auto Pick The Substitute Part Number.", _
'      vbExclamation, Caption
   sNewPart = Compress(cmbNew)
   sSql = "UPDATE MopkTable SET PKPARTREF='" & sNewPart & "'," _
          & "PKTYPE=23,PKPQTY=" & Trim(str(Val(txtNqt))) & ",PKUNITS='" _
          & lblNum & "'  WHERE (PKMOPART='" & Compress(cmbPrt) & "' AND " _
          & "PKMORUN=" & cmbRun & " AND PKRECORD=" & Val(PickRecord) & ")"
   clsADOCon.ExecuteSql sSql
   MouseCursor 0
   If clsADOCon.RowsAffected > 0 Then
      MsgBox "The Substitution Update Was Successful " & vbCr _
         & "And Is Now Ready To Be Picked.", vbInformation, Caption
   Else
      MsgBox "Couldn't Complete The Substitution.", vbExclamation, Caption
   End If
   On Error Resume Next
   cmbRun.SetFocus
   
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALEVEL,PATOOL " _
          & "FROM PartTable WHERE (PALEVEL = ? AND PATOOL=0)" _
          & "ORDER BY PARTREF"
   Set AdoPrt = New ADODB.Command
   AdoPrt.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adInteger
   AdoPrt.Parameters.Append AdoParameter1
   
   sSql = "SELECT DISTINCT RUNREF,RUNSTATUS,RUNNO,PKPARTREF,PKMOPART,PKMORUN," _
          & "PKAQTY FROM RunsTable,MopkTable WHERE (RUNREF=PKMOPART AND RUNREF = ? " _
          & "AND RUNNO=PKMORUN AND RUNSTATUS NOT LIKE 'C%' " _
          & "AND PKAQTY=0)"
   
'   Set RdoQry = RdoCon.CreateQuery("", sSql)
    Set AdoQry = New ADODB.Command
    AdoQry.CommandText = sSql
    
    Set ADOParameter2 = New ADODB.Parameter
    ADOParameter2.Type = adChar
    ADOParameter2.Size = 30
    
    AdoQry.Parameters.Append ADOParameter2
    
    
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   
   'this was a bug.  You can't set hundreds of thousands of records to the same value.
   'no one is complaining about a problem with the logic, so just eliminate this call.
   'sSql = "UPDATE MopkTable SET PKRECORD=0 WHERE PKRECORD>0"
   'clsADOCon.ExecuteSql sSql
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter1 = Nothing
   Set ADOParameter2 = Nothing
   Set AdoQry = Nothing
   Set AdoPrt = Nothing
   Set PickMCe02a = Nothing
   
End Sub



Private Sub FillCombo()
   'Dim RdoCmb As ADODB.Recordset
   Dim b As Byte
   
   On Error GoTo DiaErr1
   sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   If b = 0 Then
      MsgBox "There Is No Open Inventory Journal For This Period.", _
         vbExclamation, Caption
      Sleep 500
      Unload Me
      Exit Sub
   End If
   sProcName = "fillcombo"
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PKMOPART " _
          & "FROM PartTable,MopkTable WHERE (PARTREF=PKMOPART AND " _
          & "PKAQTY=0) ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetRuns() As Byte
   Dim RdoRns As ADODB.Recordset
   Dim iList As Integer
   
   ClearBoxes
   cmbRun.Clear
   On Error GoTo DiaErr1
   'RdoQry(0) = Compress(cmbPrt)
   AdoQry.Parameters(0).Value = Compress(cmbPrt)
   On Error GoTo 0
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoQry)
   If bSqlRows Then
      With RdoRns
         cmbRun = Format(!Runno, "####0")
         lblStat = "" & !RUNSTATUS
         Do Until .EOF
            If iList <> !Runno Then
               AddComboStr cmbRun.hWnd, Format$(!Runno, "####0")
               iList = !Runno
            End If
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
      GetRuns = True
   Else
      GetRuns = False
   End If
   Set RdoRns = Nothing
   If GetRuns Then bGoodMo = GetPart()
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetPart() As Byte
   Dim RdoRun As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "Qry_GetPartsNotTools '" & Compress(cmbPrt) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         ClearResultSet RdoRun
      End With
      GetPart = 1
   Else
      GetPart = 0
   End If
   Set RdoRun = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub ClearBoxes()
   bGoodPart = False
   cmbOld.Clear
   cmbNew.Clear
   cmbNew.Enabled = False
   cmdSub.Enabled = False
   txtNqt.Enabled = False
   lblOds = ""
   lblNds = ""
   lblQOH = ""
   lblQty = ""
   lblOum = ""
   lblNum = ""
   txtNqt = ""
   PickRecord = ""
   
End Sub

Private Function GetPicks(FinalSet As Byte) As Boolean
   Dim RdoPck As ADODB.Recordset
   Dim iList As Integer
   
   cmbOld.Clear
   cmbNew.Clear
   cmbNew.Enabled = True
   iOldLevel = 99
   Erase vItems
   iList = -1
   
   ' don't try to update a clustered index!
   ' use the numbers that are there
   '    If FinalSet Then
   '        On Error Resume Next
   '        'THIS WAS A BUG -- should compare on PKMOPART, NOT PKPARTREF
   '        'IT MOSTLY DIDN'T DO ANYTHING, UNLESS THERE WAS AN MO THE PICKED PART WITH THE RUN SELECTED
   '        sSql = "UPDATE MopkTable SET PKRECORD=0 WHERE PKRECORD>0 " _
   '            & "AND PKPARTREF='" & Compress(cmbPrt) & "' AND PKMORUN=" & cmbRun & " "
   '        RdoCon.Execute sSql, rdExecDirect
   '    End If
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL," _
          & "PAUNITS,PKPARTREF,PKMOPART,PKMORUN,PKREV,PKPQTY," _
          & "PKAQTY,PKRECORD FROM PartTable,MopkTable WHERE PARTREF=PKPARTREF " _
          & "AND PKMOPART='" & Compress(cmbPrt) & "' AND (PKMORUN=" _
          & Trim(Val(cmbRun)) & " AND PKAQTY=0) ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPck, ES_KEYSET)
   If bSqlRows Then
      GetPicks = True
      'On Error Resume Next
      With RdoPck
         cmbOld = "" & Trim(!PartNum)
         lblOds = "" & Trim(!PADESC)
         lblOum = "" & Trim(!PAUNITS)
         lblQty = Format(0 + !PKPQTY, ES_QuantityDataFormat)
         lblLvl = "" & !PALEVEL
         If Err = 40009 Then
            lblOds = ""
            lblOum = ""
            lblQty = "0.000"
            lblLvl = ""
         End If
         Do Until .EOF
            iList = iList + 1
            AddComboStr cmbOld.hWnd, "" & Trim(!PartNum)
            vItems(iList, PICKITEM_PartDescription) = "" & Trim(!PADESC)
            vItems(iList, PICKITEM_PartUnits) = "" & Trim(!PAUNITS)
            vItems(iList, PICKITEM_PickQuantity) = Format(0 + !PKPQTY, ES_QuantityDataFormat)
            vItems(iList, PICKITEM_PartType) = "" & Trim(!PALEVEL)
            'If FinalSet Then
            '                .Edit
            '                !PKRECORD = iList + 1
            'Debug.Print "(" & iList & ") PKRECORD=" & !PKRECORD & " " & !PARTNUM
            '                .Update
            'End If
            
            '                If !PKRECORD <> iList + 1 Then
            '                    Debug.Print "not equal"
            '                End If
            
            'vItems(iList, PICKITEM_RecordNo) = iList + 1
            'JUST USE THE EXISTING NUMBERS
            vItems(iList, PICKITEM_RecordNo) = !PKRECORD
            
            .MoveNext
         Loop
         PickRecord = vItems(0, PICKITEM_RecordNo)
         ClearResultSet RdoPck
      End With
   Else
      cmbOld = ""
      lblOds = ""
      lblOum = ""
      lblQty = Format(0, ES_QuantityDataFormat)
      lblLvl = ""
      GetPicks = False
   End If
   Set RdoPck = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpicks"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FillNewParts()
   Dim RdoSub As ADODB.Recordset
   'don't run it if the part type hasn't changed
   txtNqt = ""
   'bGoodQty = False
   'If iOldLevel = Val(Trim(lblLvl)) Then Exit Sub
   'iOldLevel = Val(Trim(lblLvl))
   
   cmbNew.Clear
   On Error GoTo DiaErr1
   'RdoPrt(0) = Val(lblLvl)
   AdoPrt.Parameters(0).Value = Val(lblLvl)
   bSqlRows = clsADOCon.GetQuerySet(RdoSub, AdoPrt)
   If bSqlRows Then
      With RdoSub
         cmdSub.Enabled = True
         Do Until .EOF
            If Trim(!PartNum) <> cmbPrt Then _
                    AddComboStr cmbNew.hWnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoSub
      End With
      If cmbNew.ListCount > 0 Then cmbNew = cmbNew.List(0)
   Else
      cmdSub.Enabled = False
   End If
   Set RdoSub = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillnewPa"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetNewPart()
   Dim RdoNew As ADODB.Recordset
   Dim sGetPart
   sGetPart = Compress(cmbNew)
   
   On Error Resume Next
   If Len(sGetPart) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALEVEL,PAQOH," _
             & "PALOTTRACK,PATOOL FROM PartTable WHERE (PARTREF='" & sGetPart & "' " _
             & "AND PALEVEL=" & lblLvl & " AND PATOOL=0)"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoNew)
      If bSqlRows Then
         With RdoNew
            cmbNew = "" & Trim(!PartNum)
            lblNds = "" & Trim(!PADESC)
            lblQOH = !PAQOH
            lblNum = "" & Trim(!PAUNITS)
            optLot.Value = !PALOTTRACK
            ClearResultSet RdoNew
         End With
         bGoodPart = True
      Else
         MsgBox "Part Wasn't Found Or Is Wrong Type.", vbExclamation, Caption
         cmbNew = ""
         bGoodPart = False
      End If
      On Error Resume Next
   Else
      cmbNew = ""
   End If
   Set RdoNew = Nothing
   
End Sub

Private Sub txtNqt_LostFocus()
   txtNqt = CheckLen(txtNqt, 9)
   txtNqt = Format(Abs(Val(txtNqt)), ES_QuantityDataFormat)
   'If Val(txtNqt) > 0 Then bGoodQty = 1 Else bGoodQty = 0
   
End Sub
