VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHf05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open A Closed Manufacturing Order"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHf05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Qualified Part Numbers (CL)"
      Top             =   720
      Width           =   3545
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Contains Qualified Runs"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5040
      TabIndex        =   2
      ToolTipText     =   " Reopen thel MO"
      Top             =   480
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5040
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2235
      FormDesignWidth =   6075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Closed"
      Height          =   255
      Index           =   1
      Left            =   4000
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5040
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   765
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "ShopSHf05a"
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
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter
Dim bOnLoad As Byte
Dim bGoodRun As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Click()
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   GetRuns
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) Then
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
      GetRuns
   End If
   
End Sub


Private Sub cmbRun_Click()
   bGoodRun = GetCurrRun()
   
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   bGoodRun = GetCurrRun()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   cmbPrt = ""
   
End Sub


Private Sub cmdDel_Click()
   If bGoodRun = 0 Then
      MsgBox "Requires A Valid Run. See Help.", _
         vbExclamation, Caption
   Else
      OpenMO
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4154
      cmdHlp = False
      MouseCursor 0
   End If
   
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
   FormLoad Me, ES_DONTLIST, ES_RESIZE
   FormatControls
   
   sSql = "SELECT RUNREF,RUNNO FROM RunsTable WHERE " _
          & "RUNREF= ? AND RUNSTATUS='CL' "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   AdoQry.parameters.Append AdoParameter
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   
   Set ShopSHf05a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   Dim b As Byte
   sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   If b = 0 Then
      MsgBox "There Is No Open Inventory Journal For The Period.", _
         vbExclamation, Caption
      Sleep 500
      Unload Me
      Exit Sub
   End If
   cmbPrt.Clear
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF FROM " _
          & "RunsTable,PartTable WHERE PARTREF=RUNREF AND " _
          & "RUNSTATUS='CL' ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc, True)
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Sub

Private Sub GetRuns()
   Dim RdoRns As ADODB.Recordset
   cmbRun.Clear
   AdoQry.parameters(0).Value = Compress(cmbPrt)
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoQry, ES_FORWARD)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr cmbRun.hWnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
   End If
   Set RdoRns = Nothing
   If cmbRun.ListCount > 0 Then
      cmbRun = cmbRun.List(0)
      bGoodRun = GetCurrRun()
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetCurrRun() As Byte
   Dim RdoRun As ADODB.Recordset
   Dim lRunno As Long
   Dim sPart As String
   
   lRunno = Val(cmbRun)
   sPart = Compress(cmbPrt)
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,RUNCLOSED FROM RunsTable " _
          & "WHERE RUNREF='" & sPart & "' AND RUNNO=" & lRunno & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      lblStat = "" & Trim(RdoRun!RUNSTATUS)
      lblDte = Format(RdoRun!RUNCLOSED, "mm/dd/yy")
   Else
      lblStat = "**"
      lblDte = ""
   End If
   If lblStat = "CL" Then
      GetCurrRun = 1
   Else
      GetCurrRun = 0
      lblDte = ""
      lblStat = "**"
   End If
   Set RdoRun = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcurrrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub lblStat_Change()
   If lblStat = "**" Then
      lblStat.ForeColor = ES_RED
   Else
      lblStat.ForeColor = Es_TextForeColor
   End If
   
End Sub


Private Sub OpenMO()
   Dim rdoAct As ADODB.Recordset
   Dim iList As Integer
   Dim b As Byte
   Dim bResponse As Byte
   Dim iEntries As Integer
   Dim iTrans As Integer
   Dim iRef As Integer
   Dim lRunno As Long
   Dim sMsg As String
   Dim sPart As String
   Dim sRun As String
   Dim vEntry(100, 8) As Variant
   Dim lotNumber As String
   
   sPart = Compress(cmbPrt)
   lRunno = Val(cmbRun)
   sRun = Trim$(cmbRun)
   b = 5 - Len(sRun)
   
   'don't allow if inventory journal is closed
   'NOTE: this INGLPOSTED and INGLJOURNAL don't appear to ever be set.
   Dim rdo As ADODB.Recordset
   sSql = "select INGLJOURNAL, INGLPOSTED, INLOTNUMBER from InvaTable " & vbCrLf _
          & "where INPART='" & sPart & "' and INMORUN=" & sRun _
          & " and INTYPE=" & IATYPE_MoCompletion
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If bSqlRows Then
      With rdo
         lotNumber = Trim(!INLOTNUMBER)
         If !INGLPOSTED <> 0 Then
            MsgBox "Can't re-open MO.  Inventory Journal has been closed"
            Set rdo = Nothing
            Exit Sub
         End If
      End With
   End If
   Set rdo = Nothing
   
   'don't allow if lot is allocated
   sSql = "select LOTORIGINALQTY, LOTREMAININGQTY " & vbCrLf _
          & "from invatable join lohdtable on inlotnumber = lotnumber" & vbCrLf _
          & "and lotoriginalqty <> lotremainingqty" & vbCrLf _
          & "where INPART='" & sPart & "' and INMORUN=" & sRun _
          & " and INTYPE=" & IATYPE_MoCompletion
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      MsgBox "Can't re-open MO.  Part or all of the lot has been allocated."
      Set rdoAct = Nothing
      Exit Sub
   End If
   Set rdoAct = Nothing
   
   On Error GoTo DiaErr1
   sMsg = "This Reopens The MO To All Functions." & vbCr _
          & "Do You Really Want To Reopen This MO?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      If sJournalID <> "" Then iTrans = GetNextTransaction(sJournalID)
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      sSql = "UPDATE RunsTable SET RUNSTATUS='CO'," _
             & "RUNCLOSED=NULL " _
             & "WHERE RUNREF='" & sPart & "' AND " _
             & "RUNNO=" & lRunno & " "
      clsADOCon.ExecuteSQL sSql
      
'      'delete inventory activity record
'      sSql = "delete from  InvaTable " & vbCrLf _
'             & "where INPART='" & sPart & "' and INMORUN=" & sRun _
'             & " and INTYPE=" & IATYPE_MoCompletion
'      clsAdoCon.ExecuteSQL sSql
'
'      'delete lotitem record
'      If Len(lotNumber) > 0 Then
'         sSql = "delete from  LoitTable " & vbCrLf _
'                & "where LOINUMBER='" & lotNumber & "'"
'         clsAdoCon.ExecuteSQL sSql
'
'         'now delete lotheader record
'         sSql = "delete from  LohdTable " & vbCrLf _
'                & "where LOTNUMBER='" & lotNumber & "'"
'         clsAdoCon.ExecuteSQL sSql
'
'      End If
      
      clsADOCon.ExecuteSQL sSql
      If iTrans > 0 Then
         sSql = "SELECT DCHEAD,DCTRAN,DCREF,DCDEBIT," _
                & "DCCREDIT,DCACCTNO,DCDATE,DCDESC,DCPARTNO," _
                & "DCRUNNO FROM JritTable WHERE (DCPARTNO='" _
                & sPart & "' AND DCRUNNO=" & lRunno & " AND " _
                & "DCDESC='Close MO') ORDER BY DCREF DESC"
         bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
         If bSqlRows Then
            With rdoAct
               Do Until .EOF
                  iEntries = iEntries + 1
                  vEntry(iEntries, 0) = "" & Trim(sJournalID)
                  vEntry(iEntries, 1) = Format(!DCDEBIT, ES_QuantityDataFormat)
                  vEntry(iEntries, 2) = Format(!DCCREDIT, ES_QuantityDataFormat)
                  vEntry(iEntries, 3) = "" & Trim(!DCACCTNO)
                  vEntry(iEntries, 5) = "Reopen Closed MO"
                  vEntry(iEntries, 6) = "" & Trim(!DCPARTNO)
                  vEntry(iEntries, 7) = Format(!DCRUNNO, "#####0")
                  .MoveNext
               Loop
               ClearResultSet rdoAct
            End With
            Set rdoAct = Nothing
            'Just in case it is reclosed and reopened again
            If iEntries > 0 Then
               sMsg = "Close MO-" & Trim(str(iTrans))
               sSql = "Update JritTable SET DCDESC='" & sMsg & "' " _
                      & "WHERE DCPARTNO='" & sPart & "' AND DCRUNNO=" _
                      & lRunno & " AND DCDESC='Close MO'"
               clsADOCon.ExecuteSQL sSql
            End If
            For iList = 1 To iEntries
               'Reversing entries
               iRef = iRef + 1
               sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                      & "DCDEBIT,DCCREDIT,DCACCTNO,DCDATE,DCDESC,DCPARTNO," _
                      & "DCRUNNO) " _
                      & "VALUES('" _
                      & Trim(sJournalID) & "'," _
                      & iTrans & "," _
                      & iRef & "," _
                      & Val(vEntry(iList, 2)) & "," _
                      & Val(vEntry(iList, 1)) & ",'" _
                      & vEntry(iList, 3) & "','" _
                      & Format(ES_SYSDATE, "mm/dd/yy") & "','" _
                      & "Reopen Closed MO" & "','" _
                      & vEntry(iList, 6) & "'," _
                      & Val(vEntry(iList, 7)) & ")"
               clsADOCon.ExecuteSQL sSql
            Next
         End If
      End If
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         sMsg = "The Status Was Changed From CL To CO." & vbCr _
                & "Additional MO Actions Can Be Executed."
         MsgBox sMsg, vbInformation, Caption
         FillCombo
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "Couldn't Change The Run To Complete (CO).", vbExclamation, Caption
         GoTo DiaErr1
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "openmo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
