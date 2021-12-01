VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHf12a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adjust OH Rate For Open Manufacturing Order"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPOH 
      Height          =   285
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   16
      Tag             =   "1"
      ToolTipText     =   "New Po Number"
      Top             =   4200
      Width           =   855
   End
   Begin VB.ComboBox cmbOP 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "Contains Qualified Runs"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtFOH 
      Height          =   285
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "New Po Number"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHf12a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Qualified Part Numbers (CO)"
      Top             =   960
      Width           =   3545
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Contains Qualified Runs"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdOHCost 
      Caption         =   "Update OH"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      ToolTipText     =   " Update OH Cost for Open MO"
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   4320
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4800
      FormDesignWidth =   7200
   End
   Begin VB.Label lblCurPOH 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   19
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current OH Percent Rate"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   18
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New OH Percent Rate"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   17
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run OP"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   14
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblStat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3600
      TabIndex        =   13
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New OH Fixed Rate"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current OH Fixed Rate"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblCurFOH 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
End
Attribute VB_Name = "ShopSHf12a"
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
'Dim rdoQry As rdoQuery
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim bCantClose As Byte
Dim bOnLoad As Byte
Dim bGoodPrt As Byte
Dim bGoodRun As Byte

Dim lRunno As Long
Dim sPartNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd



Private Sub cmbPrt_Click()
   bGoodPrt = GetRunPart()
   GetRuns
   'GetOHCost
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) Then
      bGoodPrt = GetRunPart()
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
   PopulateRunOP
End Sub

Private Sub cmbOP_Click()
   GetOverheadRate
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdUpdOHCost_Click()
   
   Dim RdoEmpRt As ADODB.Recordset
   Dim strPart As String
   Dim lngRun As Long
   Dim iOp As Integer
   Dim iOHFRate As Currency
   Dim iOHPRate As Currency
   Dim iCalOH As Currency
   
   Dim Rate As Currency
   Dim empno As String
   
   On Error GoTo DiaErr1
   strPart = cmbPrt
   lngRun = Val(cmbRun)
   iOp = Val(cmbOP)
   
   If Trim(txtFOH.Text) = "" Then
      iOHFRate = 0
   Else
      iOHFRate = Val(txtFOH.Text)
   End If
   
   If Trim(txtPOH.Text) = "" Then
      iOHPRate = 0
   Else
      iOHPRate = Val(txtPOH.Text)
   End If
   
   ' Get emp rate
   sSql = "SELECT PREMPAYRATE,TCEMP from EmplTable,tcitTable where PREMNUMBER = TCEMP AND " _
            & " TCPARTREF = '" & Compress(cmbPrt) & "' AND TCRUNNO = " & lngRun _
            & " AND TCOPNO = " & iOp

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmpRt, ES_FORWARD)
   If bSqlRows Then
      With RdoEmpRt
         Do Until .EOF
            Rate = !PREMPAYRATE
            empno = !TCEMP
            iCalOH = (iOHPRate * Rate) / 100
              
            sSql = "UPDATE tcitTable SET TCOHFIXED = " & iOHFRate _
                        & ", TCOHRATE = " & iCalOH _
                     & " FROM tcitTable WHERE TCPARTREF = '" _
                        & Compress(cmbPrt) & "' AND TCRUNNO = " & cmbRun _
                        & " AND TCOPNO = " & iOp _
                        & " AND TCEMP = '" & empno & "'"
            
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
            .MoveNext
         Loop
   
      End With
   End If
   
   Set RdoEmpRt = Nothing

   Exit Sub
   
DiaErr1:
   sProcName = "cmdUpdOHCost"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub Form_Activate()
   Dim b As Byte
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sSql = "SELECT RUNREF,RUNNO FROM RunsTable WHERE " _
          & "RUNREF= ? AND RUNSTATUS NOT LIKE 'C%'"
'   Set rdoQry = RdoCon.CreateQuery("", sSql)
   
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   
   AdoQry.Parameters.Append AdoParameter
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   
   FormUnload
   Set ShopSHf12a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbPrt.Clear
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF FROM " _
          & "RunsTable,PartTable WHERE PARTREF=RUNREF AND " _
          & "RUNSTATUS NOT LIKE 'C%' ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      bGoodPrt = GetRunPart()
      GetRuns
      PopulateRunOP
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetRuns()
   Dim RdoRns As ADODB.Recordset
   cmbRun.Clear
   sPartNumber = Compress(cmbPrt)
   
   AdoQry.Parameters(0).Value = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoQry, ES_KEYSET)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr cmbRun.hwnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
   End If
   Set RdoRns = Nothing
   If cmbRun.ListCount > 0 Then
      cmbRun = cmbRun.List(0)
      If GetPreferenceValue("AutoSelectLastRun") = "1" Then cmbRun = cmbRun.List(cmbRun.ListCount - 1)
      bGoodRun = GetCurrRun()
      PopulateRunOP
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
   
   lRunno = Val(cmbRun)
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS FROM RunsTable " _
          & "WHERE RUNREF='" & Compress(cmbPrt) & "' AND RUNNO=" & lRunno & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         lblStat = "" & Trim(!RUNSTATUS)
         ClearResultSet RdoRun
      End With
   Else
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
   If Left(lblDsc, 10) = "*** Part N" Then
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

Private Function GetRunPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   cmbRun.Clear
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PAPRODCODE," _
          & "PASTDCOST FROM PartTable WHERE PARTREF='" & Compress(cmbPrt) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         ClearResultSet RdoPrt
         GetRunPart = 1
      End With
   Else
      GetRunPart = 0
      lblDsc = "*** Part Number Wasn't Found ****"
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrunpart"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub PopulateRunOP()
   'load the open ops for a run into a combobox
   
   Dim sPart As String
   Dim nRun As Long
   
   Dim rdo As ADODB.Recordset
   Dim bTCSerOp As Boolean
   cmbOP.Clear
   
   sPart = cmbPrt
   nRun = cmbRun
   
   bTCSerOp = GetTCServiceOp()
    If (bTCSerOp = True) Then
        sSql = "Select OPNO from RnopTable " & vbCrLf _
               & " WHERE OPREF = '" & Compress(sPart) & "'" & vbCrLf _
               & " AND OPRUN = " & nRun & vbCrLf _
               & " ORDER BY OPNO"
    Else
      sSql = "SELECT OPNO FROM RnopTable " & vbCrLf _
               & " WHERE OPREF = '" & Compress(sPart) & "'" & vbCrLf _
               & " AND OPRUN = " & nRun & vbCrLf _
               & " AND (LTRIM(RTRIM(OPSERVPART)) = '' OR OPSERVPART IS NULL) " & vbCrLf _
               & " ORDER BY OPNO"
    End If
    
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
         While Not .EOF
            cmbOP.AddItem CStr(!opNo)
            .MoveNext
         Wend
      End With
      If cmbOP.ListCount > 0 Then
         cmbOP.ListIndex = 0
         GetOverheadRate
      End If
   Else
      MsgBox "No open operations for this MO", vbExclamation ', sSysCaption
   End If
   
End Sub

Private Function GetTCServiceOp() As Boolean
    ' get COPOTIMESERVOP flag
    Dim RdoGet As ADODB.Recordset
    Dim bSerFlg As Boolean
   On Error GoTo whoops
    
    sSql = "SELECT COPOTIMESERVOP FROM ComnTable WHERE COREF=1"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_KEYSET)
    If bSqlRows Then
       With RdoGet
          bSerFlg = "" & Trim(!COPOTIMESERVOP)
          GetTCServiceOp = bSerFlg
       End With
       RdoGet.Close
    Else
        ' record not found
        GetTCServiceOp = True
    End If
   Exit Function
   
whoops:
   Exit Function
End Function

Private Function GetOverheadRate()
   Dim rdoOverHead As ADODB.Recordset
   Dim strPart As String
   Dim lngRun As Long
   Dim intOp As Integer
   
   strPart = cmbPrt
   lngRun = cmbRun
   intOp = cmbOP
   
   sSql = "SELECT RnopTable.OPREF, RnopTable.OPRUN, RnopTable.OPNO, RnopTable.OPSHOP, " _
          & " RnopTable.OPCENTER, WcntTable.WCNOHFIXED, WcntTable.WCNOHPCT, WcntTable.WCNSTDRATE " _
          & " FROM RnopTable INNER JOIN WcntTable ON RnopTable.OPSHOP = WcntTable.WCNSHOP AND " _
          & " RnopTable.OPCENTER = WcntTable.WCNREF " _
          & "WHERE (RnopTable.OPREF = '" & Compress(strPart) & "') AND (RnopTable.OPRUN = " & lngRun & ") AND (RnopTable.OPNO = " & intOp & ")"
   
   Debug.Print sSql
   
   gblnSqlRows = clsADOCon.GetDataSet(sSql, rdoOverHead)
   If gblnSqlRows Then
      With rdoOverHead
         lblCurFOH = !WCNOHFIXED
         lblCurPOH = !WCNOHPCT
      End With
   Else
      lblCurFOH = 0
      lblCurPOH = 0
   End If
End Function


'Private Function GetEmployeeRate() As Currency
'
'   Dim empno As Long
'   Dim empRate As Currency
'   'get employee information
'   Dim rdo As ADODB.Recordset
'   sSql = "select PREMPAYRATE, rtrim(PREMACCTS) as PREMACCTS from EmplTable where PREMNUMBER = " & empno
'   If clsADOCon.GetDataSet(sSql, rdo) Then
'      empRate = rdo!PREMPAYRATE * Me.GetTimeCodeMultiplier(TimeCode)
'      empAccount = rdo!PREMACCTS
'   Else
'      empRate = 0
'      empAccount = ""
'   End If
'   Set rdo = Nothing
'End Function
'
