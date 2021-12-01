VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PickMCp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pick Expediting"
   ClientHeight    =   4455
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkRL 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4080
      TabIndex        =   26
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox chkSC 
      Caption         =   "SC"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox chkPP 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3360
      TabIndex        =   24
      Top             =   3120
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox chkPL 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   3120
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Select Run Number"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PickMCp02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   2160
      TabIndex        =   17
      Top             =   2280
      Width           =   2535
      Begin VB.OptionButton optUop 
         Caption         =   "Used On Parts"
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton optMon 
         Caption         =   "MO Numbers"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   3720
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   3975
      Width           =   735
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains MO's PP and PL"
      Top             =   960
      Width           =   3545
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PickMCp02a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "PickMCp02a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4455
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PP"
      Height          =   255
      Index           =   12
      Left            =   3000
      TabIndex        =   29
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "RL"
      Height          =   255
      Index           =   11
      Left            =   3720
      TabIndex        =   28
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SC"
      Height          =   255
      Index           =   10
      Left            =   4560
      TabIndex        =   27
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PL"
      Height          =   255
      Index           =   9
      Left            =   2160
      TabIndex        =   22
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   7
      Left            =   3480
      TabIndex        =   21
      Top             =   1460
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report By:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   2
      Left            =   5880
      TabIndex        =   16
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sched Pick Start Before"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   12
      Top             =   3975
      Width           =   2055
   End
   Begin VB.Label lblSel 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Number(s)"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "PickMCp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Dim sPartNumber As String
Dim sOldPart As String
Dim bGoodRuns As Byte
Dim strRunRef As String
Dim strRunNo As String
Dim cRunqty As Currency
Dim bSqlRows1 As Boolean




Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   If cmbPrt <> "" Then
      bGoodRuns = GetRuns()
   End If
   If Len(Trim(cmbPrt)) = 0 Then
    cmbPrt = "ALL"
    cmbRun = "ALL"
   End If
   
End Sub

Private Function GetRuns() As Byte
   Dim RdoRns As ADODB.Recordset
   If sOldPart <> cmbPrt Then
      cmbRun.Clear
      sOldPart = cmbPrt
   Else
      GetRuns = 1
      Exit Function
   End If
   sOldPart = cmbPrt
   sPartNumber = Compress(cmbPrt)
   On Error GoTo DiaErr1
   AdoQry.Parameters(0).Value = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoQry)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr cmbRun.hWnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
      GetRuns = 1
   Else
      sPartNumber = ""
      GetRuns = 0
   End If
   
   If cmbRun.ListCount > 0 Then cmbRun = cmbRun.List(0)
   Set RdoRns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbRun_LostFocus()
   If Len(Trim(cmbRun)) = 0 Or (cmbPrt = "ALL") Then
    cmbRun = "ALL"
   End If

End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub FillCombo1()
   On Error GoTo DiaErr1
   cmbPrt.Clear
   sSql = "SELECT DISTINCT RUNREF,PARTREF,PARTNUM FROM " _
          & "RunsTable,PartTable WHERE RUNREF=PARTREF " _
          & "AND (RUNSTATUS='PL' OR RUNSTATUS='PP') " _
          & "ORDER BY RUNREF"
   LoadComboBox cmbPrt, 1
   cmbPrt = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo1"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      If optMon.Value = True Then FillCombo1 Else FillCombo2
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO,RUNPLDATE FROM " _
          & "RunsTable WHERE RUNREF = ? " _
          & "AND RUNSTATUS<>'CA'"
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Size = 30
   AdoParameter.Type = adChar
   AdoQry.Parameters.Append AdoParameter
   
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set PickMCp02a = Nothing
   
End Sub

Private Sub PrintReport1()
   Dim sPrtNumber As String
   Dim sRunNo As String
   Dim sBegDate As String
   Dim sEndDate As String
   MouseCursor 13
   
   On Error GoTo DiaErr1
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sqlRunStat As String
   
   sBegDate = "1995,01,01"
   sEndDate = Format(txtBeg, "yyyy,mm,dd")
   
   If cmbPrt <> "ALL" Then
    sPrtNumber = Compress(cmbPrt)
   Else
    sPrtNumber = ""
   End If
   
   If cmbRun <> "ALL" Then
    sRunNo = Compress(cmbRun)
   Else
    sRunNo = ""
   End If
   
   sqlRunStat = ""
   If (chkPP.Value = vbChecked) Then
      sqlRunStat = "{RunsTable.RUNSTATUS} = 'PP'"
   End If
   
   If (chkPL.Value = vbChecked) Then
      If (sqlRunStat <> "") Then
            sqlRunStat = sqlRunStat & " OR "
      End If
      sqlRunStat = sqlRunStat & "{RunsTable.RUNSTATUS} = 'PL'"
   End If
   
   If (chkSC.Value = vbChecked) Then
      If (sqlRunStat <> "") Then
            sqlRunStat = sqlRunStat & " OR "
      End If
      sqlRunStat = sqlRunStat & "{RunsTable.RUNSTATUS} = 'SC'"
   End If
   
   If (chkRL.Value = vbChecked) Then
      If (sqlRunStat <> "") Then
            sqlRunStat = sqlRunStat & " OR "
      End If
      sqlRunStat = sqlRunStat & "{RunsTable.RUNSTATUS} = 'RL'"
   End If
   
   If (sqlRunStat = "") Then
      MouseCursor 0
      MsgBox "Please select at least one Run Status.", vbInformation, Caption
      Exit Sub
   End If
   
   sCustomReport = GetCustomReport("prdma03")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowPartDesc"
   aFormulaName.Add "ShowExtDesc"
    
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & cmbPrt & "...., To " & txtBeg & "...'")
   aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
   aFormulaValue.Add CStr("'" & optDsc.Value & "'")
   aFormulaValue.Add CStr("'" & optExt.Value & "'")
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{RunsTable.RUNREF} LIKE '" & sPrtNumber & "*' AND " _
         & "{RunsTable.RUNPKSTART} in Date(" & sBegDate & ") " _
         & "to Date(" & sEndDate & ") AND (" & sqlRunStat & ")"
         
         '& "({RunsTable.RUNSTATUS} = 'PL' OR {RunsTable.RUNSTATUS} = 'PP')"
   '          & " {MopkTable.PKAQTY} = 0.00 and {MopkTable.PKPQTY}>{PartTable.PAQOH} AND "
   
   If (sRunNo <> "") Then
          sSql = sSql & " AND {RunsTable.RunNo} = " & sRunNo
   End If
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   Dim bByte As Byte
   Dim sOptions As String
   'Save by Menu Option
   If optMon.Value = True Then bByte = 1 Else bByte = 0
   sOptions = Trim(str(bByte)) _
              & Trim(str(optDsc.Value)) _
              & Trim(str(optExt.Value))
   SaveSetting "Esi2000", "EsiProd", "ma03", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   'Get By Menu Option
   sOptions = GetSetting("Esi2000", "EsiProd", "ma03", sOptions)
   If Len(sOptions) > 0 Then
      optMon.Value = Val(Left(sOptions, 1))
      optDsc.Value = Val(Mid(sOptions, 2, 1))
      optExt.Value = Val(Mid(sOptions, 3, 1))
   Else
      optMon.Value = True
      optDsc.Value = vbChecked
      optExt.Value = vbChecked
   End If
   If optMon.Value = False Then optUop.Value = True
   txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub optDis_Click()
   If optMon.Value = True Then
      CreatePickList
      PrintReport1
   Else
      PrintReport2
   End If
   
End Sub

Private Sub CreatePickList()
   
   Dim RdoLst As ADODB.Recordset
   Dim RdoRun As ADODB.Recordset
   Dim RdoMoPk As ADODB.Recordset
   Dim bsqlMoPkRows
   
   Dim sPrtNumber As String
   Dim sRunNo As String
   Dim sBegDate As String
   Dim sEndDate As String
   Dim sBomRev As String
   
   On Error GoTo DiaErr1
   Dim sqlRunStat As String
   Dim sqlRunNo As String
   Dim strPartRef As String
   Dim strPQty As String
   Dim strAQty As String
   Dim cQuantity As Currency
   Dim cConversion As Currency
   Dim cSetup As Currency
   Dim cRunqty As Currency
    
   
   Dim dDate As Date
   dDate = Format(ES_SYSDATE, "mm/dd/yyyy")
   sBegDate = "1/1/1995"
   sEndDate = Format(txtBeg, "mm/dd/yyyy")
   
   If cmbPrt <> "ALL" Then
    sPrtNumber = Compress(cmbPrt)
   Else
    sPrtNumber = ""
   End If
   
   If cmbRun <> "ALL" And cmbRun <> "" Then
    sRunNo = Compress(cmbRun)
    sqlRunNo = " AND RUNNO = '" & sRunNo & "'"
   Else
    sRunNo = ""
    sqlRunNo = ""
   End If
   
   sqlRunStat = ""
   If (chkPP.Value = vbChecked) Then
      sqlRunStat = "RUNSTATUS = 'PP'"
   End If
   
   If (chkPL.Value = vbChecked) Then
      If (sqlRunStat <> "") Then
            sqlRunStat = sqlRunStat & " OR "
      End If
      sqlRunStat = sqlRunStat & "RUNSTATUS = 'PL'"
   End If
   
   If (chkSC.Value = vbChecked) Then
      If (sqlRunStat <> "") Then
            sqlRunStat = sqlRunStat & " OR "
      End If
      sqlRunStat = sqlRunStat & "RUNSTATUS = 'SC'"
   End If
   
   If (chkRL.Value = vbChecked) Then
      If (sqlRunStat <> "") Then
            sqlRunStat = sqlRunStat & " OR "
      End If
      sqlRunStat = sqlRunStat & "RUNSTATUS = 'RL'"
   End If
   
   If (sqlRunStat = "") Then
      MouseCursor 0
      MsgBox "Please select at least one Run Status.", vbInformation, Caption
      Exit Sub
   End If

   sSql = "DELETE FROM EsReportPickExpedite"
   clsADOCon.ExecuteSQL sSql

   sSql = "SELECT DISTINCT RUNREF, RUNNO, RUNQTY FROM RunsTable WHERE RUNREF LIKE '" & sPrtNumber & "%' " _
            & sqlRunNo & " AND RUNPKSTART BETWEEN '" & sBegDate & "' AND " _
          & "'" & sEndDate & "' " _
            & " AND (" & sqlRunStat & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_STATIC)
   
   If bSqlRows Then
      With RdoRun
         Do Until .EOF
            strRunRef = Trim(!RUNREF)
            strRunNo = Trim(!Runno)
            cRunqty = Format(!RUNQTY, ES_QuantityDataFormat)
            
            sSql = "SELECT PKPARTREF, PKMOPART, PKMORUN,PKPQTY, PKAQTY " _
                     & " FROM MopkTable WHERE PKMOPART = '" & strRunRef & "' " _
                        & " AND PKMORUN = '" & strRunNo & "'"
            
            bsqlMoPkRows = clsADOCon.GetDataSet(sSql, RdoMoPk, ES_STATIC)
            If bsqlMoPkRows Then
               With RdoMoPk
                  Do Until .EOF
                     strPartRef = Trim(Compress(!PKPARTREF))
                     strPQty = Trim(!PKPQTY)
                     strAQty = Trim(!PKAQTY)
                     
                     sSql = "INSERT INTO EsReportPickExpedite (tPKMOPART, tPKMORUN, tPKPARTREF, " _
                              & " tPKPQTY, tPKAQTY, tMOPicked ) " _
                              & " VALUES ('" & strRunRef & "','" & strRunNo & "','" _
                               & strPartRef & "','" & strPQty & "','" & strAQty & "','1')"
                     
                     clsADOCon.ExecuteSQL sSql
                     
                     .MoveNext
                  Loop
               End With
               ClearResultSet RdoMoPk
            Else
               'MoPk not found - creat a list for each MO item
               ' Empty for now
               sBomRev = ""
               'determine whether any part list for this part and rev
               sSql = "SELECT BMHREF,BMHREV,BMHOBSOLETE,BMHRELEASED,BMHEFFECTIVE " & vbCrLf _
                      & "FROM BmhdTable" & vbCrLf _
                      & "WHERE BMHREF='" & strRunRef & "' AND BMHREV='" & sBomRev & "' " & vbCrLf _
                      & "AND (BMHOBSOLETE IS NULL OR BMHOBSOLETE >='" & dDate & "') AND BMHRELEASED=1"
               bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
               If bSqlRows Then
                 ClearResultSet RdoLst
               
                  sSql = "SELECT * FROM BmplTable" & vbCrLf _
                     & "WHERE BMASSYPART='" & strRunRef & "'" & vbCrLf _
                     & "AND BMREV='" & sBomRev & "'" & vbCrLf _
                     & "ORDER BY BMSEQUENCE"
                  bSqlRows1 = clsADOCon.GetDataSet(sSql, RdoLst, ES_STATIC)
                  If bSqlRows1 Then
                     With RdoLst
                        Do Until .EOF
                           If Not IsNull(!BMSETUP) Then
                              cSetup = !BMSETUP
                           Else
                              cSetup = 0
                           End If
                           cQuantity = Format((cRunqty * (!BMQTYREQD + !BMADDER) + cSetup), "######0.000")
                           If !BMCONVERSION <> 0 Then
                              cQuantity = cQuantity / !BMCONVERSION
                           End If
            
                           'if phantom item, then explode it
                           If !BMPHANTOM = 1 Then
                              
                              InsertPhantom Trim(!BMPARTREF), Trim(!BMPARTREV), cQuantity
                              
                              sSql = "INSERT INTO EsReportPickExpedite (tPKMOPART, tPKMORUN, tPKPARTREF, " _
                                       & " tPKPQTY, tPKAQTY, tMOPicked ) " _
                                       & " VALUES ('" & strRunRef & "','" & strRunNo & "','" _
                                        & Trim(!BMPARTREF) & "','" & cQuantity & "','0.00','0')"
                           
                              clsADOCon.ExecuteSQL sSql
                           'else add non-phantom item to pick list
                           Else
                              sSql = "INSERT INTO EsReportPickExpedite (tPKMOPART, tPKMORUN, tPKPARTREF, " _
                                       & " tPKPQTY, tPKAQTY, tMOPicked ) " _
                                       & " VALUES ('" & strRunRef & "','" & strRunNo & "','" _
                                        & Trim(!BMPARTREF) & "','" & cQuantity & "','0.00','0')"
                              
                              clsADOCon.ExecuteSQL sSql
                              
                           End If
                           
                           .MoveNext
                        Loop
                        ClearResultSet RdoLst
                     End With
                  End If
               End If
            End If
            .MoveNext
         Loop
      ClearResultSet RdoRun
      End With
   End If
   Set RdoLst = Nothing
   Set RdoRun = Nothing
   Set RdoMoPk = Nothing
   Exit Sub

DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub
Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optMon_Click()
   If optMon Then
      cmbPrt.ToolTipText = "Contains MO's PP and PL"
      lblSel(0) = "MO Number(s)"
      FillCombo1
   End If
   
End Sub

Private Sub optPrn_Click()
   If optMon.Value = True Then
      PrintReport1
   Else
      PrintReport2
   End If
   
End Sub


Private Sub optUop_Click()
   If optUop.Value = True Then
      cmbPrt.ToolTipText = "Contains Required Parts"
      lblSel(0) = "Used On Part(s)"
      FillCombo2
   End If
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   txtBeg = CheckDateEx(txtBeg)
   
End Sub



Private Sub FillCombo2()
   On Error GoTo DiaErr1
   cmbPrt.Clear
   sSql = "SELECT DISTINCT PKPARTREF,PARTREF,PARTNUM FROM " _
          & "MopkTable,PartTable WHERE PKPARTREF=PARTREF " _
          & "AND (PKPQTY>0 AND PKAQTY=0)"
   LoadComboBox cmbPrt, 1
   cmbPrt = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo2"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport2()
   Dim sPrtNumber As String
   Dim sBegDate As String
   Dim sEndDate As String
   Dim sRunNo As String
   MouseCursor 13
   
   On Error GoTo DiaErr1
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sqlRunStat As String
   
   sBegDate = "1995,01,01"
   sEndDate = Format(txtBeg, "yyyy,mm,dd")
   If cmbPrt <> "ALL" Then
    sPrtNumber = Compress(cmbPrt)
   Else
    sPrtNumber = ""
   End If
   
   If cmbRun <> "ALL" Then
    sRunNo = Compress(cmbRun)
   Else
    sRunNo = ""
   End If
   
   sqlRunStat = ""
   If (chkPP.Value = vbChecked) Then
      sqlRunStat = "{RunsTable.RUNSTATUS} = 'PP'"
   End If
   
   If (chkPL.Value = vbChecked) Then
      If (sqlRunStat <> "") Then
            sqlRunStat = sqlRunStat & " OR "
      End If
      sqlRunStat = sqlRunStat & "{RunsTable.RUNSTATUS} = 'PL'"
   End If
   
   If (chkSC.Value = vbChecked) Then
      If (sqlRunStat <> "") Then
            sqlRunStat = sqlRunStat & " OR "
      End If
      sqlRunStat = sqlRunStat & "{RunsTable.RUNSTATUS} = 'SC'"
   End If
   
   If (chkRL.Value = vbChecked) Then
      If (sqlRunStat <> "") Then
            sqlRunStat = sqlRunStat & " OR "
      End If
      sqlRunStat = sqlRunStat & "{RunsTable.RUNSTATUS} = 'RL'"
   End If
   
   If (sqlRunStat = "") Then
      MouseCursor 0
      MsgBox "Please select at least one Run Status.", vbInformation, Caption
      Exit Sub
   End If
   
   sCustomReport = GetCustomReport("prdma03b")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowPartDesc"
   aFormulaName.Add "ShowExtDesc"
    
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & cmbPrt & "...., To " & txtBeg & "...'")
   aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
   aFormulaValue.Add CStr("'" & optDsc.Value & "'")
   aFormulaValue.Add CStr("'" & optExt.Value & "'")
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{MopkTable.PKPARTREF} LIKE '" & sPrtNumber & "*' AND " _
          & "{RunsTable.RUNPKSTART} in Date(" & sBegDate & ") " _
          & "to Date(" & sEndDate & ") AND " _
          & " {MopkTable.PKAQTY} = 0.00 and {MopkTable.PKPQTY}>{PartTable.PAQOH} AND " _
          & "(" & sqlRunStat & ") "
   
    If (sRunNo <> "") Then
           sSql = sSql & " AND {MopkTable.PKMORUN} = " & sRunNo
    End If

   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub




Public Sub InsertPhantom(AssyPart As String, AssyRev As String, AssyQuantity As Currency)
   Dim RdoPhn As ADODB.Recordset
   Dim iList As Integer
   Dim iTotalPhantom As Integer
   Dim cPQuantity As Currency
   Dim cPConversion As Currency
   Dim cPSetup As Currency
   
   sSql = "SELECT * FROM BmplTable" & vbCrLf _
      & "WHERE BMASSYPART='" & AssyPart & "'" & vbCrLf _
      & "AND BMREV='" & AssyRev & "'" & vbCrLf _
      & "ORDER BY BMSEQUENCE"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPhn, ES_STATIC)
   iList = -1
   If bSqlRows Then
      With RdoPhn
         Do Until .EOF
            iList = iList + 1
            If Not IsNull(!BMSETUP) Then
               cPSetup = !BMSETUP
            Else
               cPSetup = 0
            End If
            cPQuantity = Format((AssyQuantity * (!BMQTYREQD + !BMADDER) + cPSetup), "######0.000")
            
            If !BMCONVERSION <> 0 Then
               cPQuantity = cPQuantity / !BMCONVERSION
            End If
            
            If !BMPHANTOM = 1 Then
               InsertPhantom Trim(!BMPARTREF), Trim(!BMPARTREV), cPQuantity
            Else
            
               sSql = "INSERT INTO EsReportPickExpedite (tPKMOPART, tPKMORUN, tPKPARTREF, " _
                        & " tPKPQTY, tPKAQTY, tMOPicked ) " _
                        & " VALUES ('" & strRunRef & "','" & strRunNo & "','" _
                         & Trim(!BMPARTREF) & "','" & cPQuantity & "','" & cPQuantity & "','0')"
               
               clsADOCon.ExecuteSQL sSql
            
            End If
            .MoveNext
         Loop
         ClearResultSet RdoPhn
      End With
   End If
   Set RdoPhn = Nothing
End Sub


