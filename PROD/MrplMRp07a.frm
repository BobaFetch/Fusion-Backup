VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form MrplMRp07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculate ROLT by Part"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRefresh 
      Cancel          =   -1  'True
      Caption         =   "&RefreshExport"
      Height          =   480
      Left            =   7680
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1785
   End
   Begin VB.TextBox txtPrt 
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Tag             =   "3"
      Top             =   240
      Width           =   2895
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2880
      TabIndex        =   31
      Top             =   240
      Width           =   2895
   End
   Begin VB.CheckBox cbAssignedRoutings 
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   2400
      Width           =   615
   End
   Begin VB.CheckBox cbCalcRouting 
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   3000
      Width           =   495
   End
   Begin VB.CheckBox cbPartType 
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   5040
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.CheckBox cbPartType 
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.CheckBox cbPartType 
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.CheckBox cbPartType 
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2520
      TabIndex        =   24
      Top             =   4200
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtAdminDays 
      Height          =   285
      Left            =   2880
      TabIndex        =   11
      Top             =   3480
      Width           =   615
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   8400
      TabIndex        =   20
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   2880
      TabIndex        =   8
      ToolTipText     =   "Select Product Class From List"
      Top             =   1800
      Width           =   975
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   2880
      TabIndex        =   7
      ToolTipText     =   "Select Product Code From List"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Height          =   375
      Left            =   5880
      Picture         =   "MrplMRp07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Find Part"
      Top             =   240
      Visible         =   0   'False
      Width           =   395
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "&Close"
      Height          =   360
      Left            =   8400
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.TextBox cmbPrt1 
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      ToolTipText     =   "Requires A Valid Part Number"
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblRecords 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2520
      TabIndex        =   30
      Top             =   4680
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Use Only Assigned Routings"
      Height          =   375
      Index           =   11
      Left            =   240
      TabIndex        =   29
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "(Based on Recommended Run Qty)"
      Height          =   255
      Index           =   10
      Left            =   3480
      TabIndex        =   28
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Queue/Setup and Run Times"
      Height          =   375
      Index           =   9
      Left            =   240
      TabIndex        =   27
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Calculate from Routing"
      Height          =   375
      Index           =   8
      Left            =   240
      TabIndex        =   26
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Part Type(s)"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   25
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Admin Days to Add to Calculation"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   23
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "(Blank for ALL)"
      Height          =   255
      Index           =   6
      Left            =   4440
      TabIndex        =   19
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "(Blank for ALL)"
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   18
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Product Classes"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Product Codes"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "(Blank for ALL)"
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   15
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2880
      TabIndex        =   14
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "MrplMRp07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'03/24/2011 New

Option Explicit

Dim bOnLoad As Byte
Dim bAtLeastOneDefaultRouting As Byte
Dim cAvgWorkWeekHrs As Currency

Dim iWorkDays As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd



Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If cmbCde = "" Then cmbCde = "ALL"
End Sub



Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 6)
   If cmbCls = "" Then cmbCls = "ALL"
End Sub

Private Sub cmbPrt_Click()
    If Len(cmbPrt) = 0 Or cmbPrt = "ALL" Then
        lblDsc = "** ALL PARTS **"
        cmbPrt = "ALL"
    Else
        GetPart
    End If
End Sub

Private Sub cmbPrt_Change()
    If Len(cmbPrt) = 0 Or cmbPrt = "ALL" Then
        lblDsc = "** ALL PARTS **"
        cmbPrt = "ALL"
    Else
        GetPart
    End If
End Sub

Private Sub cmbPrt_LostFocus()
    If Len(cmbPrt) = 0 Or cmbPrt = "ALL" Then
        lblDsc = "** ALL PARTS **"
        cmbPrt = "ALL"
    Else
        GetPart
    End If
End Sub


Private Sub cmdCan_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    ViewParts.lblControl = "CMBPRT"
    ViewParts.txtPrt = cmbPrt
    ViewParts.Show
End Sub

Private Sub cmdRefresh_Click()
   On Error GoTo DiaErr1
   RemoveReportData
   If GenerateReportData Then
      MsgBox "Generated ROLT report to view in Excel.", vbInformation
   Else
      MsgBox "Couln't generated ROLT data.", vbInformation
   End If
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "RefreshData"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      cmbCde.AddItem "ALL"
      FillProductCodes
      If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
      cmbCls.AddItem "ALL"
      FillProductClasses
      If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
      
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillPartCombo cmbPrt
      
      cmbPrt = "ALL"
   End If
   bOnLoad = 0
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   SetupReportTables
   cAvgWorkWeekHrs = GetAvgWorkWeekHrs
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
    Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set MrplMRp07a = Nothing
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim i As Integer
   
   On Error Resume Next
       
   sOptions = Left(Trim(cmbPrt) & Space(30), 30)
   sOptions = sOptions & Left(Trim(cmbCde) & Space(6), 6)
   sOptions = sOptions & Left(Trim(cmbCls) & Space(4), 4)
   sOptions = sOptions & Left(Trim(txtAdminDays) & Space(6), 6)
   For i = 1 To 4
    sOptions = sOptions & cbPartType(i).Value
   Next i
   sOptions = sOptions & cbAssignedRoutings.Value
   sOptions = sOptions & cbCalcRouting.Value
   
   SaveSetting "Esi2000", "EsiProd", "mrp07a", Trim(sOptions)
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim i As Integer

   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "mrp07a", Trim(sOptions))
   If Len(sOptions) > 0 Then cmbPrt = Trim(Left(sOptions, 30)) Else cmbPrt = "ALL"
   If Len(sOptions) > 30 Then cmbCde = Trim(Mid(sOptions, 31, 6)) Else cmbCde = "ALL"
   If Len(sOptions) > 36 Then cmbCls = Trim(Mid(sOptions, 37, 4)) Else cmbCls = "ALL"
   If Len(sOptions) > 40 Then txtAdminDays = Trim(Mid(sOptions, 41, 6)) Else txtAdminDays = ""
   For i = 1 To 4
    If Len(sOptions) > 46 Then
        cbPartType(i).Value = Mid(sOptions, 46 + i, 1)
    Else
        cbPartType(i).Value = 1
    End If
   Next i
   If Len(sOptions) > 50 Then cbAssignedRoutings = Mid(sOptions, 51, 1) Else cbAssignedRoutings = 1
   If Len(sOptions) > 51 Then cbCalcRouting = Mid(sOptions, 52, 1) Else cbCalcRouting = 1

End Sub



Private Sub GetPart()
   Dim RdoPrt As ADODB.Recordset
   sSql = "Qry_GetPartNumberBasics '" & Compress(cmbPrt) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(.Fields(1))
         If Len(cmbPrt) > 0 Then
            lblDsc = "" & Trim(.Fields(2))
         Else
            lblDsc = "*** Part Number Wasn't Found ***"
         End If
         ClearResultSet RdoPrt
      End With
   Else
      lblDsc = "*** Part Number Wasn't Found ***"
   End If
   Set RdoPrt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub SetupReportTables()
    On Error Resume Next
    
    If TableExists("EsReportROLT") Then
        If Not ColumnExists("EsReportROLT", "OriginalPartROLT") Then
            sSql = "DROP Table EsReportROLT"
            clsADOCon.ExecuteSQL sSql
        End If
    End If
    
    If TableExists("EsReportROLTDetail") Then
        If Not ColumnExists("EsReportROLTDetail", "RoltPartQueueHrs") Then
            sSql = "DROP Table EsReportROLTDetail"
            clsADOCon.ExecuteSQL sSql
        End If
    End If
    
    
    'If the table doesn't exist, create it
    If Not TableExists("EsReportROLT") Then
        sSql = "CREATE TABLE EsReportROLT (ROLTUser varchar(4) NULL, PartNumber char(30) NOT NULL, PartType tinyint NULL, PartCode char(6) NULL, PartClass char(4) NULL, " & _
               "RecRunQty decimal(12, 4) NULL, OriginalPartROLT decimal(12,4) NULL, AdminDays decimal(12, 4) NULL, LongLeadTime decimal(12, 4) NULL, PreMfgLeadTime decimal(12, 4) NULL, " & _
               "MfgFlowTime decimal(12, 4) NULL, ROLTCalc decimal(12, 4) NULL, ROLTCurr decimal(12, 4) NULL, ROLTDefRouting smallint not null default 0)  ON [PRIMARY]"
        clsADOCon.ExecuteSQL sSql
        
        sSql = "CREATE NONCLUSTERED INDEX [USERPART] ON [dbo].[EsReportROLT] ([ROLTUser] ASC,[PartNumber] ASC) WITH (SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF) ON [PRIMARY]"
        clsADOCon.ExecuteSQL sSql
    End If
    
    If Not TableExists("EsReportROLTDetail") Then
        sSql = "CREATE TABLE EsReportROLTDetail (ROLTUser varchar(4) NULL, ROLTLevel smallint NULL, ROLTAssembly char(30) NULL, " & _
            "ROLTPartRef char(30) NULL, ROLTRevision char(4) NULL, ROLTSequence smallint NOT NULL DEFAULT 0, ROLTSortKey varchar(256) NULL, ROLTMakeBuy char(1) NULL, " & _
            "ROLTPartTime decimal(12, 4) NULL, ROLTRoutingMoveHrs decimal(12,4), ROLTRoutingQueueHrs decimal(12,4), ROLTRoutingSetupHrs decimal(12,4), ROLTRoutingRunHrs decimal(12,4), " & _
            "ROLTDefRouting smallint not null DEFAULT 0) ON [PRIMARY]"
        clsADOCon.ExecuteSQL sSql
    End If
    
 
    
End Sub

Private Sub RemoveReportData()
    'Remove all data from table
    sSql = "DELETE FROM EsReportROLT WHERE ROLTUser = '" & sInitials & "'"
    clsADOCon.ExecuteSQL sSql
    
    
    'Remove all data from table
    sSql = "DELETE FROM EsReportROLTDetail WHERE ROLTUser = '" & sInitials & "'"
    clsADOCon.ExecuteSQL sSql

End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
    PrintReport
End Sub



Private Sub PrintReport()
   If cbPartType(1).Value = 0 And cbPartType(2).Value = 0 And cbPartType(3) = 0 And cbPartType(4) = 0 Then
      MsgBox "You Must Select at Least One Part Type", vbOKOnly
      Exit Sub
   End If
   
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   RemoveReportData
   If Not GenerateReportData Then
        MouseCursor 0
        Exit Sub
   End If
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sPart As String
   Dim sYesNo As String
   
   
   If Len(Compress(cmbPrt)) = 0 Or cmbPrt = "ALL" Then sPart = "" Else sPart = Compress(cmbPrt)
   
  
   sCustomReport = GetCustomReport("prdmr07")
    
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath

   cCRViewer.SetReportTitle = "prdmr07"
   cCRViewer.ShowGroupTree False
    
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes1"
   aFormulaName.Add "Includes2"
   aFormulaName.Add "Includes3"
   aFormulaName.Add "Includes4"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr("Include Part Number: " & sPart) & "'")
   aFormulaValue.Add CStr("'" & CStr("Product Code: " & cmbCde & "  Product Class: " & cmbCls & "  Admin Days: " & txtAdminDays.Text) & "'")
   If cbAssignedRoutings.Value = 1 Then sYesNo = "Yes" Else sYesNo = "No"
   aFormulaValue.Add CStr("'" & CStr("Use Only Assigned Routings? " & sYesNo) & "'")
   If cbCalcRouting.Value = 1 Then sYesNo = "Yes" Else sYesNo = "No"
   aFormulaValue.Add CStr("'" & CStr("Calculate from Queue/Setup/Run Times? " & sYesNo) & "'")
      
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   'cCRViewer.SetReportSelectionFormula sSql
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
    ' print the copies
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




Private Function GenerateReportData() As Boolean
    Dim sPart As String
    Dim RdoPart As ADODB.Recordset
    Dim i As Integer
    Dim sPartTypeIN As String
    Dim cRecRunQty As Currency
    Dim iAdminDays As Integer
    
    
    Dim cLongLeadItem As Currency
    Dim cPreMfgLeadTime As Currency
    Dim cMaxFlowTime As Currency
    Dim cRoltCalculated As Currency
    Dim cLastROLT As Currency
    Dim cOriginalPartROLT As Currency

    Dim sCurrPartNum As String
    
    Dim cQueueHrs, cMoveHrs, cUnitHrs, cSetupHrs As Currency
    
    sPartTypeIN = ""
    For i = 1 To 4
        If cbPartType(i).Value = vbChecked Then
            If Len(sPartTypeIN) = 0 Then sPartTypeIN = LTrim(str(i)) Else sPartTypeIN = sPartTypeIN & "," & LTrim(str(i))
        End If
    Next i
    

    If Len(Compress(cmbPrt)) = 0 Or Compress(cmbPrt) = "ALL" Then sPart = "" Else sPart = Compress(cmbPrt)
    
    sSql = "SELECT PARTREF, PARTNUM, PABOMREV, PADESC, PALEVEL, PAMAKEBUY, PACLASS, PAPRODCODE, " & _
            "PAFLOWTIME, PALEADTIME, PARRQ, PALASTROLT, ISNULL(RTQUEUEHRS, 0) RTQUEUEHRS, " & _
               "ISNULL(RTMOVEHRS, 0) RTMOVEHRS, ISNULL(RTSETUPHRS, 0) RTSETUPHRS, " & _
               "ISNULL(RTUNITHRS , 0) RTUNITHRS FROM PartTable " & _
       "LEFT OUTER JOIN RTHDTABLE ON RTREF=PAROUTING " & _
        " WHERE PARTREF LIKE '" & sPart & "'" '%

Debug.Print sSql

'    If sPart <> "" Then
'      sSql = sSql & " WHERE PARTREF  = '" & sPart & "'"
'    Else
'      sSql = sSql & ""
'    End If
    
    If Compress(cmbCls) <> "ALL" Then sSql = sSql & " AND PACLASS = '" & Compress(cmbCls) & "' "
    If Compress(cmbCde) <> "ALL" Then sSql = sSql & " AND PAPRODCODE = '" & Compress(cmbCde) & "' "
    sSql = sSql & " AND PALEVEL IN (" & sPartTypeIN & ") "
    
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoPart, ES_STATIC)
    GenerateReportData = bSqlRows
    If bSqlRows Then
    
    ProgressBar1.min = 0
    ProgressBar1.Max = RdoPart.RecordCount
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    'lblRecords.Visible = True
    
    While Not RdoPart.EOF
        With RdoPart
            'lblRecords.Caption = "Record " & ProgressBar1.Value & " of " & ProgressBar1.Max
            'DoEvents
            cLongLeadItem = 0
            cPreMfgLeadTime = 0
            cMaxFlowTime = 0
            cRoltCalculated = 0
            bAtLeastOneDefaultRouting = 0
            ' MM
            ' MM cAvgWorkWeekHrs = 8
            cOriginalPartROLT = ((!RTQUEUEHRS + !RTMOVEHRS) / cAvgWorkWeekHrs) + ((!RTUNITHRS + !RTSETUPHRS) / cAvgWorkWeekHrs)
            iAdminDays = Val(txtAdminDays.Text)
            
            Debug.Print "*** Original Part : " & "" & RdoPart!PartRef & "  Original Part Rolt: " & cOriginalPartROLT
            
            'Remove all data from table
            sSql = "DELETE FROM EsReportBomROLTTable WHERE BOMUser = '" & sInitials & "'"
            clsADOCon.ExecuteSQL sSql
    
            If GetROLTTimeForPart("" & !PartRef, "" & !PABOMREV, sInitials, cLongLeadItem, _
                           cPreMfgLeadTime, cMaxFlowTime, cRoltCalculated) Then
            
                cRecRunQty = Val("" & !PARRQ)
                cLastROLT = Val("" & !PALASTROLT)
                
                sSql = "INSERT INTO EsReportROLT (ROLTUser, PartNumber,PartType,PartCode,PartClass,RecRunQty, OriginalPartROLT, AdminDays,LongLeadTime,PreMfgLeadTime,MfgFlowTime,RoltCalc,ROLTDefRouting) " & _
                   " Values ('" & sInitials & "','" & "" & !PartNum & "'," & "" & !PALEVEL & ",'" & "" & !PAPRODCODE & "','" & "" & !PACLASS & "'," & cRecRunQty & "," & cOriginalPartROLT & "," & iAdminDays & "," & cLongLeadItem & "," & cPreMfgLeadTime & "," & cMaxFlowTime + cOriginalPartROLT & "," & cRoltCalculated & "," & bAtLeastOneDefaultRouting & ")"
'Debug.Print "FlowTime: " & !PAFLOWTIME & "   LeadTime: " & !PALEADTIME
                clsADOCon.ExecuteSQL sSql
                           
                sSql = "UPDATE PartTable SET PALASTROLT = " & cRoltCalculated & " WHERE PARTREF = '" & "" & !PartRef & "' "
                clsADOCon.ExecuteSQL sSql
                
            End If
            
            ProgressBar1.Value = ProgressBar1.Value + 1
            .MoveNext
        End With
    Wend
    Else
        MsgBox "No Parts Exist with this criteria. Check the Part Type, Class, etc.", vbInformation
    End If
    Set RdoPart = Nothing
    ProgressBar1.Visible = False
    'lblRecords.Visible = False


End Function

Private Function GetROLTTimeForPart(ByVal sPartNum As String, ByVal sPartRev As String, _
                  ByVal sInitials As String, ByRef cLongestLeadItem As Currency, _
                  ByRef cPreManufacturedLeadTime As Currency, _
                  ByRef cMfgFlowTime As Currency, ByRef cRoltCalcWeeks) As Boolean
   'return = True if successful
    
   Dim RdoROLT As ADODB.Recordset
   Dim RdoROLT1 As ADODB.Recordset
   
   Dim cSumAssDays As Currency
   Dim cTotalPurchaseTime As Currency
   Dim cRoltDays As Currency
   
   sSql = "RptROLT '" & sPartNum & "','" & sInitials & "'"
   
   clsADOCon.ExecuteSQL sSql ' rdExecDirect

   ' Just get the LeadTime only
   sSql = "SELECT BomUser,BomAssembly,iPartDays, iPartMakeDays," & _
            " ISNULL(SumAssmDays, 0) SumAssmDays, " & _
            "ISNULL(SumPartMakeDays, 0) SumPartMakeDays ,ISNULL(RoltDays, 0) RoltDays, ExplodedQty " & _
         " FROM EsReportBomROLTTable" & _
            " WHERE BomUser = '" & sInitials & "' AND " & _
            " BomParentPartRef = '" & sPartNum & "'" & _
            " AND BomLevel = 0 " & _
         "order by BomSortKey"
         
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoROLT, ES_FORWARD)
   If bSqlRows Then
       With RdoROLT
          'cLongestLeadItem = !SumAssmDays
          'cPreManufacturedLeadTime = !iPartDays
          'cMfgFlowTime = !SumPartMakeDays
          'ClearResultSet RdoROLT
          
          cSumAssDays = !SumAssmDays
          cPreManufacturedLeadTime = !iPartDays
          cTotalPurchaseTime = cSumAssDays + cPreManufacturedLeadTime
          cMfgFlowTime = (!SumPartMakeDays + !iPartMakeDays)
          cRoltDays = !RoltDays
          If (cSumAssDays <> 0) Then
          ' For Round up = 0.5
            'cRoltCalcWeeks = Round(0.5 + (cPreManufacturedLeadTime + cSumAssDays) / IIf(iWorkDays = 0, 1, iWorkDays), 0)
            cRoltCalcWeeks = Round(((0.5 + cRoltDays) / IIf(iWorkDays = 0, 5, iWorkDays)), 0)
          Else
            cRoltCalcWeeks = 0
          End If
          ClearResultSet RdoROLT
          
       End With
   End If
   
   Set RdoROLT = Nothing
   
   ' Just get the LeadTime only
   sSql = "SELECT MAX(paleadtime) LongestLeadTime " & _
         " FROM EsReportBomROLTTable" & _
            " WHERE BomUser = '" & sInitials & "' AND " & _
            " BomParentPartRef = '" & sPartNum & "'"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoROLT1, ES_FORWARD)
   If bSqlRows Then
       With RdoROLT1
       
          cLongestLeadItem = !LongestLeadTime
          'cRoltCalcWeeks = cLongestLeadItem / IIf(iWorkDays = 0, 1, iWorkDays)
          ClearResultSet RdoROLT1
          
          'cLongestLeadItem = !LongestLeadTime
          'cRoltCalcWeeks = cLongestLeadItem / IIf(iWorkDays = 0, 1, iWorkDays)
          'ClearResultSet RdoROLT1
       End With
   End If
   
   Set RdoROLT1 = Nothing
   
   
   GetROLTTimeForPart = True
   MouseCursor 0
   Exit Function
   

DiaErr1:

   sProcName = "GetROLTTimeForPart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function


'Private Function GetROLTTimeForPart(ByVal sPartNum As String, ByVal sPartRev As String, ByRef cLongestLeadItem As Currency, ByRef cPreManufacturedLeadTime As Currency, ByRef cMfgFlowTime As Currency, ByRef cRoltCalcWeeks) As Boolean
'   'return = True if successful
'   Dim RdoRolt As ADODB.Recordset
'   Dim aRoltPart(0 To 11) As Currency
'   Dim i As Integer
'   Dim iLstLvl As Integer
'   Dim iCurrLvl As Integer
'   Dim cCurrentPartROLT As Currency
'   Dim cStartingROLT As Currency
'   Dim bProcessed As Byte
'
'   Dim cRoutingTime, cPartTime As Currency
'   Dim sMakeBuy As String
'
'   On Error GoTo DiaErr1
'   'DeleteDataForThisUser
'
'   'insert part to be exploded
'   'Dim assy As String
'   'assy = Compress(cmbPls)
'   sSql = "INSERT INTO EsReportROLTDetail" & vbCrLf _
'      & "(ROLTUser,ROLTLevel,ROLTAssembly,ROLTPartRef,ROLTRevision," & vbCrLf _
'      & "ROLTSequence, ROLTSortKey, ROLTMakeBuy, ROLTPartTime, ROLTRoutingMoveHrs, ROLTRoutingQueueHrs, ROLTRoutingSetupHrs, ROLTRoutingRunHrs, ROLTDefRouting)" & vbCrLf _
'      & "SELECT TOP 1 '" & sInitials & "',0,PARTREF,PARTREF,PABOMREV," & vbCrLf _
'      & "0,'',PAMAKEBUY, " & vbCrLf _
'      & "CASE WHEN PAMAKEBUY='M' THEN PAFLOWTIME WHEN PAMAKEBUY='B' THEN PALEADTIME WHEN PAMAKEBUY='E' THEN PAFLOWTIME+PALEADTIME ELSE 0 END, " & vbCrLf _
'      & "CASE WHEN PAROUTING='' THEN 0 ELSE RTMOVEHRS END AS TOTROUTINGMOVEHRS, " & vbCrLf _
'      & "CASE WHEN PAROUTING='' THEN 0 ELSE RTQUEUEHRS END AS TOTROUTINGQUEUEHRS, " & vbCrLf _
'      & "CASE WHEN PAROUTING='' THEN 0 ELSE RTSETUPHRS END AS TOTROUTINGSETUPHRS, " & vbCrLf _
'      & "CASE WHEN PAROUTING='' THEN 0 ELSE (PARRQ * RTUNITHRS) END AS TOTROUTINGRUNHRS, " & vbCrLf _
'      & "CASE WHEN RTREF IS NULL THEN 1 WHEN RTRIM(PAROUTING)='DEFAULT' THEN 1 ELSE 0 END AS USEDEFROUTING " & vbCrLf _
'      & "FROM PartTable" & vbCrLf _
'      & "LEFT JOIN LohdTable on LOTPARTREF = PARTREF" & vbCrLf _
'      & "LEFT OUTER JOIN RTHDTable ON RTREF=PAROUTING " & vbCrLf _
'      & "WHERE PARTREF = '" & Compress(sPartNum) & "' ORDER BY LOTADATE DESC"
'   clsADOCon.ExecuteSQL sSql
'
'   'keep inserting next level parts until there are no more
'   Dim level As Integer
'   For level = 0 To 10
'      sSql = "INSERT INTO EsReportROLTDetail" & vbCrLf _
'         & "(ROLTUser, ROLTLevel, ROLTAssembly,ROLTPartRef,ROLTRevision," & vbCrLf _
'         & "ROLTSequence, ROLTSortKey, ROLTMakeBuy, ROLTPartTime, ROLTRoutingMoveHrs, ROLTRoutingQueueHrs, ROLTRoutingSetupHrs, ROLTRoutingRunHrs, ROLTDefRouting) " & vbCrLf _
'         & "SELECT '" & sInitials & "'," & level + 1 & ",BMASSYPART,BMPARTREF,BMPARTREV," & vbCrLf _
'         & "BMSEQUENCE, ROLTSortKey " & vbCrLf _
'         & "+ left(PT1.partref,15) " & vbCrLf _
'         & " + cast(BMSEQUENCE as varchar(4)) " & vbCrLf _
'         & " + cast( " & vbCrLf _
'         & " (select count(*) FROM BmplTable b, EsReportROLTDetail, PartTable pt2" & vbCrLf _
'         & "    Where b.BMASSYPART = ROLTPartRef" & vbCrLf _
'         & "       AND b.BMREV = ROLTRevision" & vbCrLf _
'         & "       AND ROLTLevel = " & level & vbCrLf _
'         & "       AND pt2.PARTREF = b.BMPARTREF" & vbCrLf _
'         & "       AND b.BMPARTREF <= a.BMPARTREF" & vbCrLf _
'         & "       AND ROLTUser = '" & sInitials & "') as varchar(4)), PT1.PAMAKEBUY, "
'     sSql = sSql _
'         & " CASE WHEN PT1.PAMAKEBUY='M' THEN PT1.PAFLOWTIME WHEN PT1.PAMAKEBUY='B' THEN PT1.PALEADTIME WHEN PT1.PAMAKEBUY='E' THEN PT1.PAFLOWTIME+PT1.PALEADTIME ELSE 0 END, " & vbCrLf _
'         & " CASE WHEN PT1.PAMAKEBUY='B' THEN 0 WHEN PT1.PAROUTING='' THEN 0 ELSE RTMOVEHRS END AS TOTROUTINGMOVEHRS, " & vbCrLf _
'         & " CASE WHEN PT1.PAMAKEBUY='B' THEN 0 WHEN PT1.PAROUTING='' THEN 0 ELSE RTQUEUEHRS END AS TOTROUTINGQUEUEHRS, " & vbCrLf _
'         & " CASE WHEN PT1.PAMAKEBUY='B' THEN 0 WHEN PT1.PAROUTING='' THEN 0 ELSE RTSETUPHRS END AS TOTROUTINGSETUPHRS, " & vbCrLf _
'         & " CASE WHEN PT1.PAMAKEBUY='B' THEN 0 WHEN PT1.PAROUTING='' THEN 0 ELSE (PT1.PARRQ * RTUNITHRS) END AS TOTROUTINGRUNHRS, " & vbCrLf _
'         & " CASE WHEN RTREF IS NULL THEN 1 WHEN RTRIM(PAROUTING)='DEFAULT' THEN 1 ELSE 0 END AS USEDEFROUTING " & vbCrLf _
'         & "FROM BmplTable a" & vbCrLf _
'         & "JOIN EsReportRoltDetail D2 on a.BMASSYPART = D2.ROLTPartRef" & vbCrLf _
'         & "AND a.BMREV = D2.ROLTRevision AND D2.ROLTLevel = " & level & vbCrLf _
'         & "JOIN PartTable PT1 on PT1.PARTREF = BMPARTREF" & vbCrLf _
'         & "LEFT OUTER JOIN RTHDTable ON RTREF=PT1.PAROUTING " & vbCrLf _
'         & "WHERE D2.ROLTUser = '" & sInitials & "'" & vbCrLf _
'         & "ORDER BY D2.ROLTPARTREF"
'
'  'Debug.Print sSql
'
'
'      clsADOCon.ExecuteSQL sSql
'      If clsADOCon.RowsAffected = 0 Then
'         Exit For
'      End If
'   Next
'
'
'
''Now it's time to loop through the records and get the max times
'   For i = LBound(aRoltPart) To UBound(aRoltPart)
'    aRoltPart(i) = 0
'   Next i
'   cLongestLeadItem = 0
'   cPreManufacturedLeadTime = 0
'   cMfgFlowTime = 0
'   cCurrentPartROLT = 0
'   iLstLvl = 0
'   On Error Resume Next
'
'    sSql = "SELECT * FROM EsReportROLTDetail ORDER BY ROLTSortKey"
'    bSqlRows = clsADOCon.GetDataSet(sSql, RdoRolt, ES_FORWARD)
'    If bSqlRows Then
'      With RdoRolt
'        Do While Not .EOF
'            bProcessed = 0
'            sMakeBuy = "" & !ROLTMakeBuy
'            If "" & !ROLTDefRouting = 1 Then bAtLeastOneDefaultRouting = 1
'            If (!ROLTDefRouting = 0 And sMakeBuy = "M") Or ("" & !ROLTDefRouting = 1 And cbAssignedRoutings.Value = 0 And sMakeBuy = "M") Or sMakeBuy = "B" Then
'Debug.Print "Part " & !ROLTPARTREF & "   MoveHRs: " & !ROLTRoutingMoveHrs & "  Queue Hrs: " & !ROLTRoutingQueueHrs & " SetupHrs: " & !ROLTRoutingSetupHrs & "  Run Hrs: " & !ROLTRoutingRunHrs
'
'                cRoutingTime = "" & ((!ROLTRoutingMoveHrs + !ROLTRoutingQueueHrs) / 24) + ((!ROLTRoutingSetupHrs + !ROLTRoutingRunHrs) / cAvgWorkWeekHrs)
'                cPartTime = "" & !ROLTPARTTime
'                iCurrLvl = Val("" & !ROLTLevel)
'                If iCurrLvl = 0 Then
'                    If cbCalcRouting.Value = 1 Then cStartingROLT = cRoutingTime Else cStartingROLT = cPartTime
'                    If sMakeBuy = "M" And cStartingROLT > cLongestLeadItem Then cLongestLeadItem = cStartingROLT
'                    If sMakeBuy = "B" And cStartingROLT > cPreManufacturedLeadTime Then cPreManufacturedLeadTime = cStartingROLT
'                    bProcessed = 1
'                Else
'                    If iCurrLvl <= iLstLvl Then
'                        cCurrentPartROLT = AddArrayElements(aRoltPart)
'                        For i = iCurrLvl + 1 To UBound(aRoltPart)
'                            aRoltPart(i) = 0
'                        Next i
'                        If cbCalcRouting.Value = 1 Then aRoltPart(iCurrLvl) = cRoutingTime Else aRoltPart(iCurrLvl) = cPartTime
'                        If cCurrentPartROLT > cMfgFlowTime Then cMfgFlowTime = cCurrentPartROLT
'                    ElseIf iCurrLvl > iLstLvl Then
'                        If cbCalcRouting.Value = 1 Then aRoltPart(iCurrLvl) = cRoutingTime Else aRoltPart(iCurrLvl) = cPartTime
'                    End If
'                    bProcessed = 1
'                    iLstLvl = iCurrLvl
'                End If
'                If sMakeBuy = "M" And aRoltPart(iCurrLvl) > cLongestLeadItem Then cLongestLeadItem = aRoltPart(iCurrLvl)
'                If sMakeBuy = "B" And cPartTime > cPreManufacturedLeadTime Then cPreManufacturedLeadTime = cPartTime
'
'            End If  'If not default routing
'
'
'            .MoveNext
'        Loop
'        If bProcessed = 0 Then
'            If sMakeBuy = "M" And aRoltPart(iCurrLvl) > cLongestLeadItem Then cLongestLeadItem = aRoltPart(iCurrLvl)
'            If sMakeBuy = "B" And cPartTime > cPreManufacturedLeadTime Then cPreManufacturedLeadTime = cPartTime
'        End If
'
'
'
'      End With
'      ClearResultSet RdoRolt
'    End If
'
'
'    cRoltCalcWeeks = (cMfgFlowTime + cPreManufacturedLeadTime) / 5
'    cLongestLeadItem = cPreManufacturedLeadTime     'This is a cheat for now
'
'
'   Set RdoRolt = Nothing
'   GetROLTTimeForPart = True
'   MouseCursor 0
'   Exit Function
'
'DiaErr1:
'
'   sProcName = "GetROLTTimeForPart"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'End Function
'

Private Function AddArrayElements(cMyArray() As Currency) As Currency
    Dim i As Integer
    AddArrayElements = 0
    On Error Resume Next
    
    For i = LBound(cMyArray) To UBound(cMyArray)
        AddArrayElements = AddArrayElements + cMyArray(i)
    Next i
    
End Function


Function GetAvgWorkWeekHrs() As Currency
   Dim RdoCal As ADODB.Recordset
   Dim AvgHrs(1 To 7) As Currency
   Dim i As Integer
   Dim cTemp As Currency
      
   GetAvgWorkWeekHrs = 0
   For i = 1 To 7
     AvgHrs(i) = 0
   Next i
   iWorkDays = 0
   
   sSql = "SELECT * FROM CctmTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCal)
   If bSqlRows Then
      With RdoCal
        AvgHrs(1) = !CALSUNHR1 + !CALSUNHR2 + !CALSUNHR3 + !CALSUNHR4
        If AvgHrs(1) > 0 Then iWorkDays = iWorkDays + 1
        
        AvgHrs(2) = !CALMONHR1 + !CALMONHR2 + !CALMONHR3 + !CALMONHR4
        If AvgHrs(2) > 0 Then iWorkDays = iWorkDays + 1
        
        AvgHrs(3) = !CALTUEHR1 + !CALTUEHR2 + !CALTUEHR3 + !CALTUEHR4
        If AvgHrs(3) > 0 Then iWorkDays = iWorkDays + 1
        
        AvgHrs(4) = !CALWEDHR1 + !CALWEDHR2 + !CALWEDHR3 + !CALWEDHR4
        If AvgHrs(4) > 0 Then iWorkDays = iWorkDays + 1
        
        AvgHrs(5) = !CALTHUHR1 + !CALTHUHR2 + !CALTHUHR3 + !CALTHUHR4
        If AvgHrs(5) > 0 Then iWorkDays = iWorkDays + 1
      
        AvgHrs(6) = !CALFRIHR1 + !CALFRIHR2 + !CALFRIHR3 + !CALFRIHR4
        If AvgHrs(6) > 0 Then iWorkDays = iWorkDays + 1
      
        AvgHrs(7) = !CALSATHR1 + !CALSATHR2 + !CALSATHR3 + !CALSATHR4
        If AvgHrs(7) > 0 Then iWorkDays = iWorkDays + 1
      
      End With
   End If
   Set RdoCal = Nothing
   GetAvgWorkWeekHrs = (AvgHrs(1) + AvgHrs(2) + AvgHrs(3) + AvgHrs(4) + AvgHrs(5) + AvgHrs(6) + AvgHrs(7)) / iWorkDays

End Function

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPrt.Visible = False
      txtPrt.Visible = True
      cmdFind.Visible = True
   Else
      cmbPrt.Visible = True
      txtPrt.Visible = False
      cmdFind.Visible = False
   End If
End Function

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
   
End Sub


Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"
   cmbPrt = txtPrt
End Sub


Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPrt = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPrt
   ViewParts.Show
End Sub

