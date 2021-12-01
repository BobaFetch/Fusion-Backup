VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form diaWipOH 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work In Process Report - Calculated OH"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   360
      Width           =   1335
      Begin VB.CommandButton optPrn 
         Height          =   390
         Left            =   675
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   390
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.ComboBox cboAsOf 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Tag             =   "4"
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox cboClass 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox cboCode 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaWipOH.frx":0000
      PictureDn       =   "diaWipOH.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4920
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2925
      FormDesignWidth =   6795
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   8
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaWipOH.frx":028C
      PictureDn       =   "diaWipOH.frx":03D2
   End
   Begin ComctlLib.ProgressBar Prg1 
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Class"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "As Of "
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Of"
      Height          =   285
      Index           =   7
      Left            =   1920
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Record"
      Height          =   285
      Index           =   8
      Left            =   0
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblRuns 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "diaWipOH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RdoWip As ADODB.Recordset
Dim bOnLoad As Byte
Dim lTotalRuns As Long

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   cboAsOf = Format(ES_SYSDATE, "mm/dd/yy")
   Label2 = ""
End Sub

Private Sub cboAsOf_DropDown()
   ShowCalendar Me
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Function DataReady() As Boolean
   DataReady = False
   lTotalRuns = 0
   lblRuns = "0"
   sSql = "TRUNCATE TABLE EsReportWIPOH"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   If GetWipRuns Then
      If BuildReport Then
         DataReady = True
      End If
   Else
      MsgBox "No WIP Found With The Selected Parameters.", vbInformation, Caption
   End If
  
End Function

Private Sub Form_Activate()
   If bOnLoad Then
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   GetOptions
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   bOnLoad = 1
   
   PopulateCombo cboClass, "PACLASS", "PartTable"
   PopulateCombo cboCode, "PAPRODCODE", "PartTable"
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub


Private Sub PopulateCombo(cbo As ComboBox, sColumn As String, sTable As String)
   'populate combobox from database table values of a specific column
   
   cbo.Clear
   cbo.AddItem "<ALL>"
   
   Dim rdo As ADODB.Recordset
   sSql = "select " & sColumn & " from " & sTable & " GROUP BY " & sColumn
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If bSqlRows Then
      With rdo
         Do Until .EOF
            If Trim(.Fields(0)) = "" Then
               cbo.AddItem "<BLANK>"
            Else
               cbo.AddItem Trim(.Fields(0))
            End If
            .MoveNext
         Loop
      End With
   End If
   cbo.ListIndex = 0
End Sub

Private Function GetWipRuns() As Boolean
   Dim RdoMos As ADODB.Recordset
   Dim bLots As Byte
   Dim bShowLots As Byte
   
   bLots = CheckLotStatus()
   sSql = "INSERT INTO EsReportWIPOH" & vbCrLf _
      & "(WIPRUNREF,WIPRUNNO,WIPRUNQTY,WIPRUNSTATUS,WIPCOSTTYPE)" & vbCrLf _
      & "select RUNREF,RUNNO,RUNQTY, RUNSTATUS," & vbCrLf _
      & "case when " & bLots & " = 1 and PALOTTRACK = 1 then 'LOT' else 'STD' end" & vbCrLf _
      & "from RunsTable" & vbCrLf _
      & "join PartTable on RUNREF=PARTREF " & vbCrLf _
      & "where (RUNCOMPLETE IS NULL OR RUNCOMPLETE > '" & cboAsOf & "') " & vbCrLf _
      & " AND RUNCREATE <  '" & cboAsOf & "' " & vbCrLf _
      & " and RUNSTATUS<>'CA'" & vbCrLf _
      & "union " & vbCrLf _
      & "select RUNREF,RUNNO,RUNQTY,RUNSTATUS," & vbCrLf _
      & "case when " & bLots & " = 1 and PALOTTRACK = 1 then 'LOT' else 'STD' end" & vbCrLf _
      & "from RunsTable" & vbCrLf _
      & "join PartTable on RUNREF=PARTREF " & vbCrLf _
      & "where (RUNSTATUS = 'CA' AND RUNCANCELED > '" & cboAsOf & "') " & vbCrLf _
      & " AND RUNCREATE <  '" & cboAsOf & "' " & vbCrLf _
      
      ' Add Runs which got canceled and canceled after requested date.
      
  Debug.Print sSql
  
      '"where (RUNCLOSED IS NULL OR RUNCLOSED > '" & cboAsOf & "') "
   clsADOCon.ExecuteSQL sSql
   
   lTotalRuns = clsADOCon.RowsAffected
   
   If lTotalRuns > 0 Then
      GetWipRuns = True
   Else
      GetWipRuns = False
   End If
   lblRuns = lTotalRuns
   lblRuns.Refresh
   
End Function

'Private Sub txtend_DropDown()
'   ShowCalendar Me
'
'End Sub
'
Public Function BuildReport() As Boolean
   'return True if successful
   Dim A As Integer
   Dim cCounter As Currency
   Dim cValue As Currency
   Dim lList As Long
   
   Prg1.Visible = True
   On Error GoTo DiaErr1
   cValue = 100 / lTotalRuns
   A = 5
   Prg1.Value = A
   MouseCursor 13
   
   BatchMarkUnInvoicedPoItems    ' set WIPMISSEXP for missing expense items
   BatchMarkOpenPickList         ' set WIPMISSMATL for open pick list items
   BatchExpenseCosts
   BatchLaborCosts
   GetMaterialCosts              ' set WIPMISSMATL for lot-tracked parts with uncosted lots
                                 ' and non-lot-tracked parts with uncosed picks
   
   ' Calcualte the OH
   CalcOHRate
   
   Prg1.Value = 100
   MouseCursor 0
   
   Prg1.Visible = False
   BuildReport = True
   Exit Function
   
DiaErr1:
   sProcName = "BuildReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub BatchMarkUnInvoicedPoItems()
   'Uninvoiced PO Items marked for expensed, but use as desired
   
   sSql = "update EsReportWIPOH set WIPMISSEXP = 1" & vbCrLf _
      & "where exists( select PINUMBER" & vbCrLf _
      & "from PoitTable" & vbCrLf _
      & "join PartTable on PARTREF = PIPART" & vbCrLf _
      & "where PIRUNPART = WIPRUNREF and PIRUNNO = WIPRUNNO" & vbCrLf _
      & "and PIAQTY = 0 and PITYPE not in (16,17))"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
End Sub

Private Sub BatchMarkOpenPickList()
   'Flag Open Pick items

   sSql = "update EsReportWIPOH set WIPMISSMATL = 1" & vbCrLf _
      & "where exists(select PKTYPE" & vbCrLf _
      & "from MopkTable" & vbCrLf _
      & "where PKMOPART = WIPRUNREF and PKMORUN = WIPRUNNO" & vbCrLf _
      & "and (PKTYPE in (9,23) or (PKTYPE <> 12 and PKAQTY = 0)))"
   clsADOCon.ExecuteSQL sSql

End Sub

Private Sub BatchExpenseCosts()
   
   'Expense Items with and without PO's
   sSql = "update EsReportWIPOH set WIPEXP=" & vbCrLf _
      & "(select isnull(sum(PIAQTY * PIAMT),0)" & vbCrLf _
      & "from PoitTable" & vbCrLf _
      & "join PartTable on PIPART = PARTREF" & vbCrLf _
      & "where PIRUNPART = WIPRUNREF and PIRUNNO = WIPRUNNO" & vbCrLf _
      & "and PITYPE = 17 and PALEVEL = 7 and " & "PIADATE <= '" & cboAsOf & "')" & vbCrLf _
      & "+ (select isnull(sum(VITQTY*VITCOST),0)" & vbCrLf _
      & "from ViitTable where VITMO = WIPRUNREF and VITMORUN = WIPRUNNO" & vbCrLf _
      & "and VITPO = 0 and VITPOITEM = 0)"
   
Debug.Print sSql
   clsADOCon.ExecuteSQL sSql
   
   'Tax and freight
   sSql = "update EsReportWIPOH set WIPFREIGHT = " & vbCrLf _
      & "(select isnull(sum(VIFREIGHT),0)" & vbCrLf _
      & "from ViitTable" & vbCrLf _
      & "join VihdTable on VITNO = VINO" & vbCrLf _
      & "where VITMO = WIPRUNREF and VITMORUN = WIPRUNNO)," & vbCrLf _
      & "WIPTAX = " & vbCrLf _
      & "(select isnull(sum(VITAX),0)" & vbCrLf _
      & "from ViitTable" & vbCrLf _
      & "join VihdTable on VITNO = VINO" & vbCrLf _
      & "where VITMO = WIPRUNREF and VITMORUN = WIPRUNNO)"
Debug.Print sSql
   clsADOCon.ExecuteSQL sSql
   
End Sub
Public Sub BatchLaborCosts()
   'Labor and Overhead
   
   'get labor and overhead
   sSql = "update EsReportWIPOH set WIPLABOR = " & vbCrLf _
      & "(select cast(isnull(sum(TCHOURS * TCRATE),0) as decimal(12,2)) as Labor" & vbCrLf _
      & "from TcitTable" & vbCrLf _
      & "join TchdTable on TCCARD = TMCARD" & vbCrLf _
      & "where TCPARTREF = WIPRUNREF and TCRUNNO = WIPRUNNO" & vbCrLf _
      & "and TMDAY <= '" & cboAsOf & "')," & vbCrLf _
      & "WIPOH = (select cast(isnull(sum(TCHOURS * (TCOHRATE + TCOHFIXED)),0)" & vbCrLf _
      & "as decimal(12,2)) as Overhead" & vbCrLf _
      & "from TcitTable" & vbCrLf _
      & "join TchdTable on TCCARD = TMCARD" & vbCrLf _
      & "where TCPARTREF = WIPRUNREF and TCRUNNO = WIPRUNNO" & vbCrLf _
      & "and TMDAY <= '" & cboAsOf & "')"
   
   clsADOCon.ExecuteSQL sSql
   
   sSql = "update EsReportWIPOH set WIPMISSTIME = 1" & vbCrLf _
      & "where exists (select TCRATE" & vbCrLf _
      & "from TcitTable" & vbCrLf _
      & "join TchdTable on TCCARD = TMCARD" & vbCrLf _
      & "where TCPARTREF = WIPRUNREF and TCRUNNO = WIPRUNNO" & vbCrLf _
      & "and TMDAY <= '" & cboAsOf & "'" & vbCrLf _
      & "and (TCRATE = 0 Or (TCOHRATE = 0 And TCOHFIXED = 0)))"
   
   clsADOCon.ExecuteSQL sSql

End Sub

Private Sub GetMaterialCosts()
   
   'Get costed Lots (Picks) less canceled picks
   ' Added type 16 and 11
   sSql = "update EsReportWIPOH set WIPMATL = " & vbCrLf _
      & "(select cast(isnull(sum(-LOIQUANTITY * LOTUNITCOST),0) as decimal(12,2))" & vbCrLf _
      & "from LoitTable" & vbCrLf _
      & "join LohdTable on LOINUMBER = LOTNUMBER " & vbCrLf _
      & "INNER JOIN PartTable ON LoitTable.LOIPARTREF = PartTable.PARTREF " & vbCrLf _
      & "AND LohdTable.LOTPARTREF = PartTable.PARTREF " & vbCrLf _
      & "LEFT OUTER JOIN PoitTable ON LohdTable.LOTPOITEMREV = PoitTable.PIREV AND " & vbCrLf _
      & "LohdTable.LOTPO = PoitTable.PINUMBER " & vbCrLf _
      & "AND LohdTable.LOTPOITEM = PoitTable.PIITEM " & vbCrLf _
      & "where LOIMOPARTREF = WIPRUNREF and LOIMORUNNO = WIPRUNNO" & vbCrLf _
      & "and LOITYPE in (10, 12, 21, 16,11) and LOTUNITCOST > 0" & vbCrLf _
      & "And PALEVEL <> 7" & vbCrLf _
      & "and LOIADATE < dateadd(d, 1, '" & cboAsOf & "'))" & vbCrLf _
      & "where WIPCOSTTYPE = 'LOT'"

Debug.Print sSql

   clsADOCon.ExecuteSQL sSql
   
   'WIP cost type = STD
   sSql = "update EsReportWIPOH set WIPMATL = " & vbCrLf _
      & "(select cast(isnull(sum(PKAQTY * PKAMT),0) as decimal(12,2))" & vbCrLf _
      & "from MopkTable" & vbCrLf _
      & "join PartTable on PKPARTREF = PARTREF" & vbCrLf _
      & "where PKAQTY > 0" & vbCrLf _
      & "and PKMOPART = WIPRUNREF and PKMORUN = WIPRUNNO" & vbCrLf _
      & "and PKPDATE <= '" & cboAsOf & "')" & vbCrLf _
      & "where WIPCOSTTYPE = 'STD'"
   
Debug.Print sSql
   
   clsADOCon.ExecuteSQL sSql
   
   'mark MO's for lot-tracked parts that have uncosted lots
   sSql = "update EsReportWIPOH" & vbCrLf _
      & "set WIPMISSMATL = 1," & vbCrLf _
      & "WIPUNCOSTED = 1" & vbCrLf _
      & "where exists (select LOINUMBER from LoitTable" & vbCrLf _
      & "join LohdTable on LOINUMBER = LOTNUMBER" & vbCrLf _
      & "where LOIMOPARTREF = WIPRUNREF and LOIMORUNNO = WIPRUNNO" & vbCrLf _
      & "and LOITYPE in (10, 12) and LOTUNITCOST = 0)"
   
   clsADOCon.ExecuteSQL sSql

End Sub

Private Sub CalcOHRate()
   sSql = "WIPCalculatedOH"
   clsADOCon.ExecuteSQL sSql

End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(cboAsOf.Text)
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub


Private Sub GetOptions()
   Dim sOptions As String
   Dim dToday As Integer
   
   ' Get By Menu Option
   On Error Resume Next
   dToday = CInt(Mid(Format(Now, "mm/dd/yy"), 4, 2))
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
     If dToday >= 20 Then
       cboAsOf = GetMonthEnd(Format(Now, "mm/dd/yy"))
     Else
'       txtDte = GetMonthEnd(Format(DateAdd("m", -1, Now), "mm/dd/yy"))
        cboAsOf = Mid(sOptions, 1, 8)
     End If

   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub


Private Sub optDis_Click()
   If DataReady() Then
      PrintReport
   End If
End Sub

Private Sub optPrn_Click()
   If DataReady() Then
      PrintReport
   End If
End Sub

Private Sub PrintReport()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   sProcName = "printreport"
'   SetMdiReportsize MdiSect, True
  Set cCRViewer = New EsCrystalRptViewer
  cCRViewer.Init
  sCustomReport = GetCustomReport("finwipOH.rpt")
  cCRViewer.SetReportTitle = sCustomReport
  cCRViewer.SetReportFileName sCustomReport, sReportPath
'   MdiSect.crw.ReportFileName = sReportPath & sCustomReport

   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Title3"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'Work In Process Inventory As Of " & CStr(cboAsOf) & "'")
   aFormulaValue.Add CStr("'Part Class " & CStr(cboClass _
                        & " And Product Code " & cboCode) & "'")
  
  cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = ""
   If UCase(cboClass) = "<BLANK>" Then
      sSql = "{PartTable.PACLASS} = ''"
   ElseIf UCase(cboClass) <> "<ALL>" Then
      sSql = "{PartTable.PACLASS} = '" & Compress(cboClass) & "'"
   End If

   If UCase(cboCode) <> "<ALL>" Then
      If Len(sSql) > 0 Then
         sSql = sSql & " and "
      End If
      If cboCode = "<BLANK>" Then
         sSql = sSql & "{PartTable.PAPRODCODE} = ''"
      Else
         sSql = sSql & "{PartTable.PAPRODCODE} = '" & Compress(cboCode) & "'"
      End If
   End If
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.CRViewerSize Me
    cCRViewer.SetDbTableConnection
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

