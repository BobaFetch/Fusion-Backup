VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRp11a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Outside Services Requirements"
   ClientHeight    =   3135
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRp11a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optSel 
      Caption         =   "No Purchase Order Issued"
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   5
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CheckBox optSel 
      Caption         =   "Not Received"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   4
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox optSel 
      Caption         =   "Received"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Contains Used Part Type 7's. Leading Character Search Or Select Contains Part Numbers On Purchase Orders"
      Top             =   1080
      Width           =   3545
   End
   Begin VB.CheckBox optDet 
      Caption         =   "PO Item Comments"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1680
      TabIndex        =   6
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PurcPRp11a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "PurcPRp11a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   2880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3135
      FormDesignWidth =   7215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   3
      Left            =   5640
      TabIndex        =   18
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "OR"
      Height          =   285
      Index           =   7
      Left            =   3120
      TabIndex        =   17
      Top             =   1920
      Width           =   465
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show:"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   16
      Top             =   1800
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Op Sched From"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   4
      Left            =   3000
      TabIndex        =   14
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   5640
      TabIndex        =   13
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Tag             =   " "
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Part(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "PurcPRp11a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT OPSERVPART,PARTREF,PARTNUM FROM " _
          & "RnopTable,PartTable WHERE OPSERVPART=PARTREF AND PALEVEL=7 " _
          & "ORDER BY PARTREF"
   LoadComboBox cmbPrt, 1
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Trim(cmbPrt) = "" Then cmbPrt = "ALL"
   
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



Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad = 1 Then
      FillCombo
      CreateTable
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
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
   Set ShopSHp08a = Nothing
   
End Sub

Private Sub PrintReport()
    Dim bOpts(3) As Byte
    Dim sBeg As String
    Dim sEnd As String
    Dim sPart As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   If optSel(2).Value = vbChecked Then
      bOpts(2) = 1
      bOpts(0) = 0
      bOpts(1) = 0
   Else
      bOpts(2) = 0
      bOpts(0) = optSel(0).Value
      bOpts(1) = optSel(1).Value
   End If
   
   MouseCursor 13
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Not IsDate(txtBeg) Then
      sBeg = "1995,01,01"
   Else
      sBeg = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEnd = "2024,12,31"
   Else
      sEnd = Format(txtEnd, "yyyy,mm,dd")
   End If
   If Trim(cmbPrt) = "" Then cmbPrt = "ALL"
   If cmbPrt <> "ALL" Then sPart = Compress(cmbPrt)
   
   On Error GoTo DiaErr1

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowDetails"
   
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Part Number(s)" & CStr(cmbPrt & "... " _
                        & txtBeg & " Though " & txtEnd) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add optDet.Value
    
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   If bOpts(2) = 1 Then
      sCustomReport = GetCustomReport("prdpr13a") 'Without PO's
   Else
      sCustomReport = GetCustomReport("prdpr13b") 'With PO's
   End If
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
   sSql = "{PartTable.PARTREF} LIKE '" & sPart & "*' AND " _
          & "{RnopTable.OPSCHEDDATE} in Date(" & sBeg & ") to Date(" & sEnd & ") "
   sSql = sSql & " AND {PartTable.PALEVEL} = 7"
   If bOpts(2) = 0 Then
      If bOpts(0) = 1 Then sSql = sSql & "AND {PoitTable.PIAQTY}>0 "
      If bOpts(1) = 1 Then sSql = sSql & "AND {PoitTable.PIAQTY}=0 "

'      If optDet.value = 1 Then
'         MDISect.Crw.SectionFormat(0) = "GROUPFTR.0.0;T;;;"
'      Else
'         MDISect.Crw.SectionFormat(0) = "GROUPFTR.0.0;F;;;"
'      End If
   End If

'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
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
   cmbPrt = "ALL"
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 4)
   
End Sub

Private Sub SaveOptions()
   
End Sub

Private Sub GetOptions()
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   If optSel(2).Value = vbChecked Then FillTable
   PrintReport
   
End Sub


Private Sub optPrn_Click()
   If optSel(2).Value = vbChecked Then FillTable
   PrintReport
   
End Sub





Private Sub CreateTable()
   Dim RdoTst As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT OPREF FROM dbo.EsReportServOps WHERE " _
          & "OPREF='FOOBAR' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTst, ES_FORWARD)
   If Err = 40002 Then
      sSql = "CREATE TABLE dbo.EsReportServOPs " _
             & "(OPCOUNTER SMALLINT NULL DEFAULT(0)," _
             & "OPREF CHAR(30) NULL DEFAULT('')," _
             & "OPRUN INT NULL DEFAULT(0)," _
             & "OPNO SMALLINT NULL DEFAULT(0)," _
             & "OPSERVPART CHAR(30) NULL DEFAULT(''))"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "CREATE UNIQUE CLUSTERED INDEX OpRef ON " _
             & "dbo.EsReportServOps(OPCOUNTER) WITH FILLFACTOR = 80"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "CREATE INDEX OpMo ON dbo.EsReportServOps(OPREF) WITH FILLFACTOR = 80"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "CREATE INDEX OpSp ON dbo.EsReportServOps(OPSERVPART) WITH FILLFACTOR = 80"
      clsADOCon.ExecuteSQL sSql
   Else
      ClearResultSet RdoTst
   End If
   Set RdoTst = Nothing
   
End Sub

Private Sub FillTable()
   Dim RdoSer As ADODB.Recordset
   Dim iCounter As Integer
   MouseCursor 13
   
   sSql = "truncate table EsReportServOps"
   clsADOCon.ExecuteSQL sSql
   
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT OPREF,OPRUN,OPNO,OPSERVPART FROM " _
          & "RnopTable LEFT JOIN PoitTable ON RnopTable.OPREF=PoitTable.PIRUNPART " _
          & "AND RnopTable.OPRUN=PoitTable.PIRUNNO WHERE OPSERVPART<>'' AND " _
          & "(PoitTable.PIRUNPART Is Null)AND (PoitTable.PIRUNNO Is Null) " _
          & "ORDER BY OPSERVPART,OPREF,OPRUN,OPNO"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSer, ES_FORWARD)
   If bSqlRows Then
      With RdoSer
         Do Until .EOF
            iCounter = iCounter + 1
            sSql = "INSERT INTO EsReportServOps (OPCOUNTER,OPREF,OPRUN,OPNO,OPSERVPART) " _
                   & "VALUES (" & iCounter & ",'" & Trim(!OPREF) & "'," & !OPRUN & "," _
                   & !opNo & ",'" & Trim(!OPSERVPART) & "')"
            clsADOCon.ExecuteSQL sSql
            .MoveNext
            If iCounter > 30000 Then Exit Do
         Loop
         ClearResultSet RdoSer
      End With
   End If
   Set RdoSer = Nothing
   On Error GoTo 0
   Exit Sub
   
DiaErr1:
   sProcName = "filltable"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optSel_Click(Index As Integer)
   If optSel(0).Value = vbChecked Then optSel(2).Value = vbUnchecked
   If optSel(1).Value = vbChecked Then optSel(2).Value = vbUnchecked
   If optSel(0).Value = vbUnchecked And optSel(1).Value = vbUnchecked Then _
             optSel(2).Value = vbChecked
   If optSel(2).Value = vbChecked Then
      optSel(0).Value = vbUnchecked
      optSel(1).Value = vbUnchecked
   End If
   
End Sub

Private Sub optSel_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(txtEnd) > 3 Then
      txtEnd = CheckDateEx(txtEnd)
   Else
      txtEnd = "ALL"
   End If
   
End Sub
