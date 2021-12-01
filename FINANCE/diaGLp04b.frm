VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLp04b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detailed General Ledger (Report)"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkShowUnposted 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   3300
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5400
      TabIndex        =   23
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Picture         =   "diaGLp04b.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Picture         =   "diaGLp04b.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6120
      Top             =   3240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4290
      FormDesignWidth =   6720
   End
   Begin VB.ComboBox cmbEndAct 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Tag             =   "3"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ComboBox cmbStartAct 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CheckBox chkPageBreaks 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   3600
      Width           =   735
   End
   Begin VB.ComboBox cboEndDate 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Tag             =   "4"
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cboStartDate 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Tag             =   "4"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CheckBox chkShowInactive 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5520
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaGLp04b.frx":0308
      PictureDn       =   "diaGLp04b.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   17
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
      PictureUp       =   "diaGLp04b.frx":0594
      PictureDn       =   "diaGLp04b.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Unposted Entries"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   24
      Top             =   3300
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Inactive Accounts"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   22
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page Break By Account"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   21
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblEndAct 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   20
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label lblStartAct 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   16
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   15
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Account"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Beginning Account"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Ending"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Beginning"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "diaGLp04b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

' See the UpdateTables prodecure for database revisions

'*************************************************************************************
'
' diaGLp04b - Trial Balance (Report) / Detail GL (Report)
'
' Notes: Same form used for both reports.
'
' Created: 03/20/01 (nth)
' Revisions:
' 09/17/03 (nth) Added beginning balance to Detail GL per WCK.
' 09/17/03 (nth) Revised and updated trial balance.
' 01/19/05 (nth) Corrected beginning and ending account filter.
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte

' bForm documentation
' 0 = Detail GL
' 1 = Trial Balance
Dim bForm As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


'Private Sub CreateActTable()
'    Dim NewTb1 As TableDef
'    Dim NewTb2 As TableDef
'    Dim NewIdx1 As Index
'    Dim NewIdx2 As Index
'
'    On Error Resume Next
'    JetDb.Execute "DROP TABLE AccountActivity"
'    JetDb.Execute "DROP TABLE AccountBalance"
'
'    Set NewTb1 = JetDb.CreateTableDef("AccountActivity")
'    Set NewTb2 = JetDb.CreateTableDef("AccountBalance")
'
'    With NewTb1
'        .Fields.Append .CreateField("ActRef", dbText, 12)
'        .Fields.Append .CreateField("ActDeb", dbDouble)
'        .Fields.Append .CreateField("ActCrd", dbDouble)
'        .Fields.Append .CreateField("ActJETran", dbInteger)
'        .Fields.Append .CreateField("ActJERef", dbInteger)
'        .Fields.Append .CreateField("ActJE", dbText, 12)
'        .Fields.Append .CreateField("ActJEDesc", dbText, 30)
'        .Fields.Append .CreateField("ActJEPost", dbDate)
'    End With
'
'    JetDb.TableDefs.Append NewTb1
'    With NewTb2
'        .Fields.Append .CreateField("ActNum", dbText, 12)
'        .Fields.Append .CreateField("ActRef", dbText, 12)
'        .Fields.Append .CreateField("ActDesc", dbText, 40)
'        .Fields.Append .CreateField("ActBal", dbDouble)
'    End With
'
'    JetDb.TableDefs.Append NewTb2
'
'    Set NewTb1 = Nothing
'    Set NewTb2 = Nothing
'
'    'add the table and indexes to Jet.
'
'    On Error Resume Next
'    Set NewTb1 = JetDb!AccountActivity
'        With NewTb1
'            Set NewIdx1 = .CreateIndex
'                With NewIdx1
'                    .Name = "ixaActNum"
'                    .Fields.Append .CreateField("ActRef")
'                End With
'                .Indexes.Append NewIdx1
'        End With
'
'     Set NewTb2 = JetDb!AccountBalance
'        With NewTb2
'            Set NewIdx2 = .CreateIndex
'                With NewIdx2
'                    .Name = "ixbActNum"
'                    .Unique = True
'                    .Fields.Append .CreateField("ActRef")
'                End With
'               .Indexes.Append NewIdx2
'        End With
'
'    Set NewTb1 = Nothing
'    Set NewTb2 = Nothing
'    Set NewIdx2 = Nothing
'    Set NewIdx1 = Nothing
'    Exit Sub
'
'DiaErr1:
'    sProcName = "CreateActTable"
'    CurrError.Number = Err.Number
'    CurrError.description = Err.description
'    DoModuleErrors Me
'End Sub
'

Private Sub BuildAccountTotals()
   'Dim DbActivity As Recordset
   Dim DbBal As Recordset
   'Dim RdoBal As ADODB.Recordset
   'Dim rdoAct As ADODB.Recordset
   'Dim RdoSum As ADODB.Recordset
   Dim sBegDate As String
   Dim sEnddate As String
   Dim sBegAct As String
   Dim sEndAct As String
   Dim sTemp As String
   Dim iCount As Integer
   
   'Check for valid date entries
   If cboStartDate = "" Then
      MsgBox "Please Enter A Valid Starting Date.", vbInformation
      cboStartDate.SetFocus
      Exit Sub
   ElseIf cboEndDate = "" Then
      MsgBox "Please Enter A Valid Ending Date.", vbInformation
      cboEndDate.SetFocus
      Exit Sub
   End If
   
   '    MouseCursor 13
   '
   '    sBegDate = cboStartDate
   '    sEndDate = cboEndDate
   '
   '    sBegAct = Compress(cmbStartAct)
   '    sEndAct = Compress(cmbEndAct)
   '
   '    'If sBegAct = "" Then sBegAct = Trim(cmbStartAct.List(0))
   '    'If sEndAct = "" Then sEndAct = Trim(cmbEndAct.List(cmbEndAct.ListCount - 1))
   '
   '    On Error Resume Next
   '    'ReopenJet
   '    CreateActTable
   '
   '    JetDb.Execute "DELETE * FROM AccountActivity"
   '    JetDb.Execute "DELETE * FROM AccountBalance"
   '
   '    Set DbActivity = JetDb.OpenRecordset("AccountActivity", dbOpenDynaset)
   '    Set DbBal = JetDb.OpenRecordset("AccountBalance", dbOpenDynaset)
   '
   '    sSql = "SELECT SUM(JIDEB) AS Debit, SUM(JICRD) " _
   '        & "AS Credit,GJPOST,JIACCOUNT,GJNAME,JIDESC,JITRAN,JIREF " _
   '        & "FROM GjhdTable INNER JOIN GjitTable ON GJNAME=JINAME " _
   '        & "WHERE (GJPOSTED=1) " _
   '        & "GROUP BY GJNAME,JIACCOUNT,GJPOST,JIDESC,JITRAN,JIREF " _
   '        & "HAVING (GJPOST >= '" & sBegDate & "' AND GJPOST <= '" & sEndDate & "')"
   '    'If sBegAct <> "" Or sEndAct <> "" Then
   '    '    sSql = sSql & " and isnumeric(jiaccount)=1 "
   '    'End If
   '    If sBegAct <> "" Then
   '        sSql = sSql & "AND JIACCOUNT  >= '" & sBegAct & "'"
   '    End If
   '    If sEndAct <> "" Then
   '        sSql = sSql & " AND JIACCOUNT <= '" & sEndAct & "'"
   '    End If
   '    bSqlRows = clsAdoCon.GetDataSet(sSql,rdoAct, ES_FORWARD)
   '
   '    ' Dump results to temp jet database
   '    With rdoAct
   '        Do While Not .EOF
   '            DbActivity.AddNew
   '            DbActivity!ActRef = !JIACCOUNT
   '            DbActivity!ActDeb = !debit
   '            DbActivity!ActCrd = !credit
   '            DbActivity!ActJEPost = !GJPOST
   '            DbActivity!ActJE = !GJNAME
   '            DbActivity!ActJEDesc = !JIDESC
   '            DbActivity!ActJETran = !JITRAN
   '            DbActivity!ActJERef = !JIREF
   '            DbActivity.Update
   '            .MoveNext
   '        Loop
   '    End With
   '    DbActivity.Close
   '    Set rdoAct = Nothing
   '
   '    sSql = "SELECT DISTINCT GLACCTREF,GLACCTNO,GLDESCR FROM GlacTable "
   '    If Not chkShowInactive Then
   '        sTemp = sTemp & "(GLINACTIVE=0)"
   '    End If
   '    If sBegAct <> "" Then
   '        sTemp = sTemp & " AND(GLACCTREF>='" & sBegAct & "')"
   '    End If
   '    If sEndAct <> "" Then
   '        sTemp = sTemp & " AND(GLACCTREF<='" & sEndAct & "')"
   '    End If
   '    If Len(sTemp) Then
   '        sSql = sSql & " WHERE " & sTemp
   '    End If
   '
   '    bSqlRows = clsAdoCon.GetDataSet(sSql,RdoBal, ES_FORWARD)
   '    With RdoBal
   '        Do While Not .EOF
   '            iCount = iCount + 1
   '
   '            DbBal.AddNew
   '            DbBal!ActNum = !GLACCTNO
   '            DbBal!ActRef = !GLACCTREF
   '            DbBal!ActDesc = !GLDESCR
   '
   '            sSql = "SELECT SUM(GjitTable.JIDEB) AS Debit, " _
   '                & "SUM(GjitTable.JICRD) AS Credit " _
   '                & "FROM GjhdTable INNER JOIN " _
   '                & "GjitTable ON GJNAME = JINAME " _
   '                & "WHERE (JIACCOUNT = '" & !GLACCTREF _
   '                & "') AND (GJPOSTED = 1) AND (GJPOST < '" _
   '                & sBegDate & "')"
   '            bSqlRows = clsAdoCon.GetDataSet(sSql,RdoSum, ES_FORWARD)
   '
   '            DbBal!ActBal = (RdoSum!debit - RdoSum!credit)
   '            Set RdoSum = Nothing
   '            DbBal.Update
   '            .MoveNext
   '        Loop
   '    End With
   '    DbBal.Close
   '    Set RdoBal = Nothing
   '
   '    On Error Resume Next
   '    JetDb.Close
   PrintReport
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "BuildAccountTotals"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub PrintReport()
    Dim sWindows As String
    Dim sTemp As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   On Error GoTo DiaErr1
   
   If InStr(Caption, "Trial Balance") > 0 Then
      sTemp = "fingl05b.rpt"
   Else
      sTemp = "fingl04b.rpt"
      aFormulaName.Add "PrePeriod"
      aFormulaValue.Add CStr("'" & CStr(Format(DateAdd("d", -1, CDate(cboStartDate)), "m/d/yy")) & "'")
   End If
   
   sCustomReport = GetCustomReport(sTemp)
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "StartDate"
    aFormulaName.Add "EndDate"
    aFormulaName.Add "StartAccount"
    aFormulaName.Add "EndAccount"
    aFormulaName.Add "ShowInactive"
    aFormulaName.Add "ShowUnposted"
    aFormulaName.Add "PageBreaks"
    aFormulaName.Add "Title1"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    aFormulaValue.Add CStr("'" & CStr(cboStartDate) & "'")
    aFormulaValue.Add CStr("'" & CStr(cboEndDate) & "'")
    aFormulaValue.Add CStr("'" & CStr(cmbStartAct) & "'")
    aFormulaValue.Add CStr("'" & CStr(cmbEndAct) & "'")
    aFormulaValue.Add chkShowInactive.Value
    aFormulaValue.Add chkShowUnposted.Value
    aFormulaValue.Add chkPageBreaks.Value
    aFormulaValue.Add CStr("'Period Beginning " & CStr(Format(cboStartDate, "m/d/yy") _
                        & " And Ending " & Format(cboEndDate, "m/d/yy")) & "'")
   
    aFormulaName.Add "Title2"
   
   If chkShowInactive Then
     aFormulaValue.Add CStr("'Include Inactive Accounts? " & y & "'")
   Else
     aFormulaValue.Add CStr("'Include Inactive Accounts? " & "N" & "'")
   End If
    sSql = "{GlacTable.GLACCTNO} in {@StartAccount} to {@EndAccount}"
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   
   Exit Sub
   
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sWindows As String
   Dim sTemp As String
   Dim sCustomReport As String
   
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   'ReopenJet
   
   'sWindows = GetWindowsDir()
   'MdiSect.crw.DataFiles(0) = sWindows & "\temp\esifina.mdb"
   
   If InStr(Caption, "Trial Balance") > 0 Then
      sTemp = "fingl05b.rpt"
   Else
      sTemp = "fingl04b.rpt"
      MdiSect.crw.Formulas(6) = "PrePeriod='" _
                           & Format(DateAdd("d", -1, CDate(cboStartDate)), "m/d/yy") & "'"
   End If
   
   sCustomReport = GetCustomReport(sTemp)
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " _
                        & sInitials & "'"
   MdiSect.crw.Formulas(2) = "StartDate='" & cboStartDate & "'"
   MdiSect.crw.Formulas(3) = "EndDate='" & cboEndDate & "'"
   '    'MdiSect.crw.Formulas(1) = "StartDate=Date(" & Format(txtBeg, "yyyy,mm,dd") & ")"
   '    MdiSect.crw.Formulas(2) = "StartDate=Date(" & Format(cboStartDate, "yyyy,mm,dd") & ")"
   '    'MdiSect.crw.Formulas(2) = "StartDate=" & "Date(2005,1,1)"
   '    MdiSect.crw.Formulas(3) = "EndDate=" & "Date(2005,12,31)"
   
   '    If Trim(cmbStartAct) = "" Then
   '        MdiSect.crw.Formulas(4) = "StartAccount='" & cmbStartAct.List(0) & "'"
   '    Else
   MdiSect.crw.Formulas(4) = "StartAccount='" & cmbStartAct & "'"
   '    End If
   '
   '    If Trim(cmbEndAct) = "" Then
   '        MdiSect.crw.Formulas(5) = "EndAccount='" & cmbEndAct.List(cmbEndAct.ListCount - 1) & "'"
   '    Else
   MdiSect.crw.Formulas(5) = "EndAccount='" & cmbEndAct & "'"
   '    End If
   MdiSect.crw.Formulas(6) = "ShowInactive=" & chkShowInactive.Value
   MdiSect.crw.Formulas(7) = "ShowUnposted=" & chkShowUnposted.Value
   MdiSect.crw.Formulas(8) = "PageBreaks=" & chkPageBreaks.Value
   
   MdiSect.crw.Formulas(9) = "Title1='Period Beginning " _
                        & Format(cboStartDate, "m/d/yy") _
                        & " And Ending " & Format(cboEndDate, "m/d/yy") & "'"
   
   sTemp = "Title2='Include Inactive Accounts? "
   If chkShowInactive Then
      sTemp = sTemp & "Y'"
   Else
      sTemp = sTemp & "N'"
   End If
   MdiSect.crw.Formulas(10) = sTemp
   
   'SetCrystalAction Me
   Exit Sub
   
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillAccounts()
   ' Fill account combo
   Dim gl As New GLTransaction
   gl.FillComboWithAccounts cmbStartAct, cmbEndAct
   
   
   '    Dim rdoAct As ADODB.RecordSet
   '    On Error GoTo DiaErr1
   '    MouseCursor 13
   '    sSql = "Qry_FillLowAccounts"
   '    bSqlRows = clsAdoCon.GetDataSet(sSql,rdoAct, ES_FORWARD)
   '
   '    If bSqlRows Then
   '        With rdoAct
   '            Do Until .EOF
   '                AddComboStr cmbStartAct.hWnd, "" & Trim(!GLACCTNO)
   '                AddComboStr cmbEndAct.hWnd, "" & Trim(!GLACCTNO)
   '                .MoveNext
   '            Loop
   '        End With
   '    End If
   '    Set rdoAct = Nothing
   '
   '    lblEndAct = UpdateActDesc(cmbEndAct)
   '    lblStartAct = UpdateActDesc(cmbStartAct)
   '
   '    MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "FillAcounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbEndAct_Click()
   lblEndAct = UpdateActDesc(cmbEndAct)
End Sub

Private Sub cmbEndAct_LostFocus()
   lblEndAct = UpdateActDesc(cmbEndAct)
End Sub

Private Sub cmbStartAct_Click()
   lblStartAct = UpdateActDesc(cmbStartAct)
End Sub

Private Sub cmbStartAct_LostFocus()
   lblStartAct = UpdateActDesc(cmbStartAct)
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      If Me.Caption = "Trial Balance (Report)" Then
         chkPageBreaks.Visible = False
         z1(6).Visible = False
      Else
         chkPageBreaks.Visible = True
         z1(6).Visible = True
      End If
      FillAccounts
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   cboStartDate = Format(Now, "mm/01/yy")
   cboEndDate = GetMonthEnd(cboStartDate)
   GetOptions
   'ReopenJet
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   On Error Resume Next
   '    JetDb.Execute "DROP TABLE AccountActivity"
   '    JetDb.Execute "DROP TABLE AccountBalance"
   Set diaGLp04a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub optDis_Click()
   BuildAccountTotals
End Sub

Private Sub optPrn_Click()
   BuildAccountTotals
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub cboEndDate_DropDown()
   ShowCalendar Me
End Sub

Private Sub cboEndDate_LostFocus()
   cboEndDate = CheckDate(cboEndDate)
End Sub

Private Sub cbostartdate_DropDown()
   ShowCalendar Me
End Sub

Private Sub cbostartdate_LostFocus()
   cboStartDate = CheckDate(cboStartDate)
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = Trim(str(chkPageBreaks.Value) & (str(chkShowInactive.Value)))
   SaveSetting "Esi2000", "EsiFina", Me.Name & bForm, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & bForm & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name & bForm, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      chkPageBreaks = Val(Left(sOptions, 1))
      chkShowInactive = Val(Right(sOptions, 1))
   Else
      chkPageBreaks = vbUnchecked
      chkShowInactive = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & bForm _
                & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub
