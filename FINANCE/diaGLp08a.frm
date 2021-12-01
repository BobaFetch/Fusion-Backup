VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLp08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balance Sheet (Report)"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbDiv 
      Height          =   315
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   2280
      Width           =   660
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   4080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4215
      FormDesignWidth =   7125
   End
   Begin VB.CheckBox optCon 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optExcWO 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox optIna 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.CheckBox optPY 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtLvl 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Tag             =   "1"
      Top             =   1560
      Width           =   285
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Tag             =   "4"
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton CmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5760
      TabIndex        =   12
      Top             =   240
      Width           =   1335
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
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
      PictureUp       =   "diaGLp08a.frx":0000
      PictureDn       =   "diaGLp08a.frx":0146
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   25
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
      PictureUp       =   "diaGLp08a.frx":028C
      PictureDn       =   "diaGLp08a.frx":03D2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   10
      Left            =   3720
      TabIndex        =   24
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Divisionalized Reports Only)"
      Height          =   285
      Index           =   8
      Left            =   3720
      TabIndex        =   22
      Top             =   3480
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consolidated"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exclude Accounts W/O Divisions"
      Height          =   405
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Inactive Accounts"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Previous Year Difference"
      Height          =   435
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through Detail Level"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(9 For All)"
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   16
      Top             =   1560
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Beginning"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Ending"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaGLp08a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*************************************************************************************
' diaGLp08a - Balance Sheet
'
' Notes:
'
' Created: 03/22/01 (nth)
' Revisions:
'   08/07/03 (nth) Fix errors per WCK.
'   10/15/03 (nth) more revisions per WCK / now ties to MCS / ESI balance sheet
'   02/23/04 (JCW) Divisionalized Report
'   08/16/04 (nth) Added getoptions and saveoptions
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim iStart As Integer
Dim iEnd As Integer

Dim vAccounts(10, 4) As Variant

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub cmdCan_Click()
   Unload Me
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
      FillCombo
      CreateActTable
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   ReopenJet '''new
   GetOptions
   'txtBeg = Format(Now, "mm/01/yy")
   'txtEnd = GetMonthEnd(txtBeg)
   txtLvl = 9
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub cmbDiv_LostFocus()
   On Error Resume Next
   cmbDiv = CheckLen(cmbDiv, 2)
   If Trim(cmbDiv) <> "" And Not bValidElement(cmbDiv) Then
      cmbDiv = ""
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   On Error Resume Next
   JetDb.Execute "DROP TABLE ActrpTable"
   JetDb.Execute "DROP TABLE FSS"
   JetDb.Execute "DROP TABLE Current"
   JetDb.Execute "DROP TABLE Previous"
   JetDb.Execute "DROP TABLE Divisions"
   Set diaGLp08a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   If bDivisionAccounts(iStart, iEnd) Then
      sProcName = "filldivisions"
      FillDivisions Me
   Else
      cmbDiv.enabled = False
   End If
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport()
   Dim sWindows As String
   Dim sCustomReport As String
   
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   ReopenJet
   
   sWindows = GetWindowsDir()
   
   MdiSect.crw.DataFiles(0) = sWindows & "\temp\esifina.mdb"
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Title1='Level " & txtLvl _
                        & " Balance Sheet For The Period Ending " & txtEnd & "'"
   MdiSect.crw.Formulas(2) = "RequestBy='Requested By: " _
                        & sInitials & "'"
   
   MdiSect.crw.Formulas(3) = "nDetailLevel = " & Val(txtLvl)
   
   sCustomReport = GetCustomReport("fingl08.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   sSql = " trim(cstr({Divisions.Acctref})) <> '' "
   
   MdiSect.crw.SelectionFormula = sSql
   
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "PrintReport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub optDis_Click()
   BuildAccounts
End Sub

Private Sub optPrn_Click()
   BuildAccounts
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   'SysPrinters.Show
   'ShowPrinters.Value = True
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub

Private Sub txtLvl_LostFocus()
   If Val(txtLvl) < 1 Or Val(txtLvl) > 9 Then
      txtLvl = 9
   End If
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

Public Sub CreateActTable()
   
   Dim NewTb1 As TableDef
   Dim NewTb2 As TableDef
   Dim NewTb3 As TableDef
   Dim NewTb4 As TableDef
   Dim NewTb5 As TableDef
   Dim NewTb6 As TableDef
   Dim NewTb7 As TableDef
   Dim NewTb8 As TableDef
   Dim NewTb9 As TableDef
   Dim NewTb10 As TableDef
   Dim NewFld As Field
   Dim NewIdx1 As Index
   Dim NewIdx2 As Index
   Dim NewIdx3 As Index
   Dim NewIdx4 As Index
   Dim NewIdx5 As Index
   Dim Newidx6 As Index
   
   ' Create FSS table (Financial Statement Structure)
   On Error Resume Next
   JetDb.Execute "DROP TABLE FSS"
   Set NewTb1 = JetDb.CreateTableDef("FSS")
   With NewTb1
      .Fields.Append .CreateField("MASTERREF", dbInteger)
      .Fields.Append .CreateField("MASTERDESC", dbText, 12)
   End With
   JetDb.TableDefs.Append NewTb1
   Set NewTb1 = Nothing
   
   Set NewTb1 = JetDb!FSS
   With NewTb1
      Set NewIdx1 = .CreateIndex
      With NewIdx1
         .Name = "FSS_INDEX1"
         .Fields.Append .CreateField("MASTERREF")
      End With
      .Indexes.Append NewIdx1
   End With
   Set NewTb1 = Nothing
   Set NewIdx1 = Nothing
   
   ' Create fincial statement layout
   On Error Resume Next
   JetDb.Execute "DROP TABLE ActrpTable"
   Err.Clear
   
   ' Fields. Note that we allow empties
   
   Set NewTb1 = JetDb.CreateTableDef("ActrpTable")
   With NewTb1
      'Type
      .Fields.Append .CreateField("Act00", dbInteger)
      'Level
      .Fields.Append .CreateField("Act01", dbInteger)
      'AcctRef
      .Fields.Append .CreateField("Act02", dbText, 12)
      .Fields(2).AllowZeroLength = True
      'Account Number + Spaces to indent
      .Fields.Append .CreateField("Act03", dbText, 32)
      .Fields(3).AllowZeroLength = True
      'Account Desc
      .Fields.Append .CreateField("Act04", dbText, 60)
      .Fields(4).AllowZeroLength = True
      'Active
      .Fields.Append .CreateField("Act05", dbInteger)
      'GLFSLEVEL
      .Fields.Append .CreateField("Act06", dbInteger)
   End With
   
   ' Add the table and indexes to Jet.
   JetDb.TableDefs.Append NewTb1
   Set NewTb2 = JetDb!ActrpTable
   With NewTb2
      Set NewIdx1 = .CreateIndex
      With NewIdx1
         .Name = "AcctTyp"
         .Fields.Append .CreateField("Act00")
      End With
      .Indexes.Append NewIdx1
      
      Set NewIdx2 = .CreateIndex
      With NewIdx2
         .Name = "AcctNo"
         .Fields.Append .CreateField("Act02")
      End With
      .Indexes.Append NewIdx2
   End With
   ' Create the current account balance table
   JetDb.Execute "DROP TABLE Current"
   ' Fields. Note that we allow empties
   Set NewTb3 = JetDb.CreateTableDef("Current")
   With NewTb3
      ' AcctRef
      .Fields.Append .CreateField("AcctRef", dbText, 12)
      .Fields(0).AllowZeroLength = True
      ' Current Period
      .Fields.Append .CreateField("AcctBal", dbCurrency)
   End With
   ' Add the table and indexes to Jet.
   JetDb.TableDefs.Append NewTb3
   Set NewTb4 = JetDb!Current
   With NewTb4
      Set NewIdx3 = .CreateIndex
      With NewIdx3
         .Name = "AcctNo"
         .Fields.Append .CreateField("AcctRef")
      End With
      .Indexes.Append NewIdx3
   End With
   
   
   
   ' Create the YTD account balance table
   JetDb.Execute "DROP TABLE Previous"
   ' Fields. Note that we allow empties
   Set NewTb5 = JetDb.CreateTableDef("Previous")
   With NewTb5
      ' AcctRef
      .Fields.Append .CreateField("AcctRef", dbText, 12)
      .Fields(0).AllowZeroLength = True
      ' Current Period
      .Fields.Append .CreateField("AcctBal", dbCurrency)
   End With
   
   
   
   ' Add the table and indexes to Jet.
   JetDb.TableDefs.Append NewTb5
   Set NewTb6 = JetDb!Previous
   With NewTb6
      Set NewIdx4 = .CreateIndex
      With NewIdx4
         .Name = "AcctNo"
         .Fields.Append .CreateField("AcctRef")
      End With
      .Indexes.Append NewIdx4
   End With
   
   
   
   ' Create the Divisions
   JetDb.Execute "DROP TABLE Divisions"
   ' Fields. Note that we allow empties
   
   Set NewTb9 = JetDb.CreateTableDef("Divisions")
   With NewTb9
      ' AcctRef
      .Fields.Append .CreateField("AcctRef", dbText, 12)
      .Fields(0).AllowZeroLength = True
      '*        ' CurrentIncome Period
      '*       .Fields.Append .CreateField("AcctBal", dbCurrency)
   End With
   ' Add the table and indexes to Jet.
   
   JetDb.TableDefs.Append NewTb9
   Set NewTb10 = JetDb!Divisions
   With NewTb10
      Set Newidx6 = .CreateIndex
      Newidx6.Name = "AcctNo"
      Newidx6.Fields.Append Newidx6.CreateField("AcctRef")
      .Indexes.Append Newidx6
   End With
   
   
   
End Sub

Public Sub BuildAccounts()
   Dim i As Integer
   Dim RdoGlm As ADODB.Recordset
   Dim RdoAct1 As ADODB.Recordset
   Dim RdoAct2 As ADODB.Recordset
   Dim RdoAct4 As ADODB.Recordset
   Dim DbBal1 As Recordset
   Dim DbBal2 As Recordset
   Dim DbBal4 As Recordset
   Dim sSqlAdder As String
   
   MouseCursor 13
   bChart = 1
   iLevel = 9 'Val(txtLvl)
   
   If Trim(cmbDiv) <> "" Then
      sSqlAdder = " WHERE  (RIGHT(LEFT(GLACCTNO + '            ', " _
                  & iEnd & "), " & iEnd & " - (" & iStart & " - 1)) = '" & cmbDiv & "')"
   End If
   
   ' Build balance sheet account structure
   On Error GoTo DiaErr1
   sSql = "DELETE FROM FSS"
   JetDb.Execute sSql
   
   sSql = "SELECT COASSTACCT,COLIABACCT,COEQTYACCT,COINCMACCT,COCOGSACCT," _
          & "COEXPNACCT,COOINCACCT,COOEXPACCT,COFDTXACCT FROM GlmsTable " _
          & "WHERE COACCTREC=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm)
   If bSqlRows Then
      With RdoGlm
         Set DbBal1 = JetDb.OpenRecordset("FSS", dbOpenDynaset)
         For i = 0 To 8
            DbBal1.AddNew
            DbBal1!MASTERREF = i + 1
            DbBal1!MASTERDESC = Trim(.Fields(i))
            DbBal1.Update
         Next
      End With
   End If
   Set RdoGlm = Nothing
   
   sSql = "SELECT * FROM GlmsTable WHERE COACCTREC=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm)
   If bSqlRows Then
      With RdoGlm
         i = 1
         vAccounts(i, 0) = "" & Trim(!COASSTREF)
         vAccounts(i, 1) = "" & Trim(!COASSTACCT)
         vAccounts(i, 2) = "" & Trim(!COASSTDESC)
         vAccounts(i, 3) = Format(!COASSTTYPE, "0")
         
         i = 2
         vAccounts(i, 0) = "" & Trim(!COLIABREF)
         vAccounts(i, 1) = "" & Trim(!COLIABACCT)
         vAccounts(i, 2) = "" & Trim(!COLIABDESC)
         vAccounts(i, 3) = Format(!COLIABTYPE, "0")
         
         i = 3
         vAccounts(i, 0) = "" & Trim(!COEQTYREF)
         vAccounts(i, 1) = "" & Trim(!COEQTYACCT)
         vAccounts(i, 2) = "" & Trim(!COEQTYDESC)
         vAccounts(i, 3) = Format(!COEQTYTYPE, "0")
         
         i = 4
         vAccounts(i, 0) = "" & Trim(!COINCMREF)
         vAccounts(i, 1) = "" & Trim(!COINCMACCT)
         vAccounts(i, 2) = "" & Trim(!COINCMDESC)
         vAccounts(i, 3) = Format(!COINCMTYPE, "0")
         
         i = 5
         vAccounts(i, 0) = "" & Trim(!COCOGSREF)
         vAccounts(i, 1) = "" & Trim(!COCOGSACCT)
         vAccounts(i, 2) = "" & Trim(!COCOGSDESC)
         vAccounts(i, 3) = Format(!COCOGSTYPE, "0")
         
         i = 6
         vAccounts(i, 0) = "" & Trim(!COEXPNREF)
         vAccounts(i, 1) = "" & Trim(!COEXPNACCT)
         vAccounts(i, 2) = "" & Trim(!COEXPNDESC)
         vAccounts(i, 3) = Format(!COEXPNTYPE, "0")
         
         i = 7
         vAccounts(i, 0) = "" & Trim(!COOINCREF)
         vAccounts(i, 1) = "" & Trim(!COOINCACCT)
         vAccounts(i, 2) = "" & Trim(!COOINCDESC)
         vAccounts(i, 3) = Format(!COOINCTYPE, "0")
         
         i = 8
         vAccounts(i, 0) = "" & Trim(!COOEXPREF)
         vAccounts(i, 1) = "" & Trim(!COOEXPACCT)
         vAccounts(i, 2) = "" & Trim(!COOEXPDESC)
         vAccounts(i, 3) = Format(!COOEXPTYPE, "0")
         
         i = 9
         vAccounts(i, 0) = "" & Trim(!COFDTXREF)
         vAccounts(i, 1) = "" & Trim(!COFDTXACCT)
         vAccounts(i, 2) = "" & Trim(!COFDTXDESC)
         vAccounts(i, 3) = Format(!COFDTXTYPE, "0")
      End With
   End If
   iTotal = i
   Set RdoGlm = Nothing
   
   ' Populate the finacial statement account layout JET table
   sSql = "DELETE FROM ActrpTable"
   JetDb.Execute sSql
   
   bChart = 0
   Set DbAct = JetDb.OpenRecordset("ActrpTable", dbOpenDynaset)
   For i = 1 To iTotal
      iCurType = i
      FillLevel1 Format(vAccounts(i, 0))
   Next
   DbAct.Close
   
   sSql = "DELETE FROM Current"
   JetDb.Execute sSql
   
   Set DbBal1 = JetDb.OpenRecordset("Current", dbOpenDynaset)
   
   sSql = "SELECT JIACCOUNT,SUM(GjitTable.JIDEB) - SUM(GjitTable.JICRD) AS Balance " _
          & "FROM GjhdTable INNER JOIN GjitTable ON GJNAME = JINAME WHERE " _
          & "GJPOST <= '" & txtEnd & "' AND " _
          & "GJPOST >= '" & txtBeg & "' AND " _
          & "GjhdTable.GJPOSTED = 1 GROUP BY JIACCOUNT"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct1)
   
   If bSqlRows Then
      With RdoAct1
         While Not .EOF
            DbBal1.AddNew
            DbBal1!ACCTREF = "" & Trim(!JIACCOUNT)
            DbBal1!ACCTBAL = !Balance
            DbBal1.Update
            .MoveNext
         Wend
         .Cancel
      End With
   End If
   Set RdoAct1 = Nothing
   DbBal1.Close
   
   
   sSql = "DELETE FROM Previous"
   JetDb.Execute sSql
   
   Set DbBal2 = JetDb.OpenRecordset("Previous", dbOpenDynaset)
   
   sSql = "SELECT JIACCOUNT,SUM(JIDEB) - SUM(JICRD) AS Balance " _
          & "FROM GjhdTable INNER JOIN GjitTable ON GJNAME = JINAME " _
          & "WHERE GJPOST < '" & txtBeg & "' AND GJPOSTED = 1 GROUP BY JIACCOUNT"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct2)
   
   If bSqlRows Then
      With RdoAct2
         While Not .EOF
            DbBal2.AddNew
            DbBal2!ACCTREF = "" & Trim(!JIACCOUNT)
            DbBal2!ACCTBAL = !Balance
            DbBal2.Update
            .MoveNext
         Wend
         .Cancel
      End With
   End If
   Set RdoAct2 = Nothing
   DbBal2.Close
   
   
   JetDb.Execute "DELETE FROM Divisions"
   
   ' Fill Division Criteria
   Set DbBal4 = JetDb.OpenRecordset("Divisions", dbOpenDynaset)
   sSql = "SELECT GLACCTREF FROM GlacTable " & sSqlAdder
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct4)
   If bSqlRows Then
      With RdoAct4
         While Not .EOF
            DbBal4.AddNew
            DbBal4!ACCTREF = "" & Trim(!GLACCTREF)
            DbBal4.Update
            .MoveNext
         Wend
      End With
   End If
   Set RdoAct4 = Nothing
   DbBal4.Close
   
   
   JetDb.Close
   PrintReport
   Exit Sub
   
DiaErr1:
   sProcName = "buildaccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Function bDivisionAccounts(iStart As Integer, iEnd As Integer) As Boolean
   Dim RdoDiv As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "SELECT COGLDIVISIONS, COGLDIVSTARTPOS, COGLDIVENDPOS FROM ComnTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDiv)
   If bSqlRows Then
      With RdoDiv
         If Val("" & !COGLDIVISIONS) <> 0 Then
            If Val(!COGLDIVSTARTPOS) <> 0 And Val(!COGLDIVENDPOS) <> 0 Then
               iStart = Val(!COGLDIVSTARTPOS)
               iEnd = Val(!COGLDIVENDPOS)
               bDivisionAccounts = True
            End If
         End If
      End With
   End If
   Set RdoDiv = Nothing
   Exit Function
DiaErr1:
   sProcName = "bDivisionAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Function bValidElement(cmbCombo As ComboBox) As Boolean
   Dim i As Integer
   On Error GoTo DiaErr1
   If cmbCombo.ListCount > 0 Then
      For i = 0 To cmbCombo.ListCount - 1
         If Val(cmbCombo.List(i)) = Val(cmbCombo.Text) Then
            bValidElement = True
            cmbCombo.ListIndex = i
         End If
      Next
   End If
   Exit Function
   
DiaErr1:
   sProcName = "bValidElement"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(txtBeg.Text) & Trim(txtEnd.Text)
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim dToday As Integer
   
   dToday = CInt(Mid(Format(Now, "mm/dd/yy"), 4, 2))
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)

   If Len(Trim(sOptions)) > 0 Then
     
     If dToday < 21 Then
      txtBeg = Mid(sOptions, 1, 8)
      txtEnd = Mid(sOptions, 9, 8)
     Else
      txtBeg = Format(Now, "mm/01/yy")
      txtEnd = GetMonthEnd(txtBeg)
     End If
   
   End If
   

   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub
