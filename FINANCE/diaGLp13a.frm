VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLp13a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Income/Expense Comparison (Report)"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkWholeDollar 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox chkShowGraph 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.ComboBox cboYear 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Tag             =   "3"
      Top             =   660
      Width           =   900
   End
   Begin VB.ComboBox cmbDiv 
      Height          =   315
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   7
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise A Division (2 char)"
      Top             =   3000
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.CheckBox optIna 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optDiv 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optCon 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtLvl 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Tag             =   "1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Tag             =   "4"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Tag             =   "4"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   4920
      TabIndex        =   12
      Top             =   240
      Width           =   1335
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   14
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
      PictureUp       =   "diaGLp13a.frx":0000
      PictureDn       =   "diaGLp13a.frx":0146
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   15
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
      PictureUp       =   "diaGLp13a.frx":028C
      PictureDn       =   "diaGLp13a.frx":03D2
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3600
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   1785
      FormDesignWidth =   6300
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Default is Thousands)"
      Height          =   285
      Index           =   2
      Left            =   3600
      TabIndex        =   20
      Top             =   1440
      Width           =   1860
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Display in Whole Dollar Amts?"
      Height          =   285
      Index           =   0
      Left            =   300
      TabIndex        =   18
      Top             =   1440
      Width           =   2340
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Graph?"
      Height          =   285
      Index           =   7
      Left            =   300
      TabIndex        =   17
      Top             =   1080
      Width           =   1380
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Year"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   16
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaGLp13a"
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
'
' diaGLp13a - Income/Expense Comparison
'
' Notes: Requested by JEVINT
'
' Created:  06/28/04 (nth)
' Revisions:
'
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
'Dim vAccounts(10, 4)    As Variant
Dim iStart As Integer
Dim iEnd As Integer

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
      'CreateActTable
      FillComboWithYears
      GetOptions
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   'ReopenJet
   sCurrForm = Caption
   'txtDte = "01/01/" & Format(ES_SYSDATE, "yy")
   'GetOptions
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
   On Error Resume Next
   'JetDb.Execute "DROP TABLE ActrpTable"
   'JetDb.Execute "DROP TABLE CurrentIncome"
   'JetDb.Execute "DROP TABLE Previous"
   'JetDb.Execute "DROP TABLE YTD"
   'JetDb.Execute "DROP TABLE Divisions"
   'JetDb.Execute "DROP TABLE tbl_Periods"
   Set diaGLp13a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

''Private Sub SpoolPeriods()
''    Dim i As Integer
''    Dim sDate As String
''    Dim sPeriod As String
''    Dim sSqlAdder   As String
''    Dim RdoGlm      As ADODB.RecordSet
''    Dim DbBal1      As Recordset
''    Dim dbper       As Recordset
''
''    MouseCursor 13
''    On Error Resume Next
''
''    ' Clear temp tables
''    sSql = "DELETE FROM ActrpTable"
''    JetDb.Execute sSql
''    sSql = "DELETE FROM CurrentIncome"
''    JetDb.Execute sSql
''    sSql = "DELETE FROM YTD"
''    JetDb.Execute sSql
''    sSql = "DELETE FROM Previous"
''    JetDb.Execute sSql
''    sSql = "DELETE FROM Divisions"
''    JetDb.Execute sSql
''    sSql = "DELETE FROM tbl_Periods"
''    JetDb.Execute sSql
''
''    On Error GoTo DiaErr1
''
''    If Trim(cmbDiv) <> "" Then
''        sSqlAdder = " WHERE  (RIGHT(LEFT(GLACCTNO + '            ', " _
''            & iEnd & "), " & iEnd & " - (" & iStart & " - 1)) = '" & cmbDiv & "')"
''    End If
''
''    bChart = 1
''
''    If Trim(txtLvl) = "" Then txtLvl = 9
''    iLevel = 9 'Val(txtLvl)
''
''    ' Build income statement account structure
''
''    sSql = "DELETE FROM FSS"
''    JetDb.Execute sSql
''    sSql = "SELECT COINCMACCT,COCOGSACCT,COEXPNACCT,COOINCACCT,COOEXPACCT," _
''        & "COFDTXACCT FROM GlmsTable"
''    bSqlRows = clsAdoCon.GetDataSet(sSql,RdoGlm)
''    If bSqlRows Then
''        With RdoGlm
''            Set DbBal1 = JetDb.OpenRecordset("FSS", dbOpenDynaset)
''            For i = 0 To 5
''                DbBal1.AddNew
''                    DbBal1!MASTERREF = i + 4
''                    DbBal1!MASTERDESC = Trim(.Fields(i))
''                DbBal1.Update
''            Next
''        End With
''    End If
''
''    sSql = "SELECT * FROM GlmsTable WHERE COACCTREC=1"
''    bSqlRows = clsAdoCon.GetDataSet(sSql,RdoGlm)
''    If bSqlRows Then
''        With RdoGlm
''            i = 4
''            vAccounts(i, 0) = "" & Trim(!COINCMREF)
''            vAccounts(i, 1) = "" & Trim(!COINCMACCT)
''            vAccounts(i, 2) = "" & Trim(!COINCMDESC)
''            vAccounts(i, 3) = Format(!COINCMTYPE, "0")
''            i = 6
''            vAccounts(i, 0) = "" & Trim(!COEXPNREF)
''            vAccounts(i, 1) = "" & Trim(!COEXPNACCT)
''            vAccounts(i, 2) = "" & Trim(!COEXPNDESC)
''            vAccounts(i, 3) = Format(!COEXPNTYPE, "0")
''        End With
''    End If
''    iTotal = i
''    Set RdoGlm = Nothing
''
''    ' Load Income/Expense Accounts
''    Set DbAct = JetDb.OpenRecordset("ActrpTable", dbOpenDynaset)
''    bChart = 0
''    iInActive = Val(optIna)
''    For i = 4 To 6 Step 2
''        iCurType = i
''        FillLevel1 Format(vAccounts(i, 0))
''    Next
''    sDate = txtDte
''    For i = 1 To Val(txtPer)
''        sPeriod = Format(sDate, "mmm yyyy")
''        sSql = "INSERT INTO tbl_Periods(PeriodID,PeriodDesc) VALUES(" _
''        & i & ",'" & sPeriod & "')"
''        JetDb.Execute sSql
''        BuildAccounts i, Format(sDate, "mm/01/yy"), GetMonthEnd(sDate)
''        sDate = DateAdd("m", 1, sDate)
''    Next
''    DbAct.Close
''    Exit Sub
''DiaErr1:
''    sProcName = "spoolperiods"
''    CurrError.Number = Err.Number
''    CurrError.description = Err.description
''    DoModuleErrors Me
''End Sub
''
''Private Sub BuildAccounts(iPeriod As Integer, sBegin As String, sEnd As String)
''    Dim i           As Integer
''
''    Dim RdoAct1     As ADODB.RecordSet
''    Dim RdoAct2     As ADODB.RecordSet
''    Dim RdoAct3     As ADODB.RecordSet
''    Dim RdoAct4     As ADODB.RecordSet
''    Dim DbBal1      As Recordset
''    Dim DbBal2      As Recordset
''    Dim DbBal3      As Recordset
''    Dim DbBal4      As Recordset 'Divisions
''    Dim sAccount    As String
''    Dim sRatioAcct  As String
''    Dim sSqlAdder   As String
''
''    On Error GoTo DiaErr1
''
''   ' Current
''    Set DbBal1 = JetDb.OpenRecordset("CurrentIncome", dbOpenDynaset)
''    sSql = "SELECT JIACCOUNT,SUM(GjitTable.JICRD) - SUM(GjitTable.JIDEB) AS Balance " _
''        & "FROM GjhdTable INNER JOIN GjitTable ON GJNAME = JINAME " _
''        & "WHERE GJPOST <= '" & sEnd & "' AND GJPOST >= '" & sBegin & "' AND " _
''        & "GjhdTable.GJPOSTED = 1 GROUP BY JIACCOUNT"
''    bSqlRows = clsAdoCon.GetDataSet(sSql,RdoAct1)
''        If bSqlRows Then
''            With RdoAct1
''                While Not .EOF
''                    DbBal1.AddNew
''                    DbBal1!ACCTREF = "" & Trim(!JIACCOUNT)
''                    DbBal1!ACCTBAL = !Balance
''                    DbBal1!AcctPer = iPeriod
''                    DbBal1.Update
''                    .MoveNext
''                Wend
''            End With
''        End If
''    Set RdoAct1 = Nothing
''    DbBal1.Close
''
''    ' YTD
''    Set DbBal2 = JetDb.OpenRecordset("YTD", dbOpenDynaset)
''    sSql = "SELECT JIACCOUNT,SUM(GjitTable.JICRD) - SUM(GjitTable.JIDEB) AS Balance " _
''        & "FROM GjhdTable INNER JOIN GjitTable ON GJNAME = JINAME " _
''        & "WHERE (GJPOST <= '" & sEnd & "' AND GJPOST >= '" & txtDte & "') AND " _
''        & "GjhdTable.GJPOSTED = 1 GROUP BY JIACCOUNT"
''    bSqlRows = clsAdoCon.GetDataSet(sSql,RdoAct2)
''    If bSqlRows Then
''        With RdoAct2
''            While Not .EOF
''                DbBal2.AddNew
''                DbBal2!ACCTREF = "" & Trim(!JIACCOUNT)
''                DbBal2!ACCTBAL = !Balance
''                DbBal2!AcctPer = iPeriod
''                DbBal2.Update
''
''                .MoveNext
''            Wend
''        End With
''    End If
''    Set RdoAct2 = Nothing
''    DbBal2.Close
''
''    ' Previous
''    Set DbBal3 = JetDb.OpenRecordset("Previous", dbOpenDynaset)
''    sSql = "SELECT JIACCOUNT,SUM(GjitTable.JICRD) - SUM(GjitTable.JIDEB) AS Balance " _
''        & "FROM GjhdTable INNER JOIN GjitTable ON GJNAME = JINAME " _
''        & "WHERE (GJPOST <= '" & DateAdd("yyyy", -1, sEnd) & "' AND GJPOST >= '" _
''        & DateAdd("yyyy", -1, txtDte) & "') AND " _
''        & "GjhdTable.GJPOSTED = 1 GROUP BY JIACCOUNT"
''    bSqlRows = clsAdoCon.GetDataSet(sSql,RdoAct3)
''    If bSqlRows Then
''        With RdoAct3
''            While Not .EOF
''                DbBal3.AddNew
''                DbBal3!ACCTREF = "" & Trim(!JIACCOUNT)
''                DbBal3!ACCTBAL = !Balance
''                DbBal3!AcctPer = iPeriod
''                DbBal3.Update
''                .MoveNext
''            Wend
''        End With
''    End If
''    Set RdoAct3 = Nothing
''    DbBal3.Close
''
''
''    ' Division
''    Set DbBal4 = JetDb.OpenRecordset("Divisions", dbOpenDynaset)
''    sSql = "SELECT GLACCTREF FROM GlacTable " & sSqlAdder
''    bSqlRows = clsAdoCon.GetDataSet(sSql,RdoAct4)
''    If bSqlRows Then
''        With RdoAct4
''            While Not .EOF
''                DbBal4.AddNew
''                DbBal4!ACCTREF = "" & Trim(!GLACCTREF)
''                DbBal4.Update
''                .MoveNext
''            Wend
''        End With
''    End If
''    Set RdoAct4 = Nothing
''    DbBal4.Close
''    'JetDb.Close
''
''    Exit Sub
''
''DiaErr1:
''    sProcName = "buildaccou"
''    CurrError.Number = Err.Number
''    CurrError.description = Err.description
''    DoModuleErrors Me
''End Sub
''
''

''Private Sub CreateActTable()
''    Dim NewTb1  As TableDef
''    Dim NewTb2  As TableDef
''    Dim NewTb3  As TableDef
''    Dim NewTb4  As TableDef
''    Dim NewTb5  As TableDef
''    Dim NewTb6  As TableDef
''    Dim NewTb7  As TableDef
''    Dim NewTb8  As TableDef
''    Dim NewTb9  As TableDef
''    Dim NewTb10 As TableDef 'If necessary
''
''    Dim NewFld  As Field
''    Dim NewIdx1 As Index
''    Dim NewIdx2 As Index
''    Dim NewIdx3 As Index
''    Dim NewIdx4 As Index
''    Dim NewIdx5 As Index
''    Dim Newidx6 As Index
''
''
''        On Error Resume Next
''    JetDb.Execute "DROP TABLE ActrpTable"
''    JetDb.Execute "DROP TABLE CurrentIncome"
''    JetDb.Execute "DROP TABLE Previous"
''    JetDb.Execute "DROP TABLE YTD"
''    JetDb.Execute "DROP TABLE Divisions"
''    JetDb.Execute "DROP TABLE tbl_Periods"
''
''
''    ' Create FSS table (Financial Statement Structure)
''
''
''    Set NewTb1 = JetDb.CreateTableDef("FSS")
''        With NewTb1
''            .Fields.Append .CreateField("MASTERREF", dbInteger)
''            .Fields.Append .CreateField("MASTERDESC", dbText, 12)
''        End With
''    JetDb.TableDefs.Append NewTb1
''    Set NewTb1 = Nothing
''
''    Set NewTb1 = JetDb!FSS
''    With NewTb1
''        Set NewIdx1 = .CreateIndex
''        With NewIdx1
''            .Name = "FSS_INDEX1"
''            .Fields.Append .CreateField("MASTERREF")
''        End With
''        .Indexes.Append NewIdx1
''    End With
''    Set NewTb1 = Nothing
''    Set NewIdx1 = Nothing
''
''    ' Create fincial statement layout
''
''    ' Fields. Note that we allow empties
''    Set NewTb1 = JetDb.CreateTableDef("ActrpTable")
''        With NewTb1
''            'Type
''            .Fields.Append .CreateField("Act00", dbInteger)
''            'Level
''            .Fields.Append .CreateField("Act01", dbInteger)
''            'AcctRef
''            .Fields.Append .CreateField("Act02", dbText, 12)
''            .Fields(2).AllowZeroLength = True
''            'Account Number + Spaces to indent
''            .Fields.Append .CreateField("Act03", dbText, 32)
''            .Fields(3).AllowZeroLength = True
''            'Account Desc
''            .Fields.Append .CreateField("Act04", dbText, 60)
''            .Fields(4).AllowZeroLength = True
''            'Active
''            .Fields.Append .CreateField("Act05", dbInteger)
''            'GLFSLEVEL
''            .Fields.Append .CreateField("Act06", dbInteger)
''        End With
''
''    ' Add the table and indexes to Jet.
''    JetDb.TableDefs.Append NewTb1
''    Set NewTb2 = JetDb!ActrpTable
''        With NewTb2
''            Set NewIdx1 = .CreateIndex
''                With NewIdx1
''                    .Name = "I1"
''                    .Fields.Append .CreateField("Act02")
''                End With
''                .Indexes.Append NewIdx1
''        End With
''
''    ' Create the CurrentIncome account balance table
''
''    ' Fields. Note that we allow empties
''    Set NewTb3 = JetDb.CreateTableDef("CurrentIncome")
''    With NewTb3
''        ' AcctRef
''        .Fields.Append .CreateField("AcctRef", dbText, 12)
''        .Fields(0).AllowZeroLength = True
''        ' CurrentIncome Period
''        .Fields.Append .CreateField("AcctBal", dbCurrency)
''        .Fields.Append .CreateField("AcctPer", dbInteger)
''    End With
''    ' Add the table and indexes to Jet.
''    JetDb.TableDefs.Append NewTb3
''    Set NewTb4 = JetDb!CurrentIncome
''    With NewTb4
''        Set NewIdx3 = .CreateIndex
''        With NewIdx3
''            .Name = "Index1"
''            .Fields.Append .CreateField("AcctRef")
''        End With
''        .Indexes.Append NewIdx3
''    End With
''
''    Set NewTb4 = Nothing
''    Set NewIdx3 = Nothing
''
''    JetDb.TableDefs.Append NewTb3
''    Set NewTb4 = JetDb!CurrentIncome
''    With NewTb4
''        Set NewIdx3 = .CreateIndex
''        With NewIdx3
''            .Name = "Index2"
''            .Fields.Append .CreateField("AcctPer")
''        End With
''        .Indexes.Append NewIdx3
''    End With
''
''
''    ' Create the YTD account balance table
''
''    ' Fields. Note that we allow empties
''    Set NewTb5 = JetDb.CreateTableDef("YTD")
''    With NewTb5
''        ' AcctRef
''        .Fields.Append .CreateField("AcctRef", dbText, 12)
''        .Fields(0).AllowZeroLength = True
''        ' CurrentIncome Period
''        .Fields.Append .CreateField("AcctBal", dbCurrency)
''        .Fields.Append .CreateField("AcctPer", dbInteger)
''    End With
''    ' Add the table and indexes to Jet.
''    JetDb.TableDefs.Append NewTb5
''    Set NewTb6 = JetDb!YTD
''    With NewTb6
''        Set NewIdx4 = .CreateIndex
''        With NewIdx4
''            .Name = "AcctNo"
''            .Fields.Append .CreateField("AcctRef")
''        End With
''        .Indexes.Append NewIdx4
''    End With
''
''
''
''
''    ' Create the Previous account balance table
''
''    ' Fields. Note that we allow empties
''    Set NewTb7 = JetDb.CreateTableDef("Previous")
''    With NewTb7
''        ' AcctRef
''        .Fields.Append .CreateField("AcctRef", dbText, 12)
''        .Fields(0).AllowZeroLength = True
''        ' CurrentIncome Period
''        .Fields.Append .CreateField("AcctBal", dbCurrency)
''        .Fields.Append .CreateField("AcctPer", dbInteger)
''    End With
''    ' Add the table and indexes to Jet.
''    JetDb.TableDefs.Append NewTb7
''    Set NewTb8 = JetDb!Previous
''    With NewTb8
''        Set NewIdx5 = .CreateIndex
''        With NewIdx5
''            .Name = "AcctNo"
''            .Fields.Append .CreateField("AcctRef")
''        End With
''        .Indexes.Append NewIdx5
''    End With
''
''
''    ' Create the Divisions
''
''    ' Fields. Note that we allow empties
''
''    Set NewTb9 = JetDb.CreateTableDef("Divisions")
''    With NewTb9
''        ' AcctRef
''        .Fields.Append .CreateField("AcctRef", dbText, 12)
''        .Fields(0).AllowZeroLength = True
'''*        ' CurrentIncome Period
'''*       .Fields.Append .CreateField("AcctBal", dbCurrency)
''    End With
''    ' Add the table and indexes to Jet.
''
''    JetDb.TableDefs.Append NewTb9
''    Set NewTb10 = JetDb!Divisions
''    With NewTb10
''        Set Newidx6 = .CreateIndex
''        Newidx6.Name = "AcctNo"
''        Newidx6.Fields.Append Newidx6.CreateField("AcctRef")
''        .Indexes.Append Newidx6
''    End With
''
''    ' Create Fiscal Periods Table
''
''
''    Set NewTb1 = JetDb.CreateTableDef("tbl_Periods")
''        With NewTb1
''            .Fields.Append .CreateField("PeriodID", dbInteger)
''            .Fields.Append .CreateField("PeriodDesc", dbText, 12)
''        End With
''    JetDb.TableDefs.Append NewTb1
''    Set NewTb1 = Nothing
''
''    Set NewTb1 = JetDb!tbl_Periods
''    With NewTb1
''        Set NewIdx1 = .CreateIndex
''        With NewIdx1
''            .Name = "Index1"
''            .Fields.Append .CreateField("PeriodID")
''        End With
''        .Indexes.Append NewIdx1
''    End With
''    Set NewTb1 = Nothing
''    Set NewIdx1 = Nothing
''
''End Sub
''

Private Sub PrintReport()
   Dim sWindows As String
   Dim sStartDate As String
   Dim sEnddate As String
   Dim sYear As String
   
   Dim sCustomReport As String
   On Error GoTo DiaErr1
   
   'SetMdiReportsize MdiSect
   sWindows = GetWindowsDir()
   sYear = cboYear
   sStartDate = "1/1/" & sYear
   sEnddate = "12/31/" & sYear
   'ReopenJet
   
   'MdiSect.crw.DataFiles(0) = sWindows & "\temp\esifina.mdb"
   
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   'MdiSect.crw.Formulas(1) = "Title1='Level " & txtLvl _
   & " Income Statement For Year Beginning " & txtDte & "'"
   'MdiSect.crw.Formulas(2) = "Title2 = 'Period Beginning:  " _
   '    & txtBeg & " And Ending:  " & txtEnd & "'"
   'MdiSect.crw.Formulas(3) = "nDetailLevel = " & Val(txtLvl)
   MdiSect.crw.Formulas(1) = "Year= '" & cboYear & "'"
   MdiSect.crw.Formulas(2) = "StartDate= '" & sStartDate & "'"
   MdiSect.crw.Formulas(3) = "EndDate= '" & sEnddate & "'"
   MdiSect.crw.Formulas(7) = "ShowGraph=" & chkShowGraph.Value
   MdiSect.crw.Formulas(8) = "DisplayInDollars=" & chkWholeDollar.Value
   
   sCustomReport = GetCustomReport("fingl13.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   
   'sSql = "trim(cstr({Divisions.AcctRef})) <> ''"
   'MdiSect.crw.SelectionFormula = sSql
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
   'SpoolPeriods
   PrintReport
End Sub

Private Sub optPrn_Click()
   'SpoolPeriods
   PrintReport
End Sub

'Private Function bDivisionAccounts(iStart As Integer, iEnd As Integer) As Boolean
'    Dim RdoDiv      As ADODB.RecordSet
'    On Error GoTo DiaErr1
'
'    sSql = "SELECT COGLDIVISIONS, COGLDIVSTARTPOS, COGLDIVENDPOS FROM ComnTable"
'    bSqlRows = clsAdoCon.GetDataSet(sSql,RdoDiv)
'    If bSqlRows Then
'        With RdoDiv
'            If Val("" & !COGLDIVISIONS) <> 0 Then
'                If Val(!COGLDIVSTARTPOS) <> 0 And Val(!COGLDIVENDPOS) <> 0 Then
'                    iStart = Val(!COGLDIVSTARTPOS)
'                    iEnd = Val(!COGLDIVENDPOS)
'                    bDivisionAccounts = True
'                End If
'            End If
'        End With
'    End If
'    Exit Function
'
'DiaErr1:
'    sProcName = "bDivisionAccounts"
'    CurrError.Number = Err.Number
'    CurrError.description = Err.description
'    DoModuleErrors Me
'End Function
'

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

'Private Sub txtDte_LostFocus()
'    txtDte = CheckDate(txtDte)
'End Sub
'
'Private Sub txtPer_LostFocus()
'    If txtPer < 1 Then
'        txtPer = 1
'    ElseIf txtPer > 13 Then
'        txtPer = 13
'    End If
'End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiFina", Me.Name & "Year", cboYear
   
   Dim sOptions As String
   sOptions = chkShowGraph & chkWholeDollar & "00"
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   On Error Resume Next
   Dim defaultYear As String
   defaultYear = Format(Date, "yyyy")
   cboYear = GetSetting("Esi2000", "EsiFina", Me.Name & "Year", defaultYear)
   
   Dim sOptions As String
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, "0000")
   chkShowGraph.Value = Mid(sOptions, 1, 1)
   chkWholeDollar.Value = Mid(sOptions, 2, 1)
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
   
End Sub

Private Sub FillComboWithYears()
   Dim rdo As ADODB.Recordset
   cboYear.Clear
   
   sSql = "Select distinct cast(datepart( year, MJSTART ) as char(4) ) from JrhdTable " _
          & "order by cast(datepart( year, MJSTART ) as char(4) )"
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
         While Not .EOF
            cboYear.AddItem Trim(CStr(rdo(0)))
            .MoveNext
         Wend
      End With
   End If
   
   If cboYear.ListCount = 0 Then
      cboYear.AddItem Format(Date, "yyyy")
   End If
   
   If cboYear.ListCount > 0 Then
      cboYear.ListIndex = cboYear.ListCount - 1
   End If
   Set rdo = Nothing
End Sub
