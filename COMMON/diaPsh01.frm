VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaPsh01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Orders"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6240
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "Revision-Select From List"
      Top             =   1800
      Width           =   1095
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   225
      Left            =   280
      TabIndex        =   46
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaPsh01.frx":0000
      PictureDn       =   "diaPsh01.frx":0146
   End
   Begin VB.CheckBox optLst 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      ToolTipText     =   "Pick List For This Part (Printed MO's Only) Status PL"
      Top             =   3480
      Width           =   726
   End
   Begin VB.CheckBox optDoc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      ToolTipText     =   "Document List (Printed MO's Only)"
      Top             =   3240
      Width           =   735
   End
   Begin VB.CheckBox optBud 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox optFrom 
      Height          =   255
      Left            =   3960
      TabIndex        =   40
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6240
      TabIndex        =   37
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   600
         Picture         =   "diaPsh01.frx":0298
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaPsh01.frx":0422
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6240
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.CheckBox optInc 
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   12
      Left            =   2760
      TabIndex        =   9
      Top             =   3720
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   11
      Left            =   2760
      TabIndex        =   15
      Top             =   5040
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   14
      Top             =   4800
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   9
      Left            =   3960
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   7
      Left            =   2760
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   10
      Top             =   3960
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   5
      Top             =   2520
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   4
      Top             =   2280
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Caption         =   "____"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      Top             =   2040
      Width           =   735
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6240
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Select Run Number"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Part Numbers With Manufacturing Orders"
      Top             =   1080
      Width           =   3545
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   17
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
      PictureUp       =   "diaPsh01.frx":05A0
      PictureDn       =   "diaPsh01.frx":06E6
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4800
      FormDesignWidth =   7515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PL Rev"
      Height          =   255
      Index           =   18
      Left            =   5160
      TabIndex        =   47
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   45
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label lblQty 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   44
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick List For This Part"
      Height          =   255
      Index           =   17
      Left            =   480
      TabIndex        =   43
      ToolTipText     =   "Pick List For This Part (Printed MO's Only) Status PL"
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label lblSta 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6840
      TabIndex        =   42
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Budgets"
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   41
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label lblTyp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   39
      Top             =   1440
      Width           =   300
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type/Status"
      Height          =   255
      Index           =   15
      Left            =   5160
      TabIndex        =   38
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Information"
      Enabled         =   0   'False
      Height          =   255
      Index           =   14
      Left            =   480
      TabIndex        =   33
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspection Document"
      Enabled         =   0   'False
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   32
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bar Code"
      Enabled         =   0   'False
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   31
      Top             =   4800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cover Sheet (Printed MO's Only):"
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   30
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Allocations"
      Enabled         =   0   'False
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   29
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Comments"
      Enabled         =   0   'False
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   28
      Top             =   4320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Document List For This Part"
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   27
      ToolTipText     =   "Document List (Printed MO's Only)"
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SO Allocations"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   26
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Allocations"
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   3960
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Outside Service Part Numbers"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   24
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Comments"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   23
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   22
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   19
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   18
      Top             =   1440
      Width           =   3255
   End
End
Attribute VB_Name = "diaPsh01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/13/02 Added PKRECORD for new index
Option Explicit
Dim RdoQry As rdoQuery
Dim DbDoc As Recordset 'Jet
Dim DbPls As Recordset 'Jet

Dim bGoodPart As Byte
Dim bGoodMo As Byte
Dim bOnLoad As Byte
Dim bTablesCreated As Byte

Dim sBomRev As String
Dim sRunPkstart As String
Dim sPartNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetRevisions()
   cmbRev.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREV FROM BmhdTable WHERE BMHREF='" _
          & Compress(cmbPrt) & "' ORDER BY BMHREV"
   LoadComboBox cmbRev, -1
   Exit Sub
   
DiaErr1:
   sProcName = "getthisrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh01", sOptions)
   If Len(sOptions) > 0 Then
      For iList = 1 To 5
         optInc(iList) = Val(Mid(sOptions, iList, 1))
      Next
      For iList = 7 To 11
         optInc(iList) = Val(Mid(sOptions, iList, 1))
      Next
      optBud = Val(Mid(sOptions, iList, 1))
   End If
   optInc(5).Value = GetSetting("Esi2000", "EsiProd", "sh01all", optInc(5).Value)
   lblPrinter = GetSetting("Esi2000", "EsiProd", "sh01Printer", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
End Sub


Private Sub SaveOptions()
   Dim iList As Integer
   Dim sOptions As String
   'Save by Menu Option
   For iList = 1 To 5
      sOptions = sOptions & Trim(Val(optInc(iList).Value))
   Next
   For iList = 7 To 11
      sOptions = sOptions & Trim(Val(optInc(iList).Value))
   Next
   sOptions = sOptions & Trim(Val(optInc(iList).Value))
   sOptions = sOptions & Trim(Val(optBud.Value))
   SaveSetting "Esi2000", "EsiProd", "sh01", Trim(sOptions)
   SaveSetting "Esi2000", "EsiProd", "sh01all", Trim(optInc(5).Value)
   SaveSetting "Esi2000", "EsiProd", "sh01Printer", lblPrinter
   
End Sub




Private Sub cmbPrt_Click()
   bGoodPart = GetRuns()
   If bGoodPart Then GetRevisions
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   bGoodPart = GetRuns()
   If bGoodPart Then GetRevisions
   
End Sub

Private Sub PrintReport()
   MouseCursor 13
   On Error GoTo Psh01
   SetMdiReportsize MdiSect
   sProcName = "printreport"
   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   If optInc(1) Then
      sCustomReport = GetCustomReport("prdsh01")
      MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   Else
      sCustomReport = GetCustomReport("prdsh01a")
      MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   End If
   If optInc(2).Value = vbUnchecked Then
      MdiSect.Crw.SectionFormat(0) = "DETAIL.0.0;F;;;"
      MdiSect.Crw.SectionFormat(1) = "DETAIL.0.1;F;;;"
   Else
      MdiSect.Crw.SectionFormat(0) = "DETAIL.0.0;T;;;"
      MdiSect.Crw.SectionFormat(1) = "DETAIL.0.1;T;;;"
   End If
   If optInc(3).Value = vbUnchecked Then
      MdiSect.Crw.SectionFormat(2) = "GROUPFTR.0.0;F;;;"
      MdiSect.Crw.SectionFormat(3) = "GROUPFTR.0.1;F;;;"
   Else
      MdiSect.Crw.SectionFormat(2) = "GROUPFTR.0.0;T;;;"
      MdiSect.Crw.SectionFormat(3) = "GROUPFTR.0.1;T;;;"
   End If
   If optInc(1).Value = vbChecked Then
      If optInc(2).Value = vbChecked Then
         MdiSect.Crw.SectionFormat(1) = "DETAIL.0.1;F;;;"
      End If
   End If
   If optBud.Value = vbChecked Then
      MdiSect.Crw.SectionFormat(4) = "REPORTFTR.0.0;T;;;"
   Else
      MdiSect.Crw.SectionFormat(4) = "REPORTFTR.0.0;F;;;"
   End If
   sSql = "{RunsTable.RUNREF}='" & sPartNumber & "' " _
          & "AND {RunsTable.RUNNO}=" & Trim(cmbRun) & " "
   MdiSect.Crw.SelectionFormula = sSql
   SetCrystalAction Me
   MouseCursor 0
   DoEvents
   If optPrn Then
      If optInc(5) Then PrintAllocations
   End If
   Exit Sub
   
Psh01:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Psh02
Psh02:
   DoModuleErrors Me
   
End Sub

Private Sub cmbRev_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbRev = CheckLen(cmbRev, 4)
   For iList = 0 To cmbRev.ListCount - 1
      If Trim(cmbRev) = Trim(cmbRev.List(iList)) Then b = 1
   Next
   If b = 0 And cmbRev.ListCount > 0 Then
      Beep
      cmbRev = cmbRev.List(0)
   End If
   sBomRev = Trim(cmbRev)
   
End Sub


Private Sub cmbRun_Click()
   GetThisRun
   
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   GetThisRun
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs4120"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub




Private Sub Form_Activate()
   If bOnLoad Then
      ReopenJet
      CreateJetTables
      FillAllRuns cmbPrt
      If optFrom.Value = vbChecked Then
         cmbPrt = diaSrvmo.cmbPrt
         cmbRun = diaSrvmo.cmbRun
      End If
      bGoodPart = GetRuns()
      If bGoodPart Then GetRevisions
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   GetOptions
   bTablesCreated = 0
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PABOMREV,PARUN,RUNREF,RUNSTATUS," _
          & "RUNNO FROM PartTable,RunsTable WHERE PARTREF= ? " _
          & "AND PARTREF=RUNREF "
   Set RdoQry = RdoCon.CreateQuery("", sSql)
   bOnLoad = 1
   
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   ' JetDb.Execute "DROP TABLE CvrTable"
   ' JetDb.Execute "DROP TABLE PlsTable"
   Set RdoQry = Nothing
   'RdoRes.Close
   If optFrom Then diaSrvmo.Show Else FormUnload
   Set diaPsh01 = Nothing
   
End Sub




Private Function GetRuns() As Byte
   Dim RdoRns As rdoResultset
   On Error GoTo DiaErr1
   MouseCursor 13
   cmbRun.Clear
   sPartNumber = Compress(cmbPrt)
   RdoQry(0) = sPartNumber
   bSqlRows = GetQuerySet(RdoRns, RdoQry)
   If bSqlRows Then
      With RdoRns
         If optFrom Then
            cmbRun = diaSrvmo.cmbRun
         Else
            cmbRun = Format(!Runno, "####0")
         End If
         lblDsc = "" & Trim(!PADESC)
         lblTyp = Format(!PALEVEL, "#")
         cmbRev = "" & Trim(!PABOMREV)
         Do Until .EOF
            AddComboStr cmbRun.hWnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         .Cancel
      End With
      GetRuns = True
      GetThisRun
   Else
      sPartNumber = ""
      GetRuns = False
   End If
   MouseCursor 0
   Set RdoRns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblQty_Click()
   'run qty
   
End Sub

Private Sub optBud_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   If Not bGoodPart Then
      MsgBox "Couldn't Find Part Number, Run.", vbExclamation, Caption
      On Error Resume Next
      cmbPrt.SetFocus
      Exit Sub
   Else
      ReopenJet
      If bTablesCreated = 0 Then CreateJetTables
      PrintReport
      MouseCursor 0
   End If
   
End Sub

Private Sub optDoc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFrom_Click()
   'dummy to check if from Revise mo
   
End Sub



Private Sub optLst_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   Dim b As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   
   ReopenJet
   If Not bGoodPart Then
      MsgBox "Couldn't Find Part Number, Run.", vbExclamation, Caption
      On Error Resume Next
      cmbPrt.SetFocus
      Exit Sub
   Else
      If bTablesCreated = 0 Then CreateJetTables
      On Error Resume Next
      JetDb.Execute "DELETE * FROM PlsTable"
      JetDb.Execute "DELETE * FROM CvrTable"
      'Doc and Pick List only for printed reports
      If optLst.Value = vbChecked Then
         If lblSta = "SC" Or lblSta = "RL" Then
            cmbRev.enabled = True
            sMsg = "Do You Want To Print The MO Pick " & vbCr _
                   & "List And Move The Run Status To PL?"
            bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
            If bResponse = vbYes Then
               'Build the pick list and change status
               b = 1
               MouseCursor 13
               BuildPartsList
            Else
               CancelTrans
            End If
         Else
            cmbRev.enabled = False
            MouseCursor 13
            b = 1
            BuildPickList
         End If
      End If
      If optDoc = vbChecked Then
         MouseCursor 13
         b = 1
         BuildDocumentList
      End If
      If b = 1 Then PrintCover
      PrintReport
   End If
   
End Sub



Private Sub GetThisRun()
   Dim RdoRun As rdoResultset
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,RUNPKSTART,RUNQTY FROM RunsTable WHERE " _
          & "RUNREF='" & Compress(cmbPrt) & "' AND " _
          & "RUNNO=" & cmbRun & " "
   bSqlRows = GetDataSet(RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         lblSta = "" & Trim(!RUNSTATUS)
         If lblSta = "SC" Or lblSta = "RL" Then cmbRev.enabled = True _
                     Else cmbRev.enabled = False
         If Not IsNull(!RUNPKSTART) Then
            sRunPkstart = Format(!RUNPKSTART, "mm/dd/yy")
         Else
            sRunPkstart = Format(ES_SYSDATE, "mm/dd/yy")
         End If
         lblQty = Format(!RUNQTY, ES_QuantityDataFormat)
         .Cancel
      End With
   End If
   Set RdoRun = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getthisrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub CreateCovrTable()
   'Drop and create a Jet table to
   'run beside SQL Server so Crystal can handle the report
   '(Crystal isn't smart enough to handle the joins)
   Dim NewTb As TableDef
   Dim NewFld As Field
   
   On Error Resume Next
   JetDb.Execute "SELECT DLSPart1 FROM CvrTable"
   If Err > 0 Then
      'Fields. Note that we allow empties
      Set NewTb = JetDb.CreateTableDef("CvrTable")
      With NewTb
         'Documents
         '1
         .Fields.Append .CreateField("DLSPart1", dbText, 30)
         .Fields(0).AllowZeroLength = True
         .Fields.Append .CreateField("DLSRev1", dbText, 6)
         .Fields(1).AllowZeroLength = True
         .Fields.Append .CreateField("DLSType1", dbInteger)
         .Fields.Append .CreateField("DLSDocRef1", dbText, 30)
         .Fields(3).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocRev1", dbText, 6)
         .Fields(4).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocSheet1", dbText, 6)
         .Fields(5).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocClass1", dbText, 16)
         .Fields(6).AllowZeroLength = True
         '2
         .Fields.Append .CreateField("DLSPart2", dbText, 30)
         .Fields(7).AllowZeroLength = True
         .Fields.Append .CreateField("DLSRev2", dbText, 6)
         .Fields(8).AllowZeroLength = True
         .Fields.Append .CreateField("DLSType2", dbInteger)
         .Fields.Append .CreateField("DLSDocRef2", dbText, 30)
         .Fields(10).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocRev2", dbText, 6)
         .Fields(11).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocSheet2", dbText, 6)
         .Fields(12).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocClass2", dbText, 16)
         .Fields(13).AllowZeroLength = True
         '3
         .Fields.Append .CreateField("DLSPart3", dbText, 30)
         .Fields(14).AllowZeroLength = True
         .Fields.Append .CreateField("DLSRev3", dbText, 6)
         .Fields(15).AllowZeroLength = True
         .Fields.Append .CreateField("DLSType3", dbInteger)
         .Fields.Append .CreateField("DLSDocRef3", dbText, 30)
         .Fields(17).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocRev3", dbText, 6)
         .Fields(18).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocSheet3", dbText, 6)
         .Fields(19).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocClass3", dbText, 16)
         .Fields(20).AllowZeroLength = True
         '4
         .Fields.Append .CreateField("DLSPart4", dbText, 30)
         .Fields(21).AllowZeroLength = True
         .Fields.Append .CreateField("DLSRev4", dbText, 6)
         .Fields(22).AllowZeroLength = True
         .Fields.Append .CreateField("DLSType4", dbInteger)
         .Fields.Append .CreateField("DLSDocRef4", dbText, 30)
         .Fields(24).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocRev4", dbText, 6)
         .Fields(25).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocSheet4", dbText, 6)
         .Fields(26).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocClass4", dbText, 16)
         .Fields(27).AllowZeroLength = True
         '5
         .Fields.Append .CreateField("DLSPart5", dbText, 30)
         .Fields(28).AllowZeroLength = True
         .Fields.Append .CreateField("DLSRev5", dbText, 6)
         .Fields(29).AllowZeroLength = True
         .Fields.Append .CreateField("DLSType5", dbInteger)
         .Fields.Append .CreateField("DLSDocRef5", dbText, 30)
         .Fields(31).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocRev5", dbText, 6)
         .Fields(32).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocSheet5", dbText, 6)
         .Fields(33).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocClass5", dbText, 16)
         .Fields(34).AllowZeroLength = True
         
         'added
         .Fields.Append .CreateField("DLSDocDesc1", dbText, 60)
         .Fields(35).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocLoc1", dbText, 4)
         .Fields(36).AllowZeroLength = True
         
         .Fields.Append .CreateField("DLSDocDesc2", dbText, 60)
         .Fields(37).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocLoc2", dbText, 4)
         .Fields(38).AllowZeroLength = True
         
         .Fields.Append .CreateField("DLSDocDesc3", dbText, 60)
         .Fields(39).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocLoc3", dbText, 4)
         .Fields(40).AllowZeroLength = True
         
         .Fields.Append .CreateField("DLSDocDesc4", dbText, 60)
         .Fields(41).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocLoc4", dbText, 4)
         .Fields(42).AllowZeroLength = True
         
         .Fields.Append .CreateField("DLSDocDesc5", dbText, 60)
         .Fields(43).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocLoc5", dbText, 4)
         .Fields(44).AllowZeroLength = True
         
         'More
         .Fields.Append .CreateField("DLSDocEco1", dbText, 2)
         .Fields(45).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocAdcn1", dbText, 20)
         .Fields(46).AllowZeroLength = True
         
         .Fields.Append .CreateField("DLSDocEco2", dbText, 2)
         .Fields(47).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocAdcn2", dbText, 20)
         .Fields(48).AllowZeroLength = True
         
         .Fields.Append .CreateField("DLSDocEco3", dbText, 2)
         .Fields(49).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocAdcn3", dbText, 20)
         .Fields(50).AllowZeroLength = True
         
         .Fields.Append .CreateField("DLSDocEco4", dbText, 2)
         .Fields(51).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocAdcn4", dbText, 20)
         .Fields(52).AllowZeroLength = True
         
         .Fields.Append .CreateField("DLSDocEco5", dbText, 2)
         .Fields(53).AllowZeroLength = True
         .Fields.Append .CreateField("DLSDocAdcn5", dbText, 20)
         .Fields(54).AllowZeroLength = True
         
      End With
      'add the table to Jet. No indexes
      On Error GoTo DiaErr1
      JetDb.TableDefs.Append NewTb
   Else
      JetDb.Execute "DELETE * FROM CvrTable"
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "createcvrt"
   bTablesCreated = 0
   
End Sub

Private Sub CreatePlsTable()
   'Drop and create a Jet table to
   'run beside SQL Server so Crystal can handle the report
   '(Crystal isn't smart enough to handle the joins)
   Dim NewTb As TableDef
   Dim NewFld As Field
   
   On Error Resume Next
   JetDb.Execute "SELECT PLSPart1 FROM PlsTable"
   If Err > 0 Then
      Set NewTb = JetDb.CreateTableDef("PlsTable")
      'Fields. Note that we allow empties
      With NewTb
         'Documents
         '1
         .Fields.Append .CreateField("PLSPart1", dbText, 30)
         .Fields(0).AllowZeroLength = True
         .Fields.Append .CreateField("PLSDesc1", dbText, 30)
         .Fields(1).AllowZeroLength = True
         .Fields.Append .CreateField("PLSADate1", dbDate)
         .Fields.Append .CreateField("PLSAQty1", dbCurrency)
         .Fields(3).DefaultValue = 0
         .Fields.Append .CreateField("PLSUom1", dbText, 2)
         .Fields(4).AllowZeroLength = True
         .Fields.Append .CreateField("PLSLoc1", dbText, 4)
         .Fields(5).AllowZeroLength = True
         
         '2
         .Fields.Append .CreateField("PLSPart2", dbText, 30)
         .Fields(6).AllowZeroLength = True
         .Fields.Append .CreateField("PLSDesc2", dbText, 30)
         .Fields(7).AllowZeroLength = True
         .Fields.Append .CreateField("PLSADate2", dbDate)
         .Fields.Append .CreateField("PLSAQty2", dbCurrency)
         .Fields(9).DefaultValue = 0
         .Fields.Append .CreateField("PLSUom2", dbText, 2)
         .Fields(10).AllowZeroLength = True
         .Fields.Append .CreateField("PLSLoc2", dbText, 4)
         .Fields(11).AllowZeroLength = True
         
         '3
         .Fields.Append .CreateField("PLSPart3", dbText, 30)
         .Fields(12).AllowZeroLength = True
         .Fields.Append .CreateField("PLSDesc3", dbText, 30)
         .Fields(13).AllowZeroLength = True
         .Fields.Append .CreateField("PLSADate3", dbDate)
         .Fields.Append .CreateField("PLSAQty3", dbCurrency)
         .Fields(15).DefaultValue = 0
         .Fields.Append .CreateField("PLSUom3", dbText, 2)
         .Fields(16).AllowZeroLength = True
         .Fields.Append .CreateField("PLSLoc3", dbText, 4)
         .Fields(17).AllowZeroLength = True
         
         '4
         .Fields.Append .CreateField("PLSPart4", dbText, 30)
         .Fields(18).AllowZeroLength = True
         .Fields.Append .CreateField("PLSDesc4", dbText, 30)
         .Fields(19).AllowZeroLength = True
         .Fields.Append .CreateField("PLSADate4", dbDate)
         .Fields.Append .CreateField("PLSAQty4", dbCurrency)
         .Fields(21).DefaultValue = 0
         .Fields.Append .CreateField("PLSUom4", dbText, 2)
         .Fields(22).AllowZeroLength = True
         .Fields.Append .CreateField("PLSLoc4", dbText, 4)
         .Fields(23).AllowZeroLength = True
         
         '5
         .Fields.Append .CreateField("PLSPart5", dbText, 30)
         .Fields(24).AllowZeroLength = True
         .Fields.Append .CreateField("PLSDesc5", dbText, 30)
         .Fields(25).AllowZeroLength = True
         .Fields.Append .CreateField("PLSADate5", dbDate)
         .Fields.Append .CreateField("PLSAQty5", dbCurrency)
         .Fields(27).DefaultValue = 0
         .Fields.Append .CreateField("PLSUom5", dbText, 2)
         .Fields(28).AllowZeroLength = True
         .Fields.Append .CreateField("PLSLoc5", dbText, 4)
         .Fields(29).AllowZeroLength = True
         
         '6
         .Fields.Append .CreateField("PLSPart6", dbText, 30)
         .Fields(30).AllowZeroLength = True
         .Fields.Append .CreateField("PLSDesc6", dbText, 30)
         .Fields(31).AllowZeroLength = True
         .Fields.Append .CreateField("PLSADate6", dbDate)
         .Fields.Append .CreateField("PLSAQty6", dbCurrency)
         .Fields(33).DefaultValue = 0
         .Fields.Append .CreateField("PLSUom6", dbText, 2)
         .Fields(34).AllowZeroLength = True
         .Fields.Append .CreateField("PLSLoc6", dbText, 4)
         .Fields(35).AllowZeroLength = True
         
         'Added
         .Fields.Append .CreateField("PLSPQty1", dbCurrency)
         .Fields(36).DefaultValue = 0
         .Fields.Append .CreateField("PLSPQty2", dbCurrency)
         .Fields(37).DefaultValue = 0
         .Fields.Append .CreateField("PLSPQty3", dbCurrency)
         .Fields(38).DefaultValue = 0
         .Fields.Append .CreateField("PLSPQty4", dbCurrency)
         .Fields(39).DefaultValue = 0
         .Fields.Append .CreateField("PLSPQty5", dbCurrency)
         .Fields(40).DefaultValue = 0
         .Fields.Append .CreateField("PLSPQty6", dbCurrency)
         .Fields(41).DefaultValue = 0
      End With
      'add the table to Jet. No indexes
      JetDb.TableDefs.Append NewTb
   Else
      JetDb.Execute "DELETE * FROM CvrTable"
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "createplst"
   bTablesCreated = 0
   
End Sub

Private Sub BuildDocumentList()
   Dim RdoDoc As rdoResultset
   Dim RdoJet As rdoResultset
   Dim iRow As Integer
   
   On Error GoTo DiaErr1
   JetDb.Execute "DELETE * FROM CvrTable"
   Set DbDoc = JetDb.OpenRecordset("CvrTable", dbOpenDynaset)
   '   Well it ain't buying this join. SQL Server will but not the RDO
   '
   '    sSql = "SELECT DOREF,DONUM,DOREV,DOCLASS,DOSHEET,DODESCR," _
   '        & "DOECO,DOADCN,DOTYPE,DLSREF,DLSREV,DLSTYPE,DLSDOCREF," _
   '        & "DLSDOCREV,DLSDOCSHEET,DLSDOCCLASS FROM DdocTable,DlstTable" _
   '        & "WHERE (DOREF=DLSDOCREF AND DOSHEET=DLSDOCSHEET AND DOREV=" _
   '        & "DLSREV AND DLSREF='65B845892')"
   '
   '   so
   '     sSql = "SELECT DLSREF,DLSREV,DLSTYPE,DLSDOCREF," _
   '         & "DLSDOCREV,DLSDOCSHEET,DLSDOCCLASS FROM DlstTable " _
   '         & "WHERE DLSREF='" & Compress(cmbPrt) & "' ORDER BY DLSDOCREF"
   
   '5/06/04 Changed to static table
   sSql = "SELECT * FROM RndlTable WHERE RUNDLSRUNREF='" & Compress(cmbPrt) & "' " _
          & "AND RUNDLSRUNNO=" & Val(cmbRun) & " ORDER BY RUNDLSNUM"
   bSqlRows = GetDataSet(RdoDoc, ES_FORWARD)
   If bSqlRows Then
      DbDoc.AddNew
      With RdoDoc
         Do Until .EOF
            iRow = iRow + 1
            Err = 0
            Select Case iRow
               Case 1
                  DbDoc!DLSDocRef1 = "" & Trim(RdoDoc!RUNDLSDOCREFLONG)
                  DbDoc!DLSDocRev1 = "" & Trim(RdoDoc!RUNDLSDOCREV)
                  DbDoc!DLSDocSheet1 = "" & Trim(RdoDoc!RUNDLSDOCREFSHEET)
                  DbDoc!DLSDocClass1 = "" & Trim(RdoDoc!RUNDLSDOCREFCLASS)
                  DbDoc!DLSDocDesc1 = "" & Trim(RdoDoc!RUNDLSDOCREFDESC)
                  DbDoc!DLSDocLoc1 = ""
                  DbDoc!DLSDocEco1 = "" & Trim(RdoDoc!RUNDLSDOCREFECO)
                  DbDoc!DLSDocAdcn1 = "" & Trim(RdoDoc!RUNDLSDOCREFADCN)
               Case 2
                  DbDoc!DLSDocRef2 = "" & Trim(RdoDoc!RUNDLSDOCREFLONG)
                  DbDoc!DLSDocRev2 = "" & Trim(RdoDoc!RUNDLSDOCREV)
                  DbDoc!DLSDocSheet2 = "" & Trim(RdoDoc!RUNDLSDOCREFSHEET)
                  DbDoc!DLSDocClass2 = "" & Trim(RdoDoc!RUNDLSDOCREFCLASS)
                  DbDoc!DLSDocDesc2 = "" & Trim(RdoDoc!RUNDLSDOCREFDESC)
                  DbDoc!DLSDocLoc2 = ""
                  DbDoc!DLSDocEco2 = "" & Trim(RdoDoc!RUNDLSDOCREFECO)
                  DbDoc!DLSDocAdcn2 = "" & Trim(RdoDoc!RUNDLSDOCREFADCN)
               Case 3
                  DbDoc!DLSDocRef3 = "" & Trim(RdoDoc!RUNDLSDOCREFLONG)
                  DbDoc!DLSDocRev3 = "" & Trim(RdoDoc!RUNDLSDOCREV)
                  DbDoc!DLSDocSheet3 = "" & Trim(RdoDoc!RUNDLSDOCREFSHEET)
                  DbDoc!DLSDocClass3 = "" & Trim(RdoDoc!RUNDLSDOCREFCLASS)
                  DbDoc!DLSDocDesc3 = "" & Trim(RdoDoc!RUNDLSDOCREFDESC)
                  DbDoc!DLSDocLoc3 = ""
                  DbDoc!DLSDocEco3 = "" & Trim(RdoDoc!RUNDLSDOCREFECO)
                  DbDoc!DLSDocAdcn3 = "" & Trim(RdoDoc!RUNDLSDOCREFADCN)
               Case 4
                  DbDoc!DLSDocRef4 = "" & Trim(RdoDoc!RUNDLSDOCREFLONG)
                  DbDoc!DLSDocRev4 = "" & Trim(RdoDoc!RUNDLSDOCREV)
                  DbDoc!DLSDocSheet4 = "" & Trim(RdoDoc!RUNDLSDOCREFSHEET)
                  DbDoc!DLSDocClass4 = "" & Trim(RdoDoc!RUNDLSDOCREFCLASS)
                  DbDoc!DLSDocDesc4 = "" & Trim(RdoDoc!RUNDLSDOCREFDESC)
                  DbDoc!DLSDocLoc4 = ""
                  DbDoc!DLSDocEco4 = "" & Trim(RdoDoc!RUNDLSDOCREFECO)
                  DbDoc!DLSDocAdcn4 = "" & Trim(RdoDoc!RUNDLSDOCREFADCN)
               Case 5
                  DbDoc!DLSDocRef5 = "" & Trim(RdoDoc!RUNDLSDOCREFLONG)
                  DbDoc!DLSDocRev5 = "" & Trim(RdoDoc!RUNDLSDOCREV)
                  DbDoc!DLSDocSheet5 = "" & Trim(RdoDoc!RUNDLSDOCREFSHEET)
                  DbDoc!DLSDocClass5 = "" & Trim(RdoDoc!RUNDLSDOCREFCLASS)
                  DbDoc!DLSDocDesc5 = "" & Trim(RdoDoc!RUNDLSDOCREFDESC)
                  DbDoc!DLSDocLoc5 = ""
                  DbDoc!DLSDocEco5 = "" & Trim(RdoDoc!RUNDLSDOCREFECO)
                  DbDoc!DLSDocAdcn5 = "" & Trim(RdoDoc!RUNDLSDOCREFADCN)
               Case Else
                  Exit Do
            End Select
            MsgBox Err.Description
            .MoveNext
         Loop
         DbDoc.Update
         .Cancel
      End With
   Else
      DbDoc.AddNew
      DbDoc!DLSDocRef1 = "*** No Documents Recorded ***"
      DbDoc.Update
   End If
   On Error Resume Next
   DbDoc.Close
   Set RdoDoc = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "builddoclist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'SC or RL...no Pick List yet

Private Sub BuildPartsList()
   Dim RdoLst As rdoResultset
   Dim b As Byte
   Dim iPkRecord As Integer
   Dim cConversion As Currency
   Dim cQuantity As Currency
   Dim cSetup As Currency
   Dim sUnits As String
   
   sSql = "SELECT DISTINCT BMASSYPART FROM BmplTable " _
          & "WHERE BMASSYPART='" & Compress(cmbPrt) & "' "
   bSqlRows = GetDataSet(RdoLst, ES_FORWARD)
   If bSqlRows Then
      b = 1
      RdoLst.Cancel
   Else
      MouseCursor 0
      b = 0
      MsgBox "This Part Does Not Have A Parts List.", vbExclamation, Caption
   End If
   If b = 1 Then
      Set DbPls = JetDb.OpenRecordset("PlsTable", dbOpenDynaset)
      On Error GoTo DiaErr1
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALOCATION," _
             & "BMASSYPART,BMPARTREF,BMREV,BMQTYREQD,BMSETUP,BMADDER," _
             & "BMCONVERSION,BMUNITS FROM PartTable,BmplTable WHERE (" _
             & "PARTREF=BMPARTREF AND BMASSYPART='" & Compress(cmbPrt) & "' " _
             & "AND BMREV='" & sBomRev & "') "
      bSqlRows = GetDataSet(RdoLst, ES_FORWARD)
      If bSqlRows Then
         b = 0
         With RdoLst
            DbPls.AddNew
            On Error Resume Next
            RdoCon.BeginTrans
            Do Until .EOF
               If Not IsNull(!BMSETUP) Then
                  cSetup = !BMSETUP
               Else
                  cSetup = 0
               End If
               sUnits = "" & Trim(!BMUNITS)
               cQuantity = Format(!BMQTYREQD + !BMADDER, ES_QuantityDataFormat)
               cConversion = Format(!BMCONVERSION, "#####0.0000")
               If cConversion = 0 Then cConversion = 1
               cQuantity = cQuantity / cConversion
               cQuantity = (cQuantity * Val(lblQty)) + cSetup
               b = b + 1
               If b < 7 Then
                  Select Case b
                     Case 1
                        DbPls!PLSPart1 = "" & Trim(!PARTNUM)
                        DbPls!PLSDesc1 = "" & Trim(!PADESC)
                        DbPls!PLSPqty1 = cQuantity
                        DbPls!PLSUom1 = "" & Trim(!BMUNITS)
                        DbPls!PLSLoc1 = "" & Trim(!PALOCATION)
                     Case 2
                        DbPls!PLSPart2 = "" & Trim(!PARTNUM)
                        DbPls!PLSDesc2 = "" & Trim(!PADESC)
                        DbPls!PLSPqty2 = cQuantity
                        DbPls!PLSUom2 = "" & Trim(!BMUNITS)
                        DbPls!PLSLoc2 = "" & Trim(!PALOCATION)
                     Case 3
                        DbPls!PLSPart3 = "" & Trim(!PARTNUM)
                        DbPls!PLSDesc3 = "" & Trim(!PADESC)
                        DbPls!PLSPqty3 = cQuantity
                        DbPls!PLSUom3 = "" & Trim(!BMUNITS)
                        DbPls!PLSLoc3 = "" & Trim(!PALOCATION)
                     Case 4
                        DbPls!PLSPart4 = "" & Trim(!PARTNUM)
                        DbPls!PLSDesc4 = "" & Trim(!PADESC)
                        DbPls!PLSPqty4 = cQuantity
                        DbPls!PLSUom4 = "" & Trim(!BMUNITS)
                        DbPls!PLSLoc4 = "" & Trim(!PALOCATION)
                     Case 5
                        DbPls!PLSPart5 = "" & Trim(!PARTNUM)
                        DbPls!PLSDesc5 = "" & Trim(!PADESC)
                        DbPls!PLSPqty5 = cQuantity
                        DbPls!PLSUom5 = "" & Trim(!BMUNITS)
                        DbPls!PLSLoc5 = "" & Trim(!PALOCATION)
                     Case 6
                        DbPls!PLSPart6 = "" & Trim(!PARTNUM)
                        DbPls!PLSDesc6 = "" & Trim(!PADESC)
                        DbPls!PLSPqty6 = cQuantity
                        DbPls!PLSUom6 = "" & Trim(!BMUNITS)
                        DbPls!PLSLoc6 = "" & Trim(!PALOCATION)
                  End Select
               End If
               If sRunPkstart = "" Then sRunPkstart = Format(ES_SYSDATE, "mm/dd/yy")
               iPkRecord = iPkRecord + 1
               sSql = "INSERT INTO MopkTable (PKPARTREF," _
                      & "PKMOPART,PKMORUN,PKTYPE,PKPDATE," _
                      & "PKPQTY,PKBOMQTY,PKRECORD,PKUNITS,PKCOMT) VALUES('" _
                      & Trim(!PartRef) & "','" & Compress(cmbPrt) & "'," _
                      & cmbRun & ",9,'" & sRunPkstart & "'," & cQuantity _
                      & "," & cQuantity & "," & iPkRecord & ",'" & sUnits & "','') "
               RdoCon.Execute sSql, rdExecDirect
               .MoveNext
            Loop
            sSql = "UPDATE RunsTable SET RUNSTATUS='PL'," _
                   & "RUNPLDATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "' " _
                   & "WHERE RUNREF='" & Compress(cmbPrt) & "' " _
                   & "AND RUNNO=" & cmbRun & " "
            RdoCon.Execute sSql, rdExecDirect
            If Err = 0 Then
               RdoCon.CommitTrans
               lblSta = "PL"
            Else
               RdoCon.RollbackTrans
            End If
            DbPls.Update
            .Cancel
         End With
      End If
   Else
      DbPls.AddNew
      DbPls!PLSPart1 = "*** No Pick List Recorded ***"
      DbPls.Update
   End If
   On Error Resume Next
   DbPls.Close
   Set RdoLst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "buildpartslist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Cover and pick list

Private Sub PrintCover()
   Dim sWindows As String
   MouseCursor 13
   
   On Error Resume Next
   DbPls.Close
   DbDoc.Close
   
   On Error GoTo DiaErr1
   DoEvents
   sWindows = GetWindowsDir()
   SetMdiReportsize MdiSect
   sProcName = "printcover"
   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   sCustomReport = GetCustomReport("prdshcvr")
   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.Crw.DataFiles(0) = sWindows & "\temp\esiprod.mdb"
   MdiSect.Crw.Formulas(1) = "Includes='" & cmbPrt & " Run " & cmbRun & "'"
   MdiSect.Crw.Formulas(2) = "Includes2='" & lblDsc & "'"
   If optDoc.Value = vbUnchecked Then
      MdiSect.Crw.SectionFormat(0) = "GROUPHDR.0.0;F;;;"
      MdiSect.Crw.SectionFormat(1) = "GROUPHDR.0.1;F;;;"
   Else
      MdiSect.Crw.SectionFormat(0) = "GROUPHDR.0.0;T;;;"
      MdiSect.Crw.SectionFormat(1) = "GROUPHDR.0.1;F;;;"
   End If
   If optLst.Value = vbUnchecked Then
      MdiSect.Crw.SectionFormat(2) = "GROUPFTR.0.1;F;;;"
      MdiSect.Crw.SectionFormat(3) = "GROUPFTR.0.2;F;;;"
   Else
      MdiSect.Crw.SectionFormat(2) = "GROUPFTR.0.1;T;;;"
      MdiSect.Crw.SectionFormat(3) = "GROUPFTR.0.2;T;;;"
   End If
   'etCrystalAction Me
   MdiSect.Crw.Action = 1
   DoEvents
   On Error Resume Next
   ReopenJet
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Psh02
Psh02:
   DoModuleErrors Me
   
End Sub

'Pick list is active

Private Sub BuildPickList()
   Dim RdoLst As rdoResultset
   Dim b As Byte
   
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALOCATION," _
          & "PKPARTREF,PKMOPART,PKPQTY,PKADATE,PKAQTY,PKUNITS FROM PartTable," _
          & "MopkTable WHERE (PARTREF=PKPARTREF AND PKMOPART='" _
          & Compress(cmbPrt) & "' AND PKMORUN=" & cmbRun & ")"
   bSqlRows = GetDataSet(RdoLst, ES_FORWARD)
   If bSqlRows Then
      JetDb.Execute "DELETE * FROM PlsTable"
      Set DbPls = JetDb.OpenRecordset("PlsTable", dbOpenDynaset)
      With RdoLst
         DbPls.AddNew
         Do Until .EOF
            b = b + 1
            If b < 7 Then
               Select Case b
                  Case 1
                     DbPls!PLSPart1 = "" & Trim(!PARTNUM)
                     DbPls!PLSDesc1 = "" & Trim(!PADESC)
                     DbPls!PLSPqty1 = Format(!PKPQTY, ES_QuantityDataFormat)
                     DbPls!PLSAqty1 = Format(!PKAQTY, ES_QuantityDataFormat)
                     If Not IsNull(!PKADATE) Then
                        DbPls!PLSADate1 = Format(!PKADATE, "mm/dd/yy")
                     End If
                     DbPls!PLSUom1 = "" & Trim(!PKUNITS)
                     DbPls!PLSLoc1 = "" & Trim(!PALOCATION)
                  Case 2
                     DbPls!PLSPart2 = "" & Trim(!PARTNUM)
                     DbPls!PLSDesc2 = "" & Trim(!PADESC)
                     DbPls!PLSPqty2 = Format(!PKPQTY, ES_QuantityDataFormat)
                     If Not IsNull(!PKADATE) Then
                        DbPls!PLSADate2 = Format(!PKADATE, "mm/dd/yy")
                     End If
                     DbPls!PLSAqty2 = Format(!PKAQTY, ES_QuantityDataFormat)
                     DbPls!PLSUom2 = "" & Trim(!PKUNITS)
                     DbPls!PLSLoc2 = "" & Trim(!PALOCATION)
                  Case 3
                     DbPls!PLSPart3 = "" & Trim(!PARTNUM)
                     DbPls!PLSDesc3 = "" & Trim(!PADESC)
                     DbPls!PLSPqty3 = Format(!PKPQTY, ES_QuantityDataFormat)
                     If Not IsNull(!PKADATE) Then
                        DbPls!PLSADate3 = Format(!PKADATE, "mm/dd/yy")
                     End If
                     DbPls!PLSAqty3 = Format(!PKAQTY, ES_QuantityDataFormat)
                     DbPls!PLSUom3 = "" & Trim(!PKUNITS)
                     DbPls!PLSLoc3 = "" & Trim(!PALOCATION)
                  Case 4
                     DbPls!PLSPart4 = "" & Trim(!PARTNUM)
                     DbPls!PLSDesc4 = "" & Trim(!PADESC)
                     DbPls!PLSPqty4 = Format(!PKPQTY, ES_QuantityDataFormat)
                     If Not IsNull(!PKADATE) Then
                        DbPls!PLSADate4 = Format(!PKADATE, "mm/dd/yy")
                     End If
                     DbPls!PLSAqty4 = Format(!PKAQTY, ES_QuantityDataFormat)
                     DbPls!PLSUom4 = "" & Trim(!PKUNITS)
                     DbPls!PLSLoc4 = "" & Trim(!PALOCATION)
                  Case 5
                     DbPls!PLSPart5 = "" & Trim(!PARTNUM)
                     DbPls!PLSDesc5 = "" & Trim(!PADESC)
                     DbPls!PLSPqty5 = Format(!PKPQTY, ES_QuantityDataFormat)
                     If Not IsNull(!PKADATE) Then
                        DbPls!PLSADate5 = Format(!PKADATE, "mm/dd/yy")
                     End If
                     DbPls!PLSAqty5 = Format(!PKAQTY, ES_QuantityDataFormat)
                     DbPls!PLSUom5 = "" & Trim(!PKUNITS)
                     DbPls!PLSLoc5 = "" & Trim(!PALOCATION)
                  Case 6
                     DbPls!PLSPart6 = "" & Trim(!PARTNUM)
                     DbPls!PLSDesc6 = "" & Trim(!PADESC)
                     DbPls!PLSPqty6 = Format(!PKPQTY, ES_QuantityDataFormat)
                     If Not IsNull(!PKADATE) Then
                        DbPls!PLSADate6 = Format(!PKADATE, "mm/dd/yy")
                     End If
                     DbPls!PLSAqty6 = Format(!PKAQTY, ES_QuantityDataFormat)
                     DbPls!PLSUom6 = "" & Trim(!PKUNITS)
                     DbPls!PLSLoc6 = "" & Trim(!PALOCATION)
               End Select
            End If
            .MoveNext
         Loop
         DbPls.Update
         .Cancel
      End With
   Else
      DbPls.AddNew
      DbPls!PLSPart1 = "*** No Pick List Recorded ***"
      DbPls.Update
   End If
   DoEvents
   On Error Resume Next
   DbPls.Close
   Set RdoLst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "buildpicklist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintAllocations()
'   MouseCursor 13
'   On Error GoTo Psh01
'   SetMdiReportsize MdiSect
'   sProcName = "printalloc"
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   sCustomReport = GetCustomReport("prdsh17")
'   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
'   sSql = "{RunsTable.RUNREF}='" & sPartNumber & "' " _
'          & "AND {RunsTable.RUNNO}=" & Trim(cmbRun) & " "
'   MdiSect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
'   MouseCursor 0
'   Exit Sub
'
'Psh01:
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   Resume Psh02
'Psh02:
'   DoModuleErrors Me
'
   PrintReportSalesOrderAllocations Me, sPartNumber, cmbRun
End Sub





'Problems creating tables from a popup 10/18/01

Private Sub CreateJetTables()
   bTablesCreated = 1
   CreatePlsTable
   CreateCovrTable
   DoEvents
   
End Sub
