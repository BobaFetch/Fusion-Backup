VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CapaCPp07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work Centers Without Calendars"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPp07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbMon 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Month"
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox cmbYer 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      ToolTipText     =   "Year"
      Top             =   1440
      Width           =   855
   End
   Begin VB.ComboBox cmbWcn 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Or Blank For All"
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "CapaCPp07a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "CapaCPp07a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   252
      Index           =   3
      Left            =   2640
      TabIndex        =   10
      Top             =   1440
      Width           =   852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center(s)"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Select Or Enter A New Work Center"
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   4560
      TabIndex        =   7
      Top             =   960
      Width           =   1428
   End
End
Attribute VB_Name = "CapaCPp07a"
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

Private Sub cmbWcn_LostFocus()
   cmbWcn = CheckLen(cmbWcn, 12)
   If cmbWcn = "" Then cmbWcn = "ALL"
   
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
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad = 1 Then
      CreateTable
      FillCombo
   End If
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   sSql = "DROP TABLE CalendarNull"
   clsADOCon.ExecuteSQL sSql
   
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set CapaCPp07a = Nothing
   
End Sub
Private Sub PrintReport()
   Dim sBook As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Work Center(s)" & CStr(cmbWcn & " For " & cmbMon & " " & cmbYer) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("prdca14")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   SetCrystalAction Me
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
   cmbWcn = "ALL"
   
End Sub

Private Sub optDis_Click()
   FillTable
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   FillTable
   PrintReport
   
End Sub

Private Sub FillCombo()
   Dim A As Integer
   Dim iList As Integer
   sSql = "Qry_FillWorkCentersAll"
   LoadComboBox cmbWcn
   cmbMon.AddItem "Jan"
   cmbMon.AddItem "Feb"
   cmbMon.AddItem "Mar"
   cmbMon.AddItem "Apr"
   cmbMon.AddItem "May"
   cmbMon.AddItem "Jun"
   cmbMon.AddItem "Jul"
   cmbMon.AddItem "Aug"
   cmbMon.AddItem "Sep"
   cmbMon.AddItem "Oct"
   cmbMon.AddItem "Nov"
   cmbMon.AddItem "Dec"
   cmbMon = Format(ES_SYSDATE, "mmm")
   A = Format(ES_SYSDATE, "yyyy")
   For iList = A - 2 To A + 25
      AddComboStr cmbYer.hwnd, Format$(iList)
   Next
   cmbYer = Format(ES_SYSDATE, "yyyy")
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Make a temp table

Private Sub CreateTable()
   On Error Resume Next
   sSql = "Create Table CalendarNull " _
          & "(CALSHOPREF CHAR(12) NULL DEFAULT('')," _
          & "CALWCNREF CHAR(12) NULL DEFAULT('')," _
          & "CALSHOPNUM CHAR(12) NULL DEFAULT('')," _
          & "CALSHOPDSC CHAR(30) NULL DEFAULT('')," _
          & "CALWCNNUM CHAR(12) NULL DEFAULT('')," _
          & "CALWCNDSC CHAR(30) NULL DEFAULT(''))"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "CREATE  UNIQUE  CLUSTERED  INDEX CalRef ON " _
          & "CalendarNull(CALSHOPREF,CALWCNREF) WITH  FILLFACTOR = 80"
   clsADOCon.ExecuteSQL sSql
   
End Sub

Private Sub FillTable()
   Dim RdoFill As ADODB.Recordset
   Dim RdoCal As ADODB.Recordset
   Dim RdoNew As ADODB.Recordset
   Dim sWcn As String
   '  On Error Resume Next
   sSql = "TRUNCATE TABLE CalendarNull"
   clsADOCon.ExecuteSQL sSql
   
   If cmbWcn <> "ALL" Then sWcn = Compress(cmbWcn)
   sSql = "SELECT SHPREF,SHPNUM,SHPDESC,WCNREF,WCNNUM," _
          & "WCNSHOP,WCNDESC FROM ShopTable,WcntTable " _
          & "WHERE (SHPREF=WCNSHOP AND WCNREF LIKE '" _
          & sWcn & "%')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFill, ES_FORWARD)
   If bSqlRows Then
      With RdoFill
         Do Until .EOF
            sSql = "SELECT DISTINCT WCCREF,WCCSHOP,WCCCENTER FROM " _
                   & "WcclTable WHERE (WCCSHOP='" & Trim(!SHPREF) & "' AND " _
                   & "WCCCENTER='" & Trim(!WCNREF) & "') AND WCCREF='" _
                   & cmbMon & "-" & cmbYer & "' "
            bSqlRows = clsADOCon.GetDataSet(sSql, RdoCal, ES_FORWARD)
            If Not bSqlRows Then
               sSql = "select * from CalendarNull"
               bSqlRows = clsADOCon.GetDataSet(sSql, RdoNew, ES_DYNAMIC)
               With RdoNew
                  .AddNew
                  !CALSHOPREF = Trim(RdoFill!SHPREF)
                  !CALWCNREF = Trim(RdoFill!WCNREF)
                  !CALSHOPNUM = Trim(RdoFill!SHPNUM)
                  !CALSHOPDSC = Trim(RdoFill!SHPDESC)
                  !CALWCNNUM = Trim(RdoFill!WCNNUM)
                  !CALWCNDSC = Trim(RdoFill!WCNDESC)
                  .Update
               End With
            End If
            .MoveNext
         Loop
         ClearResultSet RdoFill
      End With
   End If
   Set RdoFill = Nothing
   Set RdoCal = Nothing
   Set RdoNew = Nothing
End Sub
