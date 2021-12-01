VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaPhu05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Weekly Time Charges (Report)"
   ClientHeight    =   2520
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2520
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbDept 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   14
      Tag             =   "1"
      ToolTipText     =   "Select From List Or Enter Number"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Tag             =   "4"
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox cmbEmp 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Select From List Or Enter Number"
      Top             =   1080
      Width           =   1095
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
         Picture         =   "diaPhu05.frx":0000
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
         Picture         =   "diaPhu05.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   2
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
      PictureUp       =   "diaPhu05.frx":0308
      PictureDn       =   "diaPhu05.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6300
      Top             =   1260
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2520
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   15
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   13
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Week Ending"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   360
      Width           =   1635
   End
   Begin VB.Label lblWen 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date range"
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      ToolTipText     =   "Week Ending (System Administration Setup)"
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Range"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "All Employees."
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "diaPhu05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Dim bOnLoad As Byte
Dim bGoodEmp As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbEmp_Click()
   bGoodEmp = GetEmployee
   
End Sub

Private Sub cmbEmp_LostFocus()
   cmbEmp = CheckLen(cmbEmp, 6)
   If Val(Len(cmbEmp)) Then
      cmbEmp = Format(cmbEmp, "000000")
      bGoodEmp = GetEmployee()
      cmbDept = "ALL"
   Else
      cmbEmp = "ALL"
   End If
   
End Sub

Private Sub cmbDept_LostFocus()
   If cmbDept = "" Then
      cmbDept = "ALL"
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbEmp = ""
   
End Sub


Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs907"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub FillEmployees()
   cmbEmp.Clear
   sSql = "SELECT DISTINCT TMEMP FROM TchdTable " & vbCrLf _
      & "WHERE TMDAY >= '" & DateAdd("d", -6, txtDte) & "' AND TMDAY <='" & txtDte & "'" & vbCrLf _
      & "order by TMEMP"
   LoadNumComboBox cmbEmp, "000000"
   If bSqlRows Then
      cmbEmp = "ALL"
      cmbEmp_Click
   Else
      lblNme = "No Time Charges In The Period."
   End If
   cmbEmp = "ALL"
End Sub

Private Sub FillDepartment()
   cmbDept.Clear
   sSql = "select distinct PREMDEPT from emplTable where PREMDEPT <> '' AND PREMDEPT IS NOT NULL" & vbCrLf _
      & "order by PREMDEPT"
   LoadNumComboBox cmbDept, "000000"
   If bSqlRows Then
      cmbDept = "ALL"
   End If
   cmbDept = "ALL"
End Sub



Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillEmployees
      FillDepartment
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   'txtDte = Format(Now - 1, "mm/dd/yy")
   'If sCurrDate = "" Then
   '   If Format(txtDte, "w") = 1 Then
   '      txtDte = Format(Now - 2, "mm/dd/yy")
   '   End If
   'Else
   '   txtDte = sCurrDate
   'End If
   GetWeekEnd
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
   Set diaPhu05 = Nothing
   
End Sub




Private Sub PrintReport()
   Dim sDate As String
   
   If Len(Trim(cmbEmp)) = 0 Then cmbEmp = "ALL"
   'sDate = Format(lblWen, "yyyy,mm,dd")
   MouseCursor 13
   On Error GoTo DiaErr1
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("admhu05")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbEmp) & " for " & CStr(lblWen) & "'")
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{TchdTable.TMDAY} in" & CrystalDate(DateAdd("d", -6, txtDte)) & " to " & CrystalDate(txtDte)
   If cmbEmp <> "ALL" Then
      sSql = sSql & "AND {TchdTable.TMEMP}=" & cmbEmp & " "
   End If
   
   If cmbDept <> "ALL" Then
      sSql = sSql & "AND {EmplTable.PREMDEPT}='" & cmbDept & "'"
   End If
   
   
   cCRViewer.SetReportSelectionFormula sSql
   'cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport1()
   Dim sDate As String
   
   If Len(Trim(cmbEmp)) = 0 Then cmbEmp = "ALL"
   'sDate = Format(lblWen, "yyyy,mm,dd")
   MouseCursor 13
   On Error GoTo DiaErr1
   'SetMdiReportsize MDISect
   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MDISect.Crw.Formulas(1) = "Includes='" & cmbEmp & " for " & lblWen & "'"
   sCustomReport = GetCustomReport("admhu05")
   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   'sSql = "{TchdTable.TMWEEK}=Date(" & sDate & ") "
   'sSql = "{TchdTable.TMWEEK}=" & CrystalDate(lblWen)
   sSql = "{TchdTable.TMDAY} in" & CrystalDate(DateAdd("d", -6, txtDte)) & " to " & CrystalDate(txtDte)
   If cmbEmp <> "ALL" Then
      sSql = sSql & "AND {TchdTable.TMEMP}=" & cmbEmp & " "
   End If
   MDISect.Crw.SelectionFormula = sSql
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub













Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   'remember the day of week that the week ends
   
   Dim dayWeekEnds As Integer
   dayWeekEnds = DatePart("w", txtDte)
   
   Dim sOptions As String
   sOptions = dayWeekEnds + "0000000"
   SaveSetting "Esi2000", "EsiTime", "diaPhu05", sOptions
   
End Sub

Private Sub GetOptions()

   'select most recent week-ending date
   Dim sOptions As String
   sOptions = GetSetting("Esi2000", "EsiTime", "diaPhu05", "70000000")
   Dim dayWeekEnds As Integer
   dayWeekEnds = CInt(Mid(sOptions, 1, 1))
   
'   Dim msg As String
'   Dim j As Integer
'   For j = 1 To 7
'   msg = ""
'   For dayWeekEnds = 1 To 7
   
   Dim weekEndingDate As Date, today As Date
   today = Format(Now, "mm/dd/yyyy")
'   today = DateAdd("d", j - 1, today)
   
   'if today is not the week-ending date, go backwards
   Dim todaysDayOfWeek As Integer
   Dim daysFromWeekendToToday As Integer
   todaysDayOfWeek = DatePart("w", today)
   daysFromWeekendToToday = dayWeekEnds - todaysDayOfWeek
   If daysFromWeekendToToday > 0 Then
      daysFromWeekendToToday = daysFromWeekendToToday - 7
   End If
   weekEndingDate = DateAdd("d", daysFromWeekendToToday, today)
   txtDte = Format(weekEndingDate, "mm/dd/yy")
'   msg = msg & "week ends on day " & dayWeekEnds & ".  today (" & today & ") is " & todaysDayOfWeek & ".  week ending date = " & weekEndingDate & vbCrLf
'   Next
'   MsgBox msg
'   Next j
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub txtDte_Change()
   GetWeekEnd
End Sub

Private Sub txtDte_Click()
   ' GetWeekEnd
   
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub



Private Sub GetWeekEnd()
   'display week start and end dates
   Dim startDate As Variant
   Dim endDate As Variant
   endDate = txtDte
   startDate = DateAdd("d", -6, endDate)
   lblWen = WeekdayName(DatePart("w", startDate), True) & " " _
      & startDate & " - " & WeekdayName(DatePart("w", endDate), True) & " " _
      & endDate
   FillEmployees
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   GetWeekEnd
   
End Sub



Private Function GetEmployee() As Byte
   Dim RdoEmp As ADODB.Recordset
   On Error GoTo DiaErr1
   
   If cmbEmp <> "ALL" Then
      sSql = "Qry_EmployeeName " & Val(cmbEmp)
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp)
      If bSqlRows Then
         With RdoEmp
            cmbEmp = Format(!PREMNUMBER, "000000")
            lblNme = "" & Trim(!PREMLSTNAME) & ", " _
                     & Trim(!PREMFSTNAME) & " " _
                     & Trim(!PREMMINIT)
            'lblSsn = "" & Trim(!PREMSOCSEC)
            .Cancel
            GetEmployee = True
            sCurrEmployee = cmbEmp
         End With
      Else
         GetEmployee = False
         lblNme = "No Current Employee"
         'lblSsn = ""
      End If
   Else
      lblNme = "All Employees."
      'lblSsn = ""
   End If
   Set RdoEmp = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getemploy"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
