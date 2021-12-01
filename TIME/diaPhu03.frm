VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaPhu03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Employee Time Charges"
   ClientHeight    =   2760
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2760
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbEmp 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Select From List Or Enter Number"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   4845
      TabIndex        =   2
      Tag             =   "4"
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
      TabIndex        =   5
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaPhu03.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaPhu03.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   0
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
      PictureUp       =   "diaPhu03.frx":0308
      PictureDn       =   "diaPhu03.frx":044E
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
      FormDesignHeight=   2760
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblSsn 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4125
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Card Date"
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "diaPhu03"
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

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbEmp_Click()
   GetEmployee
   
End Sub


Private Sub cmbEmp_LostFocus()
   cmbEmp = CheckLen(cmbEmp, 6)
   If Len(cmbEmp) Then
      cmbEmp = Format(cmbEmp, "000000")
      GetEmployee
   Else
      cmbEmp = "ALL"
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
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


Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillEmployees"
   LoadNumComboBox cmbEmp, "000000"
   If cmbEmp.ListCount > 0 Then
      If Trim(sCurrEmployee) = "" Then
         cmbEmp = cmbEmp.List(0)
      Else
         cmbEmp = sCurrEmployee
      End If
      GetEmployee
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   txtDte = Format(Now - 1, "mm/dd/yy")
   If sCurrDate = "" Then
      If Format(txtDte, "w") = 1 Then
         txtDte = Format(Now - 2, "mm/dd/yy")
      End If
   Else
      txtDte = sCurrDate
   End If
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   sCurrDate = txtDte
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaPhu03 = Nothing
   
End Sub


Private Sub PrintReport()
   Dim sDate As String
   sDate = Format(txtDte, "yyyy,mm,dd")
   If Len(Trim(cmbEmp)) = 0 Then cmbEmp = "ALL"
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("admhu03")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbEmp) & "'")
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{TchdTable.TMDAY}=Date(" & sDate & ") "
   If cmbEmp <> "ALL" Then
      sSql = sSql & " AND {TchdTable.TMEMP}=" & Val(cmbEmp) & " "
   End If
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
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
   sDate = Format(txtDte, "yyyy,mm,dd")
   If Len(Trim(cmbEmp)) = 0 Then cmbEmp = "ALL"
   
   MouseCursor 13
   On Error GoTo DiaErr1
   'SetMdiReportsize MDISect
   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MDISect.Crw.Formulas(1) = "Includes='" & cmbEmp & "'"
   sCustomReport = GetCustomReport("admhu03")
   MDISect.Crw.ReportFileName = sReportPath & sCustomReport

   sSql = "{TchdTable.TMDAY}=Date(" & sDate & ") "
   If cmbEmp <> "ALL" Then
      sSql = sSql & " AND {TchdTable.TMEMP}=" & Val(cmbEmp) & " "
   End If
   
'   sSql = "{TchdTable.TMEMP}=" & Val(cmbEmp) & " " _
'          & "AND {TchdTable.TMDAY}=Date(" & sDate & ") "
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
   
End Sub

Private Sub GetOptions()
   
End Sub

Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub



Private Sub GetEmployee()
   Dim RdoEmp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_EmployeeName " & Val(cmbEmp)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp)
   If bSqlRows Then
      With RdoEmp
         cmbEmp = Format(!PREMNUMBER, "000000")
         lblNme = "" & Trim(!PREMLSTNAME) & ", " _
                  & Trim(!PREMFSTNAME) & " " _
                  & Trim(!PREMMINIT)
         lblSsn = "" & Trim(!PREMSOCSEC)
         sCurrEmployee = cmbEmp
         .Cancel
      End With
   Else
      lblNme = "No Current Employee"
      lblSsn = ""
   End If
   Set RdoEmp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getemploy"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   
End Sub
