VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaPhu17 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Time Charges (Report)"
   ClientHeight    =   3570
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3570
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbDept 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   18
      Tag             =   "1"
      ToolTipText     =   "Select From List Or Enter Number"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.ComboBox cmbGroupBy 
      Height          =   315
      ItemData        =   "diaPhu17.frx":0000
      Left            =   2040
      List            =   "diaPhu17.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.ComboBox txtEndDte 
      Height          =   315
      Left            =   3960
      TabIndex        =   14
      Top             =   360
      Width           =   1095
   End
   Begin VB.CheckBox cbShowDetails 
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   2760
      Width           =   375
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
      Top             =   840
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
         Picture         =   "diaPhu17.frx":002D
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
         Picture         =   "diaPhu17.frx":01AB
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
      PictureUp       =   "diaPhu17.frx":0335
      PictureDn       =   "diaPhu17.frx":047B
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
      FormDesignHeight=   3570
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   7
      Left            =   3840
      TabIndex        =   20
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Group report by"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   15
      Top             =   360
      Width           =   315
   End
   Begin VB.Label Label1 
      Caption         =   "Show Details?"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   11
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Date"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "All Employees."
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "diaPhu17"
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
      & "WHERE TMDAY ='" & txtDte & "'" & vbCrLf _
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
   Set diaPhu17 = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sDate As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   
   If Len(Trim(cmbEmp)) = 0 Then cmbEmp = "ALL"
   MouseCursor 13
   On Error GoTo DiaErr1
  
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("admhu17")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "ShowDetails"
   aFormulaName.Add "GroupBy"
   aFormulaName.Add "DateRangeSelected"
   aFormulaValue.Add "'" & sFacility & "'"
   aFormulaValue.Add "'" & cmbEmp & " for " & txtDte & " through " & txtEndDte & "'"
   aFormulaValue.Add CStr(LTrim(str(cbShowDetails.Value)))
   aFormulaValue.Add CStr(LTrim(str(cmbGroupBy.ListIndex)))
   If txtDte = txtEndDte Then aFormulaValue.Add "0" Else aFormulaValue.Add "1"
  
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{TchdTable.TMDAY} >=" & CrystalDate(txtDte) & " AND {TchdTable.TMDAY} <=" & CrystalDate(txtEndDte)
   If cmbEmp <> "ALL" Then
      sSql = sSql & "AND {TchdTable.TMEMP}=" & cmbEmp & " "
   End If
   
   If cmbDept <> "ALL" Then
      sSql = sSql & "AND {EmplTable.PREMDEPT}='" & cmbDept & "'"
   End If
   
   cCRViewer.ShowGroupTree False
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
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

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiTime", "diaPhu17", txtDte + LTrim(str(cbShowDetails.Value)) + LTrim(str(cmbGroupBy.ListIndex)) + txtEndDte
End Sub

Private Sub GetOptions()

   'select most recent week-ending date
   Dim strSettings As String
   
   Dim strDate As String
   Dim strToday As String
   
   strToday = Format(Now, "mm/dd/yy")
   strSettings = GetSetting("Esi2000", "EsiTime", "diaPhu17")
    If Len(strSettings) = 0 Then
        txtDte = strToday
        cbShowDetails.Value = 0
    Else
        txtDte = Format(Mid(strSettings, 1, 8), "mm/dd/yy")
        If Len(strSettings) > 8 Then cbShowDetails.Value = Val(Mid(strSettings, 9, 1)) Else cbShowDetails.Value = 0
        If Len(strSettings) > 9 Then cmbGroupBy.ListIndex = Val(Mid(strSettings, 10, 1)) Else cmbGroupBy.ListIndex = 0
    End If
   
   If Len(Trim(txtEndDte)) = 0 Then txtEndDte = txtDte
   
'   strDate = GetSetting("Esi2000", "EsiTime", "diaPhu17", strToday)
'   txtDte = Format(strDate, "mm/dd/yy")
     
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

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   txtEndDte = txtDte
   FillEmployees
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



Private Sub txtEndDte_DropDown()
    ShowCalendar Me
End Sub



Private Sub txtEndDte_Validate(Cancel As Boolean)
    Dim sTemp1, sTemp2 As String
    sTemp1 = Format(txtEndDte, "yyyyMMdd")
    sTemp2 = Format(txtDte, "yyyyMMdd")
    If Len(sTemp1) = 0 Or Len(sTemp2) = 0 Then Exit Sub
    
    If sTemp2 > sTemp1 Then
        MsgBox "Start Date cannot be before end date"
        Cancel = True
    End If
End Sub
