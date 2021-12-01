VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaPhu02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employees By Number"
   ClientHeight    =   2985
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2985
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optSortByHireDate 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   2640
      Width           =   975
   End
   Begin VB.CheckBox optShowComments 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ComboBox txtSta 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "Includes Codes In Use"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtEnd 
      Height          =   285
      Left            =   3600
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtEmp 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaPhu02.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaPhu02.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   5
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
      PictureUp       =   "diaPhu02.frx":0308
      PictureDn       =   "diaPhu02.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2985
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort by Hire Date"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Comments"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   14
      Top             =   1440
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   13
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   3
      Left            =   4440
      TabIndex        =   12
      Top             =   1080
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Detail"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   10
      Top             =   1080
      Width           =   345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employees From Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1905
   End
End
Attribute VB_Name = "diaPhu02"
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
Dim lFirstEmp As Long
Dim lLastEmp As Long

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs907"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   FillStatus
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaPhu02 = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sEmployee As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   If txtEmp <> "ALL" Then sEmployee = UCase(Trim(txtEmp))
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "ShowComments"
   aFormulaName.Add "SuppressDetails"
   aFormulaName.Add "GroupBy"
   aFormulaValue.Add CStr("'Requested By: " & sFacility & "'")
   aFormulaValue.Add CStr("'Includes " & txtEmp & " To " & txtEnd & "...'")
   aFormulaValue.Add optShowComments.Value
   If optDet.Value = 0 Then aFormulaValue.Add "1" Else aFormulaValue.Add "0"
   If optSortByHireDate.Value = 0 Then aFormulaValue.Add "2" Else aFormulaValue.Add "3"
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   'If optDet Then
      sCustomReport = GetCustomReport("admhu01")
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      cCRViewer.SetReportTitle = sCustomReport
   'Else
   '   sCustomReport = GetCustomReport("admhu02a")
   '   cCRViewer.SetReportFileName sCustomReport, sReportPath
   '   cCRViewer.SetReportTitle = sCustomReport
   'End If
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{EmplTable.PREMNUMBER} in " & Val(txtEmp) & " to " & Val(txtEnd) & " " _
          & "AND {EmplTable.PREMSTATUS} LIKE '" & txtSta & "*'"
   cCRViewer.ShowGroupTree False
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
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

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = txtEmp & txtEnd & Trim(str(optDet.Value)) & Trim(str(optShowComments.Value)) & LTrim(str(optSortByHireDate.Value))
   SaveSetting "Esi2000", "EsiAdmn", "hu02", sOptions
   
End Sub

Private Sub GetOptions()
   Dim RdoEmp As ADODB.Recordset
   Dim sOptions As String
   On Error Resume Next
   sSql = "SELECT MIN(PREMNUMBER) FROM EmplTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp)
   If bSqlRows Then
      If Not IsNull(RdoEmp.Fields(0)) Then
         lFirstEmp = RdoEmp.Fields(0)
      Else
         lFirstEmp = 0
      End If
   End If
   sSql = "SELECT MAX(PREMNUMBER) FROM EmplTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp)
   If bSqlRows Then
      If Not IsNull(RdoEmp.Fields(0)) Then
         lLastEmp = RdoEmp.Fields(0)
      Else
         lLastEmp = 0
      End If
   End If
   sOptions = GetSetting("Esi2000", "EsiAdmn", "hu02", sOptions)
   txtEmp = Left(sOptions, 6)
   txtEnd = Mid(sOptions, 7, 6)
   optDet.Value = Val(Mid(sOptions, 13, 1))
   If Len(sOptions) > 13 Then optShowComments.Value = Val(Mid(sOptions, 14, 1)) Else optShowComments.Value = 0
   If Len(sOptions) > 14 Then optSortByHireDate.Value = Val(Mid(sOptions, 15, 1)) Else optSortByHireDate.Value = 0
   If Len(txtEmp) = 0 Then txtEmp = Format(lFirstEmp, "000000")
   If Len(txtEnd) = 0 Then txtEnd = Format(lLastEmp, "000000")
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   GetEmployees
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   GetEmployees
   PrintReport
   
End Sub

Private Sub optShowComments_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
End Sub

Private Sub txtEmp_LostFocus()
   txtEmp = CheckLen(txtEmp, 20)
   If Len(txtEmp) = 0 Then
      txtEmp = Format(lFirstEmp, "000000")
   Else
      txtEmp = Format(Abs(Val(txtEmp)), "000000")
   End If
   
End Sub


Private Sub txtEnd_LostFocus()
   txtEnd = CheckLen(txtEnd, 6)
   If Len(txtEnd) = 0 Then
      txtEnd = Format(lLastEmp, "000000")
   Else
      txtEnd = Format(Abs(Val(txtEnd)), "000000")
   End If
   
End Sub

Private Sub GetEmployees()
   On Error GoTo DiaErr1
   If Val(txtEmp) < lFirstEmp Then txtEmp = Format(lFirstEmp, "000000")
   If Len(txtEmp) = 0 Then
      txtEmp = Format(lFirstEmp, "000000")
   Else
      txtEmp = Format(Abs(Val(txtEmp)), "000000")
   End If
   
   If Val(txtEnd) > lLastEmp Then txtEnd = Format(lLastEmp, "000000")
   If Len(txtEnd) = 0 Then
      txtEnd = Format(lLastEmp, "000000")
   Else
      txtEnd = Format(Abs(Val(txtEnd)), "000000")
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getemploye"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtSta_LostFocus()
   txtSta = CheckLen(txtSta, 1)
   
End Sub

Private Sub FillStatus()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PREMSTATUS FROM EmplTable ORDER BY PREMSTATUS"
   LoadComboBox txtSta, -1
   Exit Sub
   
DiaErr1:
   sProcName = "fillstatus"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
