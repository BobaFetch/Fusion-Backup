VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaPhu01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employees By Name"
   ClientHeight    =   3045
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3045
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optSortByHireDate 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox optShowComments 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   14
      Top             =   2280
      Width           =   975
   End
   Begin VB.ComboBox txtSta 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Includes Codes In Use"
      Top             =   1440
      Width           =   615
   End
   Begin VB.ComboBox txtEmp 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   8
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
         Picture         =   "diaPhu01.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaPhu01.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   4
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
      PictureUp       =   "diaPhu01.frx":0308
      PictureDn       =   "diaPhu01.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3045
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort by Hire Date"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Comments"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   12
      Top             =   1560
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Detail"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Last Name Or Blank For All)"
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   9
      Top             =   1080
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Leading Character(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1665
   End
End
Attribute VB_Name = "diaPhu01"
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
   If bOnLoad = 1 Then FillCombo
   bOnLoad = 0
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
   Set diaPhu01 = Nothing
   
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
   aFormulaValue.Add CStr("'Includes " & txtEmp & "...'")
   aFormulaValue.Add optShowComments.Value
   If optDet.Value = 0 Then aFormulaValue.Add "1" Else aFormulaValue.Add "0"
   If optSortByHireDate.Value = 0 Then aFormulaValue.Add "1" Else aFormulaValue.Add "3"
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
'   If optDet Then
      sCustomReport = GetCustomReport("admhu01")
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      cCRViewer.SetReportTitle = sCustomReport
'   Else
'      sCustomReport = GetCustomReport("admhu01a")
'      cCRViewer.SetReportFileName sCustomReport, sReportPath
'      cCRViewer.SetReportTitle = sCustomReport
'   End If
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{EmplTable.PREMLSTNAME} LIKE '" & sEmployee & "*' " _
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
   SaveSetting "Esi2000", "EsiAdmn", "hu01", Trim(str(optDet.Value)) & Trim(str(optShowComments.Value)) & Trim(str(optSortByHireDate.Value))
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   sOptions = GetSetting("Esi2000", "EsiAdmn", "hu01", sOptions)
   optDet.Value = Val(Left(sOptions, 1))
   If Len(sOptions) > 1 Then optShowComments.Value = Val(Mid(sOptions, 2, 1)) Else optShowComments.Value = 0
   If Len(sOptions) > 2 Then optSortByHireDate.Value = Val(Mid(sOptions, 3, 1)) Else optSortByHireDate.Value = 0
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optDis_Click()
   If Trim(txtEmp) = "" Then txtEmp = "ALL"
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   If Trim(txtEmp) = "" Then txtEmp = "ALL"
   PrintReport
   
End Sub



Private Sub optShowComments_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
End Sub

Private Sub txtEmp_LostFocus()
   txtEmp = CheckLen(txtEmp, 20)
   If Len(txtEmp) = 0 Then txtEmp = "ALL"
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PREMLSTNAME FROM EmplTable "
   LoadComboBox txtEmp, -1
   txtEmp = "ALL"
   sSql = "SELECT DISTINCT PREMSTATUS FROM EmplTable ORDER BY PREMSTATUS"
   LoadComboBox txtSta, -1
   Exit Sub
   
DiaErr1:
   sProcName = "fillemploy"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtSta_LostFocus()
   txtSta = CheckLen(txtSta, 1)
   
End Sub
