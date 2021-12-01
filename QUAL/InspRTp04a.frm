VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InspRTp04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Corrective Action"
   ClientHeight    =   2490
   ClientLeft      =   1770
   ClientTop       =   1140
   ClientWidth     =   7155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2490
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InspRTp04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optPrn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   560
         Picture         =   "InspRTp04a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         Picture         =   "InspRTp04a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2490
      FormDesignWidth =   7155
   End
   Begin VB.ComboBox cmbTag 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "(Blank For All)"
      Top             =   1035
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   4
      Left            =   5520
      TabIndex        =   11
      Top             =   1440
      Width           =   1452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   5520
      TabIndex        =   10
      Top             =   1080
      Width           =   1452
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   9
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspection Rpt  Dates"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Report Types"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1035
      Width           =   2175
   End
End
Attribute VB_Name = "InspRTp04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/1/05 Changed date handling
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbTag_GotFocus()
   Dim b As Byte
   Dim iList As Integer
   If Len(Trim(cmbTag)) = 0 Then cmbTag = "ALL"
   For iList = 0 To cmbTag.ListCount - 1
      If cmbTag = cmbTag.List(iList) Then b = 1
   Next
   If b = 0 Then
      Beep
      cmbTag = cmbTag.List(0)
   End If
   
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

Private Sub FillCombo()
   cmbTag = "ALL"
   AddComboStr cmbTag.hwnd, "ALL"
   AddComboStr cmbTag.hwnd, "Customer"
   AddComboStr cmbTag.hwnd, "Internal"
   AddComboStr cmbTag.hwnd, "MRB"
   AddComboStr cmbTag.hwnd, "Vendor"
   
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillCombo
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
   Set InspRTp04a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sBegDate As String
   Dim sEnddate As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
   
   If Not IsDate(txtBeg) Then
      sBegDate = "1995,01,01"
   Else
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEnddate = "2024,12,31"
   Else
      sEnddate = Format(txtEnd, "yyyy,mm,dd")
   End If
   
   On Error GoTo DiaErr1
   sCustomReport = GetCustomReport("quarj04")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbTag) & " Tags'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
   If cmbTag <> "ALL" Then
      sSql = "{RjhdTable.REJTYPE} = '" & Left(cmbTag, 1) & "' AND "
   Else
      sSql = ""
   End If
   sSql = sSql & "{RjhdTable.REJDATE} In Date(" & sBegDate & ") " _
          & "To Date(" & sEnddate & ")"
   sSql = sSql & " and {RjitTable.RITACT} = 0"
   cCRViewer.SetReportSelectionFormula (sSql)
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport1()
   Dim sBegDate As String
   Dim sEnddate As String
   MouseCursor 13
   
   If Not IsDate(txtBeg) Then
      sBegDate = "1995,01,01"
   Else
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEnddate = "2024,12,31"
   Else
      sEnddate = Format(txtEnd, "yyyy,mm,dd")
   End If
   
   'SetMdiReportsize MdiSect
   On Error GoTo DiaErr1
   sCustomReport = GetCustomReport("quarj04")
   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.Crw.Formulas(1) = "Includes='" & cmbTag & " Tags'"
   MdiSect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   If cmbTag <> "ALL" Then
      sSql = "{RjhdTable.REJTYPE} = '" & Left(cmbTag, 1) & "' AND "
   Else
      sSql = ""
   End If
   sSql = sSql & "{RjhdTable.REJDATE} In Date(" & sBegDate & ") " _
          & "To Date(" & sEnddate & ")"
   MdiSect.Crw.SelectionFormula = sSql
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optDis_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDate(txtBeg)
   
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDate(txtEnd)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 3)
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(txtBeg) & Trim(txtEnd)
   SaveSetting "Esi2000", "EsiQual", "rj04", Trim(sOptions)
   
End Sub
