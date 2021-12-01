VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESp04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estimates Not Completed"
   ClientHeight    =   2415
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2415
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESp04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
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
      Top             =   1680
      Width           =   1250
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1680
      Width           =   1250
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Customers With Estimates"
      Top             =   960
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "EstiESp04a.frx":07AE
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
         Picture         =   "EstiESp04a.frx":092C
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
      FormDesignHeight=   2415
      FormDesignWidth =   7125
   End
   Begin VB.Label lblCUName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2040
      TabIndex        =   12
      Top             =   1320
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   252
      Index           =   3
      Left            =   5640
      TabIndex        =   10
      Top             =   1680
      Width           =   1404
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimates Due From"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   1788
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   8
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Nickname"
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1788
   End
End
Attribute VB_Name = "EstiESp04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCst_Click()
   GetThisCustomer
   
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If cmbCst = "" Then cmbCst = "ALL"
   GetThisCustomer
   
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
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT BIDCUST,CUREF,CUNICKNAME " _
          & "FROM EstiTable,CustTable WHERE BIDCUST=CUREF " _
          & "ORDER BY CUREF"
   LoadComboBox cmbCst, 1
   cmbCst = "ALL"
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
      GetBidDates
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
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set EstiESp04a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sBDate As String
   Dim sEDate As String
   Dim sCust As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim sEnddate As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   If Not IsDate(txtBeg) Then
      sBDate = "1995,01,01"
   Else
      sBDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEDate = "2024,12,31"
   Else
      sEDate = Format(txtEnd, "yyyy,mm,dd")
   End If
   If Trim(cmbCst) <> "ALL" Then sCust = Compress(cmbCst)
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Customer(s) " & CStr(cmbCst & " From " _
                        & txtBeg & " Through " & txtEnd) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   
   sCustomReport = GetCustomReport("enges04")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{EstiTable.BIDCUST} LIKE '" & sCust & "*' " _
          & "AND ({EstiTable.BIDDUE} In Date(" & sBDate & ") " _
          & "To Date(" & sEDate & "))"
   sSql = sSql & " and {EstiTable.BIDCOMPLETE} = 0 and {EstiTable.BIDCANCELED} = 0"
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
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

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub

Private Sub GetBidDates()
   Dim RdoDte As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT MIN(BIDDUE) FROM EstiTable WHERE BIDCOMPLETE=0"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtBeg = Format(.Fields(0), "mm/dd/yyyy")
         Else
           ' txtBeg = Format(ES_SYSDATE, "mm/dd/yy")
            txtBeg = ""
         End If
         ClearResultSet RdoDte
      End With
   End If
   
   sSql = "SELECT Max(BIDDUE) FROM EstiTable WHERE BIDCOMPLETE=0"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtEnd = Format(.Fields(0), "mm/dd/yyyy")
         Else
            'txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
            txtEnd = ""
         End If
         ClearResultSet RdoDte
      End With
   End If
   
   Set RdoDte = Nothing
End Sub
