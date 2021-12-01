VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form LotsLTp05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventory Transfers (Report)"
   ClientHeight    =   2910
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2910
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbSONum 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   16
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Customer "
      Top             =   2400
      Width           =   2040
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotsLTp05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtSplit 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Customer "
      Top             =   720
      Width           =   1555
   End
   Begin VB.ComboBox txtBeg 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "LotsLTp05a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "LotsLTp05a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   2880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2910
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SO Number"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   315
      Index           =   6
      Left            =   5160
      TabIndex        =   14
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer Comments"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   1440
      Width           =   1600
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1800
      TabIndex        =   12
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   5
      Left            =   5160
      TabIndex        =   11
      Top             =   1935
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Lots From"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   9
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "LotsLTp05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'8/24/05 New
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetTransferCustomer()
   Dim rdoCst As ADODB.Recordset
   sSql = "SELECT CUNAME FROM CustTable WHERE CUREF='" _
          & Compress(cmbCst) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst, ES_FORWARD)
   If bSqlRows Then txtNme = "" & Trim(rdoCst!CUNAME) _
                             Else txtNme = ""
   GetSplitComments
   Set rdoCst = Nothing
End Sub

Private Function GetSOName()

   Dim strSOName As String
   If (cmbSONum <> "") Then
      Dim RdoSO As ADODB.Recordset
      sSql = "SELECT DISTINCT SONAME FROM sohdTable WHERE SONUMBER ='" _
             & Trim(cmbSONum) & "'"
      
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoSO, ES_FORWARD)
      If bSqlRows Then strSOName = "" & Trim(RdoSO!SONAME) _
                                Else strSOName = ""
   End If
   
   GetSOName = strSOName
   
   
End Function

Private Sub GetSONum()
   Dim RdoSO As ADODB.Recordset
   sSql = "SELECT DISTINCT SONUMBER FROM sohdTable WHERE SOCUST ='" _
          & Compress(cmbCst) & "' AND SONUMBER <> 0"
   
   cmbSONum.Clear
   LoadComboBox cmbSONum, -1
   
   If cmbSONum.ListCount > 0 Then
      cmbSONum = ""
   End If
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT LOTCUST FROM LohdTable " _
          & "WHERE LOTCUST<>''"
   LoadComboBox cmbCst, -1
   If cmbCst.ListCount > 0 Then
      cmbCst = cmbCst.List(0)
      GetTransferCustomer
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub cmbCst_Click()
   GetTransferCustomer
   
End Sub


Private Sub cmbCst_LostFocus()
   GetTransferCustomer
   GetSONum
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
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
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
   Set LotsLTp05a = Nothing
   
End Sub




Private Sub PrintReport()
   Dim sComment As String
   Dim sBegDate As String
   Dim sEndDate As String
   Dim strSOName As String
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   If IsDate(txtBeg) Then
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   Else
      sBegDate = "1995,01,01"
   End If
   If IsDate(txtEnd) Then
      sEndDate = Format(txtEnd, "yyyy,mm,dd")
   Else
      sEndDate = "2024,12,31"
   End If
   If txtSplit = "" Then txtSplit = "ALL"
   If txtSplit <> "ALL" Then sComment = txtSplit
   
   ' get the SOName
   strSOName = GetSOName
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("invlt07")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'" & strSOName & "'")
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   
   sSql = "{LohdTable.LOTCUST} LIKE '" & Compress(cmbCst) & "*' AND " _
          & "{LohdTable.LOTSPLITCOMMENT} LIKE '" & sComment & "*' AND " _
          & "{LohdTable.LOTADATE} in Date(" & sBegDate & ") " _
          & "to Date(" & sEndDate & ")"
          
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
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub






Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtEnd = Format(Now, "mm/dd/yy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 3)
   
End Sub

Private Sub SaveOptions()
   
End Sub

Private Sub GetOptions()
   
End Sub



Private Sub optDis_Click()
   If Trim(txtNme) <> "" Then PrintReport
   
End Sub


Private Sub optPrn_Click()
   If Trim(txtNme) <> "" Then PrintReport
   
End Sub




Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Trim(txtBeg) = "" Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDate(txtBeg)
   
End Sub


Private Sub txtEnd_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Trim(txtEnd) = "" Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDate(txtEnd)
   
End Sub




Private Sub txtSplit_LostFocus()
   txtSplit = CheckLen(txtSplit, 20)
   If txtSplit = "" Then txtSplit = "ALL"
   
End Sub



Private Sub GetSplitComments()
   txtSplit.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT LOTSPLITCOMMENT FROM LohdTable " _
          & "WHERE LOTSPLITCOMMENT<>'' AND LOTCUST='" & Compress(cmbCst) & "'" _
          & "ORDER BY LOTSPLITCOMMENT"
   LoadComboBox txtSplit, -1
   If txtSplit.ListCount > 0 Then txtSplit = txtSplit.List(0)
   Exit Sub
   Exit Sub
   
DiaErr1:
   sProcName = "getsplitcomme"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
