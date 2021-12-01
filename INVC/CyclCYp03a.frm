VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CyclCYp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventory Variance Reconciliation"
   ClientHeight    =   2670
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2670
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkUnreconciled 
      Caption         =   "Include Unreconciled Cycle Counts"
      Height          =   255
      Left            =   300
      TabIndex        =   9
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CyclCYp03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtDsc 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1440
      Width           =   3375
   End
   Begin VB.ComboBox cmbCid 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "List Includes Cycle ID's Ronconciled"
      Top             =   1080
      Width           =   2115
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "CyclCYp03a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "CyclCYp03a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   3
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
      FormDesignHeight=   2670
      FormDesignWidth =   7260
   End
   Begin VB.Label lblCabc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      ToolTipText     =   "ABC Code Selected"
      Top             =   1080
      Width           =   405
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cycle Count ID"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "CyclCYp03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'1/17/04 new (Larry's reports)
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetCycleCount() As Byte
   Dim RdoCid As ADODB.Recordset
   On Error GoTo DiaErr1

   lblCabc = ""
   txtDsc = ""
   If Trim(cmbCid) = "" Then
      Exit Function
   End If


'   sSql = "SELECT CCREF,CCDESC,CCABCCODE FROM CchdTable WHERE CCREF='" & Trim(cmbCid) & "' "
'          & "AND CCUPDATED=1"
   sSql = "SELECT CCREF,CCDESC,CCABCCODE FROM CchdTable WHERE CCREF='" & Trim(cmbCid) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCid, ES_FORWARD)
   If bSqlRows Then
      With RdoCid
         lblCabc = "" & Trim(!CCABCCODE)
         txtDsc = "" & Trim(!CCDESC)
         GetCycleCount = 1
         ClearResultSet RdoCid
      End With
   Else
      GetCycleCount = 0
'      MsgBox "That Count ID Wasn't Found, Is Locked, Or Is Not Saved.", _
'         vbInformation, Caption
      MsgBox "That Count ID Wasn't Found.", _
         vbInformation, Caption
   End If
   Set RdoCid = Nothing
   Exit Function
   
DiaErr1:
   GetCycleCount = 0
   sProcName = "getcycleco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub FillCombo()
   On Error GoTo DiaErr1
'   sSql = "SELECT CCCOUNTLOCKED,CCREF FROM CchdTable WHERE " _
'          & "CCCOUNTLOCKED=1 AND CCUPDATED=1 ORDER BY CCREF"
   sSql = "SELECT CCCOUNTLOCKED,CCREF" & vbCrLf _
      & "FROM CchdTable" & vbCrLf _
      & "WHERE CCCOUNTLOCKED=1" & vbCrLf
   If Me.chkUnreconciled = vbUnchecked Then
      sSql = sSql & "AND CCUPDATED=1" & vbCrLf
   End If
   sSql = sSql & "ORDER BY CCREF"
   
   LoadComboBox cmbCid
   If cmbCid.ListCount > 0 Then
      If Trim(cmbCid) = "" Then cmbCid = cmbCid.List(0)
      'GetCycleCount
'   Else
'      MsgBox "There Are No Reconciled Counts Recorded.", _
'         vbInformation, Caption
'      Unload Me
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub chkUnreconciled_Click()
   FillCombo
End Sub

Private Sub cmbCid_Click()
   GetCycleCount
   
End Sub

Private Sub cmbCid_LostFocus()
   GetCycleCount
   
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
   Set CyclCYp03a = Nothing
   
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
   aFormulaName.Add "RequestBy"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   sCustomReport = GetCustomReport("invab03")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
   sSql = "{CchdTable.CCREF} = '" & Trim(cmbCid) & "' "
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



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDsc.BackColor = BackColor
   
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
