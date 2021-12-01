VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form AdmnADp08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User List"
   ClientHeight    =   1470
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
   ScaleHeight     =   1470
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkInactive 
      Caption         =   "Include Inactive Users"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   540
      Width           =   2415
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "AdmnADp08a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "AdmnADp08a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "AdmnADp08a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   2
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
      FormDesignHeight=   1470
      FormDesignWidth =   7260
   End
End
Attribute VB_Name = "AdmnADp08a"
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
   MDISect.lblBotPanel = Caption
   bOnLoad = 0
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub GetOptions()

End Sub

Private Sub SaveOptions()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   FormUnload
    
   'We really need to clear out the table now to prevent a security risk with having all
   'the information stored in a SQL Table. That's the reason we put it in a binary file to
   'begin with but we forgot to clean up after ourselves. This is a temp table used only for
   'this report
   On Error Resume Next
   clsADOCon.ExecuteSql "delete from EsReportUserPermissions"
   clsADOCon.ExecuteSql "delete from EsReportUsers"

End Sub

Private Sub PrintReport()
   Dim sClass As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   MouseCursor 13
   On Error GoTo DiaErr1
   ' first generate table entries
   CopyUsersToTable
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("UserList")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport

    aFormulaName.Add "CompanyName"
    aFormulaValue.Add CStr("'" & sFacility & "'")
   If chkInactive = 1 Then
    aFormulaName.Add "Includes"
    aFormulaValue.Add CStr("'Includes Inactive Users'")
   Else
    aFormulaName.Add "Includes"
    aFormulaValue.Add CStr("'Active Users Only'")
   End If
   
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "IncludeInactive"
    aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
    aFormulaValue.Add chkInactive
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    sSql = ""
    
    If (chkInactive.Value = vbUnchecked) Then
      sSql = "{EsReportUsers.Active} = true"
    Else
      sSql = ""
    End If
    
    
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
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport1()
   Dim sClass As String
   MouseCursor 13
   
   On Error GoTo DiaErr1

   ' first generate table entries
   CopyUsersToTable

   'SetMdiReportsize MDISect
   sCustomReport = GetCustomReport("UserList")
   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   If chkInactive = 1 Then
    MDISect.Crw.Formulas(1) = "Includes='Includes Inactive Users'"
   Else
    MDISect.Crw.Formulas(1) = "Includes='Active Users Only'"
   End If
   
   MDISect.Crw.Formulas(2) = "RequestBy='Requested By: " & sInitials & "'"
   MDISect.Crw.Formulas(3) = "IncludeInactive=" & chkInactive
   sSql = ""

   MDISect.Crw.SelectionFormula = sSql
   'SetCrystalAction Me

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

Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub

