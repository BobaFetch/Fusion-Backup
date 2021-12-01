VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CapaCPp04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Late Manufacturing Orders"
   ClientHeight    =   2535
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2535
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optSO 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   252
      Left            =   2640
      TabIndex        =   10
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPp04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Tag             =   "4"
      Top             =   1080
      Width           =   1250
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   252
      Left            =   2640
      TabIndex        =   1
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "CapaCPp04a.frx":07AE
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
         Picture         =   "CapaCPp04a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   2040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2535
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include SO"
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   11
      Top             =   1920
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Operation Comments"
      Height          =   288
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   2352
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(On Or Before)"
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   7
      Top             =   1080
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturing Orders Due "
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   2265
   End
End
Attribute VB_Name = "CapaCPp04a"
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
Private txtKeyPress(1) As New EsiKeyBd
Private txtGotFocus(1) As New EsiKeyBd
Private txtKeyDown(1) As New EsiKeyBd




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
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   GetOptions
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set CapaCPp04a = Nothing
   
End Sub




Private Sub PrintReport()
    Dim sDate As String
    MouseCursor 13
    
    On Error GoTo DiaErr1
    sDate = Right(txtDte, 2)
    If Val(sDate) > 80 Then
       sDate = "19" & sDate
    Else
       sDate = "20" & sDate
    End If
    sDate = sDate & "," & Left(txtDte, 2) & "," & Mid(txtDte, 4, 2)
   
    Dim sCustomReport As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    Dim strIncludes As String
    Dim strRequestBy As String
    
       sCustomReport = GetCustomReport("prdca08")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.ShowGroupTree False
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowPartDesc"
    aFormulaName.Add "ShowSO"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")

    strIncludes = "'Manufacturing Orders On Or Before " & txtDte & "...'"
    aFormulaValue.Add CStr(strIncludes)
    strRequestBy = "'Requested By: " & sInitials & "'"
    aFormulaValue.Add CStr(strRequestBy)
   
    aFormulaValue.Add CStr(OptCmt)
    aFormulaValue.Add CStr(optSO)
   
     sSql = "{RunsTable.RUNSCHED}<= Date(" & sDate & ") "
    
    ' Set Formula values
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
    cCRViewer.SetReportDistinctRecords True
    ' set the report Selection
    cCRViewer.SetReportSelectionFormula (sSql)
    cCRViewer.CRViewerSize Me
    
    ' Set report parameter
    cCRViewer.SetDbTableConnection


    cCRViewer.OpenCrystalReportObject Me, aFormulaName

    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aRptParaType
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
   On Error Resume Next
   Set txtGotFocus(0).esCmbGotfocus = txtDte
   Set txtKeyPress(0).esCmbKeyDate = txtDte
   
End Sub

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub



Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   SaveSetting "Esi2000", "EsiProd", "ca08", Trim(str(OptCmt.Value))
   SaveSetting "Esi2000", "EsiProd", "ca081", Trim(str(optSO.Value))
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim sSO As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "ca08", sOptions)
   sSO = GetSetting("Esi2000", "EsiProd", "ca081", sSO)
   OptCmt.Value = Val(sOptions)
   optSO.Value = Val(sSO)
   
End Sub

Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDateEx(txtDte)
   
End Sub
