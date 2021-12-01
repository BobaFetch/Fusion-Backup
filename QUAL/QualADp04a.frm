VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form QualADp04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendor Approval List"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox cbShowAddress 
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   2640
      Width           =   495
   End
   Begin VB.CheckBox cbShowInternalCmt 
      Height          =   195
      Left            =   1920
      TabIndex        =   12
      Top             =   3360
      Width           =   495
   End
   Begin VB.CheckBox cbShowStatusCodes 
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   3000
      Width           =   495
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   1800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3690
      FormDesignWidth =   7170
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "QualADp04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   6600
      Picture         =   "QualADp04a.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print The Report"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   6000
      Picture         =   "QualADp04a.frx":0938
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Display The Report"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.TextBox txtTyp 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Tag             =   "3"
      Text            =   "ALL"
      Top             =   1800
      Width           =   375
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      Text            =   "ALL"
      Top             =   960
      Width           =   1555
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   16
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Internal Comments"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   14
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Codes"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Vendor(s)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Vendor Type(s)"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblVEName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   3135
   End
End
Attribute VB_Name = "QualADp04a"
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

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillVendorsNone"
   LoadComboBox cmbVnd
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "qualadp04a", sOptions)
   txtTyp = Trim(Left(sOptions, 2))
   cbShowStatusCodes.value = Mid(sOptions, 3, 1)
   cbShowInternalCmt.value = Mid(sOptions, 4, 1)
   cbShowAddress.value = Mid(sOptions, 5, 1)
    
End Sub



Private Sub SaveOptions()
   Dim sOptions As String
   Dim sType As String * 2
   
   If txtTyp = "ALL" Then txtTyp = ""
   sType = Left(txtTyp & Space(2), 2)
   'Save by Menu Option
   sOptions = sType & cbShowStatusCodes.value & cbShowInternalCmt.value & cbShowAddress.value
   
   SaveSetting "Esi2000", "EsiProd", "qualadp04a", sOptions
End Sub

Private Sub cmbVnd_Click()
   GetThisVendor
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   If Len(cmbVnd) = 0 Then cmbVnd = "ALL"
   GetThisVendor
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
   Set QualADp04a = Nothing
End Sub

Private Sub PrintReport()
   Dim sVendRef As String
   Dim sType As String
   
   MouseCursor 13
   If Len(txtTyp) = 0 Then txtTyp = "ALL"
   If cmbVnd <> "ALL" Then
      sVendRef = Compress(cmbVnd)
   Else
      sVendRef = ""
   End If
   If txtTyp <> "ALL" Then
      sType = txtTyp
      sVendRef = Compress(sVendRef)
   Else
      sType = ""
   End If
   
   On Error GoTo Ppr03
   
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    
    cCRViewer.ShowGroupTree False
    sCustomReport = GetCustomReport("qualad04")
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = "qualad04"

    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowAddress"
    aFormulaName.Add "ShowInternalCmt"
    aFormulaName.Add "ShowStatusCode"
    
    
    aFormulaValue.Add CStr("'" & cmbVnd & "... " _
                        & "And Types " & txtTyp & "...'")
    aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
    aFormulaValue.Add cbShowAddress.value
    aFormulaValue.Add cbShowInternalCmt.value
    aFormulaValue.Add cbShowStatusCodes.value
    
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{VndrTable.VEREF} LIKE '" & sVendRef & "*' " _
          & "AND {VndrTable.VETYPE} LIKE '" & sType & "*' AND {StCmtTable.STATCODE_TYPE_REF}='VE' "
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub
   
Ppr03:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Ppr03a
Ppr03a:
   DoModuleErrors Me
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub txtTyp_LostFocus()
    If Len(Trim(txtTyp)) = 0 Then txtTyp = "ALL"
    If txtTyp <> "ALL" Then txtTyp = CheckLen(txtTyp, 2)
End Sub


