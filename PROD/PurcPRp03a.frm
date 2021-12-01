VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendor Directory"
   ClientHeight    =   2610
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6900
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRp03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      Text            =   "ALL"
      Top             =   960
      Width           =   1555
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "PurcPRp03a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PurcPRp03a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.TextBox txtTyp 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Tag             =   "3"
      Text            =   "ALL"
      Top             =   1680
      Width           =   375
   End
   Begin VB.CheckBox optAdr 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5760
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   2160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2610
      FormDesignWidth =   6900
   End
   Begin VB.Label lblVEName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2280
      TabIndex        =   13
      Top             =   1320
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Addresses"
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   1812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   252
      Index           =   3
      Left            =   4320
      TabIndex        =   9
      Top             =   1680
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Vendor Type(s)"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   252
      Index           =   1
      Left            =   4320
      TabIndex        =   7
      Top             =   960
      Width           =   1572
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Vendor(s)"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "PurcPRp03a"
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
   sOptions = GetSetting("Esi2000", "EsiProd", "pr03", sOptions)
   If Len(sOptions) > 0 Then
      optAdr.Value = Val(Left(sOptions, 1))
      txtTyp = Mid(sOptions, 2, 2)
   End If
   If txtTyp = "" Then txtTyp = "ALL"
   
End Sub



Private Sub SaveOptions()
   Dim sOptions As String
   Dim sType As String * 2
   If txtTyp = "ALL" Then txtTyp = ""
   sType = txtTyp
   'Save by Menu Option
   sOptions = RTrim(optAdr.Value) _
              & sType
   SaveSetting "Esi2000", "EsiProd", "pr03", Trim(sOptions)
   
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
   MDISect.lblBotPanel = Caption
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
   Set PurcPRp03a = Nothing
   
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
    sCustomReport = GetCustomReport("prdpr03")
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = "prdpr03"

    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowAddress"
    
    aFormulaValue.Add CStr("'" & cmbVnd & "... " _
                        & "And Types " & txtTyp & "...'")
    aFormulaValue.Add CStr("'Requested By: " & sInitials & "'")
    aFormulaValue.Add CStr(optAdr)
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{VndrTable.VEREF} LIKE '" & sVendRef & "*' " _
          & "AND {VndrTable.VETYPE} LIKE '" & sType & "*' "
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
   If txtTyp <> "ALL" Then txtTyp = CheckLen(txtTyp, 2)
   
End Sub
