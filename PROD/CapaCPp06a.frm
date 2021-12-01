VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CapaCPp06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shop Load By Work Center"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPp06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1440
      Width           =   1250
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      Top             =   2160
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optGrp 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   2400
      Width           =   735
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1440
      Width           =   1250
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
         Picture         =   "CapaCPp06a.frx":07AE
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
         Picture         =   "CapaCPp06a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbShp 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select From List"
      Top             =   1080
      Width           =   1815
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
      FormDesignHeight=   3060
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   3
      Left            =   5280
      TabIndex        =   16
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   288
      Index           =   1
      Left            =   3000
      TabIndex        =   15
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   5280
      TabIndex        =   14
      Top             =   1080
      Width           =   1548
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   360
      Picture         =   "CapaCPp06a.frx":0AB6
      ToolTipText     =   "Chart Results"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Chart"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Sheduled Date From"
      Height          =   645
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Shop"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   2265
   End
End
Attribute VB_Name = "CapaCPp06a"
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

Private Sub cmbShp_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbShp = CheckLen(cmbShp, 12)
   
   If Len(cmbShp) = 0 Then
        cmbShp = "ALL"
        Exit Sub
    End If
   
   
   For iList = 0 To cmbShp.ListCount - 1
      If cmbShp = cmbShp.List(iList) Then b = 1
   Next
   If b = 0 Then
      Beep
      cmbShp = cmbShp.List(0)
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
   On Error GoTo DiaErr1
   AddComboStr cmbShp.hwnd, "ALL"
   sSql = "Qry_FillShops "
   LoadComboBox cmbShp
   cmbShp = cmbShp.List(0)
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
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = "01/01/" & Right(txtDte, 4)
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
   Set CapaCPp06a = Nothing
   
End Sub
Private Sub PrintReport()
   Dim sBDate As String
   Dim sEDate As String
   Dim sShop As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   
   MouseCursor 13
   
   If Trim(txtBeg) = "" Then txtBeg = "ALL"
   If Trim(txtDte) = "" Then txtDte = "ALL"
   
   If Not IsDate(txtBeg) Then
      sBDate = "1995,01,01"
   Else
      sBDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtDte) Then
      sEDate = "2024,12,31"
   Else
      sEDate = Format(txtDte, "yyyy,mm,dd")
   End If
   On Error GoTo DiaErr1
   If cmbShp = "ALL" Then
      sShop = ""
   Else
      sShop = Compress(cmbShp)
   End If
'   SetMdiReportsize MDISect
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "Includes='" & cmbShp & " " & "From " & txtBeg _
'                        & " To " & txtDte & "...'"
'   MDISect.Crw.Formulas(2) = "RequestBy = 'Requested By: " & sInitials & "'"

    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowComments"
    aFormulaName.Add "ShowGroup"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Includes" & CStr(cmbShp & " " & "From " & txtBeg _
                        & " To " & txtDte) & "...'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add OptCmt.Value
    aFormulaValue.Add optGrp.Value
    
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("prdca13")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   If optCmt.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "DETAIL.0.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "DETAIL.0.0;T;;;"
'   End If
'   If optGrp.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(1) = "REPORTFTR.0.1;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(1) = "REPORTFTR.0.1;T;;;"
'   End If

   sSql = "({RunsTable.RUNSTATUS} = 'SC' OR {RunsTable.RUNSTATUS} = 'RL' OR " _
          & "{RunsTable.RUNSTATUS} like 'P*') "
   If cmbShp <> "ALL" Then
      sSql = sSql & "AND {ShopTable.SHPREF} = '" & sShop & "' "
   End If
   sSql = sSql & "AND {RnopTable.OPSCHEDDATE}In Date(" & sBDate & ") " _
          & "To Date(" & sEDate & ")" & " AND {RnopTable.OPCOMPLETE} = 0"
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue


'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
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
   Dim sOptions As String
   'Save by Menu Option
   sOptions = Trim(str(OptCmt.Value)) _
              & Trim(str(optGrp.Value))
   SaveSetting "Esi2000", "EsiProd", "ca13", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "ca13", sOptions)
   If Len(sOptions) Then
      OptCmt.Value = Val(Mid(sOptions, 1, 1))
      optGrp.Value = Val(Mid(sOptions, 2, 1))
   End If
   
End Sub

Private Sub Image1_Click()
   If optGrp.Value = vbChecked Then
      optGrp.Value = vbUnchecked
   Else
      optGrp.Value = vbChecked
   End If
   
End Sub

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optGrp_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Trim(txtBeg) = "" Then
      txtBeg = "ALL"
   Else
      txtBeg = CheckDateEx(txtBeg)
   End If
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtDte_LostFocus()
   If Trim(txtDte) = "" Then
      txtDte = "ALL"
   Else
      txtDte = CheckDateEx(txtDte)
   End If
   
End Sub
