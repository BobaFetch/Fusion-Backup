VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHp05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Order Status"
   ClientHeight    =   3855
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
   ScaleHeight     =   3855
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHp05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Leading Character Search Or Select"
      Top             =   1080
      Width           =   3545
   End
   Begin VB.CheckBox optSta 
      Caption         =   "RL"
      Height          =   255
      Index           =   7
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   14
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox OptCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   2160
      TabIndex        =   20
      Top             =   3000
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   13
      Top             =   2520
      Width           =   735
   End
   Begin VB.CheckBox optLvl 
      Caption         =   "8"
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   12
      Top             =   1920
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optLvl 
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   11
      Top             =   1920
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optLvl 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   10
      Top             =   1920
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optLvl 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   9
      Top             =   1920
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CA"
      Height          =   255
      Index           =   6
      Left            =   6360
      TabIndex        =   8
      Top             =   1560
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CL"
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   7
      Top             =   1560
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "CO"
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   6
      Top             =   1560
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PC"
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   5
      Top             =   1560
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PP"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   4
      Top             =   1560
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "PL"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      Top             =   1560
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CheckBox optSta 
      Caption         =   "SC"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   1560
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   16
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ShopSHp05a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "ShopSHp05a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3855
      FormDesignWidth =   7260
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   7
      Left            =   5880
      TabIndex        =   27
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Comments"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   26
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   25
      Top             =   2520
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   24
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   23
      Tag             =   " "
      Top             =   2280
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Types"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Tag             =   " "
      Top             =   1920
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Status"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   18
      Tag             =   " "
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   1080
      Width           =   1545
   End
End
Attribute VB_Name = "ShopSHp05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/24/05 Corrected the report join
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd




Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If cmbPrt = "" Then cmbPrt = "ALL"
   
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
   MDISect.lblBotPanel = Caption
   If bOnLoad Then FillAllRuns cmbPrt
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
   Set CapaCPp06a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim iList As Integer
   Dim sPart As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection

   
   MouseCursor 13
   On Error GoTo DiaErr1
   If cmbPrt = "ALL" Then sPart = "" Else sPart = Compress(cmbPrt)
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & cmbPrt & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")


   sCustomReport = GetCustomReport("prdsh05")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport


   sSql = "{RunsTable.RUNREF} LIKE '" & sPart & "*' "
   If optSta(0) = vbUnchecked Then sSql = sSql & " AND {RunsTable.RUNSTATUS}<>'SC'"
   If optSta(1) = vbUnchecked Then sSql = sSql & " AND {RunsTable.RUNSTATUS}<>'PL'"
   If optSta(2) = vbUnchecked Then sSql = sSql & " AND {RunsTable.RUNSTATUS}<>'PP'"
   If optSta(3) = vbUnchecked Then sSql = sSql & " AND {RunsTable.RUNSTATUS}<>'PC'"
   If optSta(4) = vbUnchecked Then sSql = sSql & " AND {RunsTable.RUNSTATUS}<>'CO'"
   If optSta(5) = vbUnchecked Then sSql = sSql & " AND {RunsTable.RUNSTATUS}<>'CL'"
   If optSta(6) = vbUnchecked Then sSql = sSql & " AND {RunsTable.RUNSTATUS}<>'CA'"
   If optSta(7) = vbUnchecked Then sSql = sSql & " AND {RunsTable.RUNSTATUS}<>'RL'"
   For iList = 0 To 2
      If optLvl(iList) = vbUnchecked Then
         sSql = sSql & " AND {PartTable.PALEVEL}<>" & str(iList + 1) & ""
      End If
   Next
   If optLvl(iList) = vbUnchecked Then
      sSql = sSql & " AND {PartTable.PALEVEL}<>8"
   End If
       aFormulaName.Add "Desc"
   If optDsc.Value = vbUnchecked Then
   '   MDISect.Crw.Formulas(3) = "Desc=''"
      aFormulaValue.Add CStr("''")
   Else
   '   MDISect.Crw.Formulas(3) = "Desc='1'"
         aFormulaValue.Add CStr("'1'")
   End If
    aFormulaName.Add "ShowComments"
    aFormulaName.Add "ShowExtendedDescription"
    aFormulaValue.Add OptCmt.Value
    aFormulaValue.Add optExt.Value

'   If OptCmt.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "GROUPFTR.1.0;F;;;"
'      MDISect.Crw.SectionFormat(1) = "GROUPFTR.1.1;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "GROUPFTR.1.0;T;;;"
'      MDISect.Crw.SectionFormat(1) = "GROUPFTR.1.1;T;;;"
'   End If
'   If optExt.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(2) = "GROUPHDR.1.0;F;;;"
'      MDISect.Crw.SectionFormat(3) = "GROUPHDR.1.1;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(2) = "GROUPHDR.1.0;T;;;"
'      MDISect.Crw.SectionFormat(3) = "GROUPHDR.1.1;T;;;"
'   End If
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

  ' MDISect.Crw.SelectionFormula = sSql
  ' SetCrystalAction Me
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
   cmbPrt = "ALL"
   
End Sub

Private Sub SaveOptions()
   Dim iList As Integer
   Dim sOptions As String
   
   For iList = 0 To 5
      sOptions = sOptions & Trim(str(optSta(iList).Value))
   Next
   sOptions = sOptions & Trim(str(optSta(iList).Value))
   
   For iList = 0 To 2
      sOptions = sOptions & Trim(str(optLvl(iList).Value))
   Next
   sOptions = sOptions & Trim(str(optSta(iList).Value))
   
   sOptions = sOptions & Trim(str(optDsc.Value))
   sOptions = sOptions & Trim(str(optExt.Value))
   sOptions = sOptions & Trim(str(OptCmt.Value))
   SaveSetting "Esi2000", "EsiProd", "sh05", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh05", sOptions)
   If Len(Trim(sOptions)) Then
      For iList = 0 To 5
         optSta(iList).Value = Val(Mid(sOptions, iList + 1, 1))
      Next
      optSta(iList).Value = Val(Mid(sOptions, iList + 1, 1))
      
      For iList = 7 To 9
         optLvl(iList - 7).Value = Val(Mid(sOptions, iList + 1, 1))
      Next
      optLvl(iList - 7).Value = Val(Mid(sOptions, iList + 1, 1))
      
      optDsc.Value = Val(Mid(sOptions, 12, 1))
      optExt.Value = Val(Mid(sOptions, 13, 1))
      OptCmt.Value = Val(Mid(sOptions, 14, 1))
   End If
End Sub

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optLvl_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub optSta_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub
