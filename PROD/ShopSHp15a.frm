VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHp15a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work Center Load"
   ClientHeight    =   2775
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2775
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboWorkCenter 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "99"
      ToolTipText     =   "Select Work Center, Leading Characters, Or Blank"
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox cboShop 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Tag             =   "99"
      ToolTipText     =   "Select From List"
      Top             =   600
      Width           =   1815
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHp15a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   4200
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CheckBox optCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   5
      Top             =   2280
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "ShopSHp15a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ShopSHp15a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
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
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2775
      FormDesignWidth =   7155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   12
      Left            =   4320
      TabIndex        =   21
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   20
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Centers"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   19
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Select From List)"
      Height          =   285
      Index           =   7
      Left            =   4320
      TabIndex        =   18
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   288
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   1692
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operations Scheduled Date From"
      Height          =   648
      Index           =   9
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Width           =   1812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   288
      Index           =   8
      Left            =   3360
      TabIndex        =   14
      Top             =   1440
      Width           =   1692
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   4
      Left            =   5760
      TabIndex        =   13
      Top             =   1440
      Width           =   1692
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Comments"
      Height          =   288
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1692
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1830
      TabIndex        =   9
      Top             =   1680
      Width           =   105
   End
End
Attribute VB_Name = "ShopSHp15a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/16/06 New
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'Private Sub FillCombo()
'   On Error GoTo DiaErr1
'   sSql = "Qry_FillWorkCentersAll"
'   LoadComboBox txtWcn
'   If txtWcn.ListCount > 0 Then txtWcn = txtWcn.List(0)
'   Exit Sub
'
'DiaErr1:
'   sProcName = "fillcombo"
'   CurrError.Number = Err
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'
Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = "01/01/" & Right(txtDte, 4)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh15a", sOptions)
   If Len(sOptions) Then
      optDsc.Value = Val(Mid(sOptions, 1, 1))
      OptCmt.Value = Val(Mid(sOptions, 2, 1))
   End If
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(str(optDsc.Value)) _
              & Trim(str(OptCmt.Value))
   SaveSetting "Esi2000", "EsiProd", "sh15a", Trim(sOptions)
   
End Sub

Private Sub cboShop_Click()
   FillWorkCenters
End Sub

Private Sub cboShop_LostFocus()
   FillWorkCenters
End Sub

Private Sub cboWorkCenter_LostFocus()
   If cboWorkCenter = "" Then cboWorkCenter = "ALL"
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
   If bOnLoad Then
      FillShops
      'cboWorkCenter = "ALL"
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
   Set ShopSHp15a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sCenter As String
   Dim sBDate As String
   Dim sEDate As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   sCenter = Compress(cboWorkCenter)
   If sCenter = "ALL" Then sCenter = ""
   
   sEDate = Right(txtDte, 2)
   If Val(sEDate) > 80 Then
      sEDate = "19" & sEDate
   Else
      sEDate = "20" & sEDate
   End If
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
   MouseCursor 13
   On Error GoTo DiaErr1
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("prdsh15")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowComments"
   aFormulaName.Add "ShowDescription"
    
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Shop" & cboShop & " WC " & cboWorkCenter & " From " & txtBeg & " To " & txtDte & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add OptCmt.Value
   aFormulaValue.Add optDsc.Value
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   If optCmt.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "DETAIL.0.0;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "DETAIL.0.0.F;;;"
'   End If
'   If optDsc.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(1) = "GROUPHDR.1.1;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(1) = "GROUPHDR.1.1;T;;;"
'   End If
   
   sSql = "({RunsTable.RUNSTATUS} = 'SC' OR {RunsTable.RUNSTATUS} = 'RL' OR " & vbCrLf _
          & "{RunsTable.RUNSTATUS} like 'P*') " & vbCrLf _
          & "AND ({RnopTable.OPSCHEDDATE} In Date(" & sBDate & ") " _
          & "To Date(" & sEDate & ") " & vbCrLf _
          & "AND {RnopTable.OPSHOP} = '" & cboShop & "'" & vbCrLf _
          & "AND {RnopTable.OPCENTER} LIKE '" & sCenter & "*')" _
          & " AND {RnopTable.OPCOMPLETE} = 0"
   
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

Private Sub optCmt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   MouseCursor 13
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

Private Sub FillShops()
   Dim wc As New ClassWorkCenter
   wc.PopulateShopCombo cboShop, cboWorkCenter
End Sub

Private Sub FillWorkCenters()
   Dim wc As New ClassWorkCenter
   wc.PoulateWorkCenterCombo cboShop, cboWorkCenter
End Sub


