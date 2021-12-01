VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CapaCPp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work Center Load"
   ClientHeight    =   3630
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3630
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbGroupBy 
      Height          =   315
      ItemData        =   "CapaCPp03a.frx":0000
      Left            =   2100
      List            =   "CapaCPp03a.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ComboBox cboShop 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2100
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select From List"
      Top             =   1020
      Width           =   1815
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CapaCPp03a.frx":002C
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cboWorkCenter 
      Height          =   315
      Left            =   2100
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Enter New (12 Char) Or Select From List"
      Top             =   1380
      Width           =   1815
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   4260
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1740
      Width           =   1250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2100
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1740
      Width           =   1250
   End
   Begin VB.CheckBox optGrp 
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2100
      TabIndex        =   5
      Top             =   3240
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optCmt 
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2100
      TabIndex        =   4
      Top             =   2880
      Width           =   735
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "CapaCPp03a.frx":07DA
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
         Picture         =   "CapaCPp03a.frx":0964
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
      Left            =   5940
      Top             =   2040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3630
      FormDesignWidth =   7155
   End
   Begin VB.Label Label1 
      Caption         =   "Group Report by"
      Height          =   255
      Left            =   300
      TabIndex        =   23
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   285
      Index           =   11
      Left            =   300
      TabIndex        =   21
      Top             =   1020
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operations Scheduled Date From"
      Height          =   465
      Index           =   9
      Left            =   300
      TabIndex        =   19
      Top             =   1740
      Width           =   1815
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   8
      Left            =   3180
      TabIndex        =   18
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   4
      Left            =   5655
      TabIndex        =   17
      Top             =   1740
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   360
      Picture         =   "CapaCPp03a.frx":0AE2
      ToolTipText     =   "Chart Results"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Chart"
      Height          =   285
      Index           =   7
      Left            =   300
      TabIndex        =   16
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Comments"
      Height          =   285
      Index           =   6
      Left            =   300
      TabIndex        =   15
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   5
      Left            =   300
      TabIndex        =   14
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Includes Load Chart By Total Hours"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Width           =   3000
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   2
      Left            =   5655
      TabIndex        =   12
      Top             =   1380
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1890
      TabIndex        =   10
      Top             =   2100
      Width           =   105
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center(s)"
      Height          =   285
      Index           =   0
      Left            =   300
      TabIndex        =   9
      Top             =   1380
      Width           =   1695
   End
End
Attribute VB_Name = "CapaCPp03a"
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

'Private Sub FillCombo()
'   On Error GoTo DiaErr1
'   sSql = "Qry_FillWorkCentersAll"
'   LoadComboBox cboWorkCenter
'   If cboWorkCenter.ListCount > 0 Then cboWorkCenter = cboWorkCenter.List(0)
'   Exit Sub
'
'DiaErr1:
'   sProcName = "fillcombo"
'   CurrError.Number = Err
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub

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
   sOptions = GetSetting("Esi2000", "EsiProd", "ca06", sOptions)
   If Len(sOptions) > 0 Then
      cboWorkCenter = Trim(Left(sOptions, 12))
   Else
      cboWorkCenter = "ALL"
   End If
   OptCmt.Value = Val(Mid(sOptions, 13, 1))
   optGrp.Value = Val(Mid(sOptions, 14, 1))
   If Len(sOptions) < 15 Then cmbGroupBy.ListIndex = 1 Else cmbGroupBy.ListIndex = Val(Mid(sOptions, 15, 1))
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sWcn As String * 12
   'Save by Menu Option
   sWcn = cboWorkCenter
   sOptions = sWcn & Trim(str(OptCmt.Value)) _
              & Trim(str(optGrp.Value)) & Trim(str(cmbGroupBy.ListIndex))
   SaveSetting "Esi2000", "EsiProd", "ca06", Trim(sOptions)
   
End Sub

Private Sub cboShop_Click()
   FillWorkCenters
End Sub

Private Sub cboShop_LostFocus()
   FillWorkCenters
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
   If bOnLoad <> 0 Then
      FillShops
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   cboWorkCenter = ""
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
   Set CapaCPp03a = Nothing
   
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
    
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "ShowComments"
    aFormulaName.Add "ShowGroup"
    
    aFormulaName.Add "GroupBy"
    

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "...'")
    aFormulaValue.Add CStr("'" & CStr(cboWorkCenter & "...  From " & txtBeg _
                        & " To " & txtDte) & "...'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
    aFormulaValue.Add OptCmt.Value
    aFormulaValue.Add optGrp.Value
   aFormulaValue.Add CStr("'" & CStr(LTrim(str(cmbGroupBy.ListIndex))) & "'")
   
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("prdca06")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue


'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'   If optCmt.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(0) = "DETAIL.0.0;F;;;"
'      MDISect.Crw.SectionFormat(1) = "DETAIL.0.1;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "DETAIL.0.0;T;;;"
'      MDISect.Crw.SectionFormat(1) = "DETAIL.0.1.F;;;"
'   End If
'   If optGrp.value = vbUnchecked Then
'      MDISect.Crw.SectionFormat(2) = "REPORTFTR.0.1;F;;;"
'   Else
'      MDISect.Crw.SectionFormat(2) = "REPORTFTR.0.1;T;;;"
'   End If
   sSql = "({RunsTable.RUNSTATUS} = 'SC' OR {RunsTable.RUNSTATUS} = 'RL' OR " _
          & "{RunsTable.RUNSTATUS} like 'P*') " & vbCrLf _
          & "AND ({RnopTable.OPSCHEDDATE} In Date(" & sBDate & ") " _
          & " To Date(" & sEDate & ") " & vbCrLf _
          & "AND {RnopTable.OPCENTER} LIKE '" & sCenter & "*'" & vbCrLf _
          & "AND {RnopTable.OPSHOP} = '" & Trim(Me.cboShop) & "'" & ")" _
          & "AND {RnopTable.OPCOMPLETE} = 0"
          
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

Private Sub Image1_Click()
   If optGrp.Value = vbChecked Then
      optGrp.Value = vbUnchecked
   Else
      optGrp.Value = vbChecked
   End If
   
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

Private Sub optGrp_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub optGrp_KeyPress(KeyAscii As Integer)
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



Private Sub cboWorkCenter_KeyPress(KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub cboWorkCenter_LostFocus()
   cboWorkCenter = CheckLen(cboWorkCenter, 12)
   If Len(cboWorkCenter) = 0 Then cboWorkCenter = "ALL"
   
End Sub

Private Sub FillShops()
   Dim wc As New ClassWorkCenter
   wc.PopulateShopCombo cboShop, cboWorkCenter
End Sub

Private Sub FillWorkCenters()
   Dim wc As New ClassWorkCenter
   wc.PoulateWorkCenterCombo cboShop, cboWorkCenter
End Sub


