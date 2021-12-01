VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DockODp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parts On Dock (Delivered)"
   ClientHeight    =   3840
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
   ScaleHeight     =   3840
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z2 
      Height          =   495
      Left            =   2160
      TabIndex        =   23
      Top             =   2400
      Width           =   3255
      Begin VB.OptionButton optServPrt 
         Caption         =   "Service Part"
         Height          =   255
         Left            =   1800
         TabIndex        =   25
         ToolTipText     =   "Planned (User Entered Date)"
         Top             =   200
         Width           =   1335
      End
      Begin VB.OptionButton optNonServ 
         Caption         =   "Non-Service Part"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Actual (System Date)"
         Top             =   200
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DockODp02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   3120
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   288
      Left            =   2160
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Leading Characters Or Select From List (Contains Parts That Require OD Insp)"
      Top             =   1680
      Width           =   3120
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4200
      TabIndex        =   3
      Tag             =   "4"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Tag             =   "4"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "DockODp02a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "DockODp02a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   288
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select, Leading Char(s) Or Blank For ALL"
      Top             =   960
      Width           =   1555
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
      FormDesignHeight=   3840
      FormDesignWidth =   7260
   End
   Begin VB.Label lblVEName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2160
      TabIndex        =   22
      Top             =   1320
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   10
      Left            =   5520
      TabIndex        =   20
      Top             =   2040
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Best Results Are Obtained By Using The On Dock (Delivered) Function"
      ForeColor       =   &H00000000&
      Height          =   288
      Index           =   9
      Left            =   240
      TabIndex        =   19
      Top             =   600
      Width           =   5748
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descriptions"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   18
      Top             =   3120
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   17
      Top             =   3360
      Width           =   1905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   288
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   1680
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   4
      Left            =   5520
      TabIndex        =   15
      Top             =   1680
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   6
      Left            =   5520
      TabIndex        =   14
      Top             =   960
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivered From"
      Height          =   288
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   1812
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   288
      Index           =   1
      Left            =   3240
      TabIndex        =   12
      Top             =   2040
      Width           =   912
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Chart Results"
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor(s)"
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1428
   End
End
Attribute VB_Name = "DockODp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/1/05 Changed date handling
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbVnd_Click()
   GetThisVendor
   
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 12)
   If cmbVnd = "" Then cmbVnd = "ALL"
   GetThisVendor
   
End Sub

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

Private Sub FillCombo()
   On Error GoTo DiaErr1
   AddComboStr cmbVnd.hWnd, "ALL"
   sSql = "Qry_FillVendors"
   LoadComboBox cmbVnd
   If cmbVnd = "" Then cmbVnd = cmbVnd.List(0)
   GetThisVendor
   
   AddComboStr cmbPrt.hWnd, "ALL"
   sSql = "SELECT DISTINCT PIPART,PITYPE,PARTREF,PARTNUM " _
          & "FROM PoitTable,PartTable WHERE (PIPART=PARTREF AND " _
          & "PITYPE=14) ORDER BY PIPART "
   LoadComboBox cmbPrt, 2
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
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
   Set DockODp02a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sPartNumber As String
   Dim sVendRef As String
   Dim sBegDate As String
   Dim sEndDate As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim strPartTpe As String
   
   MouseCursor 13
   
   If Trim(cmbVnd) <> "ALL" Then sVendRef = cmbVnd
   If Trim(cmbPrt) <> "ALL" Then sPartNumber = Compress(cmbPrt)
   On Error Resume Next
   If Not IsDate(txtBeg) Then
      sBegDate = "1995,01,01"
   Else
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEndDate = "2024,12,31"
   Else
      sEndDate = Format(txtEnd, "yyyy,mm,dd")
   End If
   On Error GoTo DiaErr1
   sCustomReport = GetCustomReport("quaod02")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.ShowGroupTree False
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowDesc"
   aFormulaName.Add "ShowExDesc"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Vendor(s) " & CStr(cmbVnd _
                        & ", Part Number(s) " & cmbPrt & ", " _
                        & "Delivered From " & txtBeg & " Through " & txtEnd) & "...'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add optDsc.Value
   aFormulaValue.Add optExt.Value
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   If (optNonServ.Value = True) Then
      strPartTpe = " and {PartTable.palevel} In [1, 2,3,4,5,8]"
   Else
      strPartTpe = " and {PartTable.palevel} = 7 "
   End If
   
   sSql = "{PohdTable.POVENDOR} LIKE '" & sVendRef & "*' AND " _
          & "{PoitTable.PIPART} LIKE '" & sPartNumber & "*' AND " _
          & "{PoitTable.PIODDELDATE} In Date(" & sBegDate & ") " _
          & "To Date(" & sEndDate & ") "
   sSql = sSql & " and {PoitTable.PITYPE} = 14 "
   sSql = sSql & strPartTpe
   sSql = sSql & " and {PoitTable.PIODDELIVERED} = 1 AND {PoitTable.PIODDELQTY} > 0 "
   sSql = sSql & " AND {PoitTable.PIONDOCKQTYACC}+{PoitTable.PIONDOCKQTYREJ}=0 "
'   sSql = sSql & " and {PoitTable.PIONDOCKQTYACC}+{PoitTable.PIONDOCKQTYREJ}>0"
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
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 5)
   cmbPrt = "ALL"
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(str(optDsc.Value)) & Trim(str(optExt.Value))
   SaveSetting "Esi2000", "EsiQual", "od02", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiQual", "od02", Trim(sOptions))
   If Len(Trim(sOptions)) > 0 Then
      optDsc.Value = Val(Left(sOptions, 1))
      optExt.Value = Val(Right(sOptions, 1))
   Else
      optDsc.Value = vbChecked
      optExt.Value = vbChecked
   End If
   
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

Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub

Private Sub txtEnd_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub

Public Sub GetThisVendor()
   Dim RdoRpt As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT VEBNAME FROM VndrTable WHERE VEREF='" _
          & Compress(MdiSect.ActiveForm.cmbVnd) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      MdiSect.ActiveForm.lblVEName = "" & Trim(RdoRpt!VEBNAME)
      ClearResultSet RdoRpt
   Else
      MdiSect.ActiveForm.lblVEName = "*** A Range Of Vendors Selected ***"
   End If
   Set RdoRpt = Nothing
   Exit Sub
modErr1:
   
   On Error GoTo 0
End Sub



