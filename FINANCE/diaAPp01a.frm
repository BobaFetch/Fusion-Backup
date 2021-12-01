VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendor Invoice (Report)"
   ClientHeight    =   4170
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4170
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optRem 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1875
      TabIndex        =   7
      Top             =   3360
      Width           =   735
   End
   Begin VB.CheckBox optDsb 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1875
      TabIndex        =   6
      Top             =   3120
      Width           =   735
   End
   Begin VB.CheckBox optAll 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1875
      TabIndex        =   5
      Top             =   2880
      Width           =   735
   End
   Begin VB.CheckBox optItm 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1875
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1875
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1875
      TabIndex        =   2
      Top             =   2160
      Width           =   735
   End
   Begin VB.ComboBox cmbInv 
      Height          =   315
      Left            =   1640
      TabIndex        =   1
      ToolTipText     =   "List Of Invoices Found"
      Top             =   1440
      Width           =   2775
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1640
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Vendors With Invoices"
      Top             =   600
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaAPp01a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaAPp01a.frx":041D
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   300
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   300
      _Version        =   65536
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      PictureDnChange =   0
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPp01a.frx":0868
      PictureDn       =   "diaAPp01a.frx":0CB8
      PictureDisabled =   "diaAPp01a.frx":0DFE
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   840
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4170
      FormDesignWidth =   7290
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   25
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPp01a.frx":124E
      PictureDn       =   "diaAPp01a.frx":169E
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices Found"
      Height          =   405
      Index           =   10
      Left            =   4680
      TabIndex        =   24
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label lblCnt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5400
      TabIndex        =   23
      Top             =   1440
      Width           =   645
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Remarks"
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   22
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disbursements"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   21
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Allocations"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   20
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Comments"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Desc"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Descriptions"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1640
      TabIndex        =   13
      Top             =   960
      Width           =   2775
   End
End
Attribute VB_Name = "diaAPp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'************************************************************************************
' diaAPp01a - Print Display a Vendor Invoice
'
' Notes:
'   Print/Display a vendor invoice
'
' Created: (cjs)
' Revisions:
'   08/02/01 (nth) Dumped the jet logic in favor of just using SQL Server tables.
'   11/06/02 (nth) Refreshed and revisited.
'   12/26/02 (nth) Added vendor invoice comments.
'   06/04/03 (nth) Added cur.currentvendor
'   10/22/03 (nth) Added customreport
'   08/16/04 (nth) Added printer to saveoptions and getoptions.
'
'***********************************************************************************

Dim RdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

Dim bExt As Byte
Dim bItm As Byte
Dim bOnLoad As Byte
Dim bGoodVendor As Boolean
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'***********************************************************************************

Private Sub cmbInv_LostFocus()
   Dim bByte As Byte
   Dim i As Integer
   cmbInv = CheckLen(cmbInv, 20)
   For i = 0 To cmbInv.ListCount - 1
      If cmbInv = cmbInv.List(i) Then
         bByte = True
      End If
   Next
   If Not bByte Then
      Beep
      If cmbInv.ListCount > 0 Then cmbInv = cmbInv.List(0)
   End If
End Sub

Private Sub cmbVnd_Click()
   bGoodVendor = FindVendor(Me)
   If bGoodVendor Then GetInvoices
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   If Len(cmbVnd) Then
      bGoodVendor = FindVendor(Me)
      If bGoodVendor Then GetInvoices
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Me.Caption
      MouseCursor 0
      cmdHlp = False
   End If
End Sub

Private Sub FillCombo()
   Dim RdoVed As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT VIVENDOR,VEREF,VENICKNAME " _
          & "FROM VihdTable,VndrTable WHERE VIVENDOR=VEREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVed)
   If bSqlRows Then
      With RdoVed
         cmbVnd = "" & Trim(!VENICKNAME)
         Do Until .EOF
            AddComboStr cmbVnd.hwnd, "" & Trim(!VENICKNAME)
            .MoveNext
         Loop
      End With
   End If
   Set RdoVed = Nothing
   If cmbVnd.ListCount > 0 Then
      cmbVnd = cUR.CurrentVendor
      bGoodVendor = FindVendor(Me)
      GetInvoices
   End If
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   sSql = "SELECT DISTINCT VINO,VIVENDOR FROM " _
          & "VihdTable WHERE VIVENDOR= ? "
   Set RdoQry = New ADODB.Command
   RdoQry.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 10
   
   RdoQry.parameters.Append AdoParameter1
   
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   GetOptions
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   If bGoodVendor Then
      cUR.CurrentVendor = cmbVnd
      SaveCurrentSelections
   End If
   FormUnload
   Set AdoParameter1 = Nothing
   Set RdoQry = Nothing
   Set diaAPp01a = Nothing
End Sub

Private Sub PrintReport()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   On Error GoTo DiaErr1
   MouseCursor 13
'   SetMdiReportsize MdiSect
'   MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
    aFormulaName.Add "CompanyName"
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   
   If optRem.Value = vbChecked Then
'      MdiSect.Crw.Formulas(1) = "ShowRemarks='1'"
      aFormulaName.Add "ShowRemarks"
      aFormulaValue.Add CStr("'1'")
   Else
'      MdiSect.Crw.Formulas(1) = "ShowRemarks='0'"
      aFormulaName.Add "ShowRemarks"
      aFormulaValue.Add CStr("'0'")
   End If
'   MdiSect.Crw.Formulas(2) = "ShowMos=" & optAll
    aFormulaName.Add "ShowMos"
    aFormulaValue.Add optAll
    sCustomReport = GetCustomReport("finap01.rpt")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
'   MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{VihdTable.VIVENDOR} = '" _
          & Compress(cmbVnd) & "' AND " _
          & "{VihdTable.VINO} = '" _
          & Trim(cmbInv) & "'"
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
   cCRViewer.CRViewerSize Me
   ' Set report parameter
   cCRViewer.SetDbTableConnection
   ' print the copies
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   cCRViewer.ShowGroupTree False
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
'   MdiSect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub PrintReport1()
   Dim sCustomerReport As String
   On Error GoTo DiaErr1
   MouseCursor 13
   'SetMdiReportsize MdiSect
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   If optRem.Value = vbChecked Then
      MdiSect.crw.Formulas(1) = "ShowRemarks='1'"
   Else
      MdiSect.crw.Formulas(1) = "ShowRemarks='0'"
   End If
   MdiSect.crw.Formulas(2) = "ShowMos=" & optAll
   sCustomReport = GetCustomReport("finap01.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{VihdTable.VIVENDOR} = '" _
          & Compress(cmbVnd) & "' AND " _
          & "{VihdTable.VINO} = '" _
          & Trim(cmbInv) & "'"
   MdiSect.crw.SelectionFormula = sSql
   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = RTrim(optDsc.Value) _
              & RTrim(optExt.Value) _
              & RTrim(optItm.Value) _
              & RTrim(optAll.Value) _
              & RTrim(optDsb.Value) _
              & RTrim(optRem.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optDsc.Value = Val(Left(sOptions, 1))
      optExt.Value = Val(Mid(sOptions, 2, 1))
      optItm.Value = Val(Mid(sOptions, 3, 1))
      optAll.Value = Val(Mid(sOptions, 4, 1))
      optDsb.Value = Val(Mid(sOptions, 5, 1))
      optRem.Value = Val(Mid(sOptions, 6, 1))
   Else
      optDsc.Value = vbChecked
      optExt.Value = vbChecked
      optItm.Value = vbChecked
      optAll.Value = vbChecked
      optDsb.Value = vbChecked
      optRem.Value = vbChecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub


Private Sub optPrn_Click()
   PrintReport
End Sub

Public Sub GetInvoices()
   Dim RdoInv As ADODB.Recordset
   Dim iTotal As Integer
   Dim sVendor As String
   On Error GoTo DiaErr1
   cmbInv.Clear
   sVendor = Compress(cmbVnd)
   'RdoQry(0) = sVendor
   RdoQry.parameters(0).Value = sVendor
   bSqlRows = clsADOCon.GetQuerySet(RdoInv, RdoQry)
   If bSqlRows Then
      With RdoInv
         Do Until .EOF
            iTotal = iTotal + 1
            AddComboStr cmbInv.hwnd, "" & Trim(!VINO)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   If cmbInv.ListCount > 0 Then cmbInv = cmbInv.List(0)
   lblCnt = iTotal
   Set RdoInv = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getinvoices"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
