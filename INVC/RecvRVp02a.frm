VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RecvRVp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receiving Log By Vendor"
   ClientHeight    =   4260
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   7230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   2040
      TabIndex        =   25
      Top             =   2400
      Width           =   855
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   2040
      TabIndex        =   24
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RecvRVp02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.Frame Z2 
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      ToolTipText     =   "Select Types - Right And Left Arrow Keys"
      Top             =   1080
      Width           =   3255
      Begin VB.OptionButton optShow 
         Caption         =   "All"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Show All Part Numbers"
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optShow 
         Caption         =   "Raw"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Part Types 4 And 5 Only"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optShow 
         Caption         =   "Service"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   20
         ToolTipText     =   "Service Part Type 7 Only"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4200
      TabIndex        =   5
      Tag             =   "4"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Tag             =   "4"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   19
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Picture         =   "RecvRVp02a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   600
         Picture         =   "RecvRVp02a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print The Report"
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Vendor From List"
      Top             =   360
      Width           =   1555
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   225
      Left            =   2040
      TabIndex        =   6
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CheckBox optMoa 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   225
      Left            =   2040
      TabIndex        =   7
      Top             =   3720
      Width           =   735
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   2880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4260
      FormDesignWidth =   7230
   End
   Begin VB.Label lblProdClass 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3360
      TabIndex        =   31
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label lblProdCode 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3360
      TabIndex        =   30
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   7
      Left            =   5760
      TabIndex        =   29
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   28
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Product Class"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Product Code"
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "(Blank for ALL)"
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Part Types"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   21
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   17
      Top             =   720
      Width           =   3720
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Nickname"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   16
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   8
      Left            =   5520
      TabIndex        =   15
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   13
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Allocations"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   1815
   End
End
Attribute VB_Name = "RecvRVp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bGoodVendor As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "rc02", sOptions)
   If Len(sOptions) > 0 Then
      optExt.Value = Val(Left(sOptions, 1))
      optMoa.Value = Val(Mid(sOptions, 2, 1))
      ' txtBeg = Mid(sOptions, 3, 8)
      ' txtEnd = Mid(sOptions, 11, 8)
   
        If Len(sOptions) > 18 Then
            cmbCls = Trim(Mid(sOptions, 19, 4))
            cmbCde = Trim(Mid(sOptions, 23, 6))
        End If
       
   
   End If
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = "01/01/" & Right(txtEnd, 4)
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sBeg As String * 8
   Dim sEnd As String * 8
   
   sBeg = txtBeg
   sEnd = txtEnd
   'Save by Menu Option
   sOptions = RTrim(optExt.Value) _
              & RTrim(optMoa.Value) & sBeg & sEnd & Left(cmbCls & Space(4), 4) & Left(cmbCde & Space(6), 6)
   SaveSetting "Esi2000", "EsiProd", "rc02", Trim(sOptions)
   
End Sub



Private Sub cmbCde_Click()
    lblProdCode = GetProductCode()

End Sub

Private Sub cmbCde_LostFocus()
    If Compress(cmbCde) = "" Then cmbCde = "ALL"
    lblProdCode = GetProductCode()
    
End Sub

Private Sub cmbCls_Click()
    lblProdClass = GetProductClass()
End Sub

Private Sub cmbCls_LostFocus()
    If Compress(cmbCls) = "" Then cmbCls = "ALL"
    lblProdClass = GetProductClass()
    
End Sub

Private Sub cmbVnd_Click()
   bGoodVendor = FindVendor(True)
   
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   bGoodVendor = FindVendor(True)

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
   If bOnLoad Then
      FillVendors
      If cmbVnd.ListCount > 0 Then cmbVnd = cmbVnd.List(0)
      If cUR.CurrentVendor <> "" Then cmbVnd = cUR.CurrentVendor
      bGoodVendor = FindVendor(True)
      
      FillProductClasses
      FillProductCodes
      
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   bOnLoad = 1
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtBeg = "01/01/" & Right(txtEnd, 4)
   GetOptions
   
   lblProdClass = GetProductClass()
   lblProdCode = GetProductCode()
   
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   cUR.CurrentVendor = cmbVnd
   SaveCurrentSelections
   FormUnload
   Set RecvRVp02a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sBegDate As String
   Dim sEndDate As String
   Dim sVendor As String
   Dim sProdCode, sProdClass As String
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   If IsDate(txtBeg) Then
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   Else
      sBegDate = "1995,01,01"
   End If
   If IsDate(txtEnd) Then
      sEndDate = Format(txtEnd, "yyyy,mm,dd")
   Else
      sEndDate = "2024,12,31"
   End If
   If Compress(cmbVnd) = "ALL" Or Compress(cmbVnd) = "" Then sVendor = "" Else sVendor = Compress(cmbVnd)

   If Compress(cmbCls) = "" Or Compress(cmbCls) = "ALL" Then sProdClass = "" Else sProdClass = Compress(cmbCls)
   If Compress(cmbCde) = "" Or Compress(cmbCde) = "ALL" Then sProdCode = "" Else sProdCode = Compress(cmbCde)
   
   
'   sVendor = Compress(cmbVnd)
   On Error GoTo Prc01
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowExDesc"
   aFormulaName.Add "ShowMOa"
   aFormulaValue.Add CStr("'Receipts From " & CStr(txtBeg & " Ending " & txtEnd) & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add optExt.Value
   aFormulaValue.Add optMoa.Value
   sCustomReport = GetCustomReport("prdrc02")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.ShowGroupTree False
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   sSql = "{PoitTable.PITYPE} in [15, 17] AND {PoitTable.PIADATE} in Date(" & sBegDate & ") to Date(" & sEndDate & ") " _
          & "AND {VndrTable.VEREF} LIKE '" & sVendor & "*' "
   If sProdClass <> "" Then sSql = sSql & " AND {PartTable.PACLASS}='" & sProdClass & "' "
   If sProdCode <> "" Then sSql = sSql & " AND {PartTable.PAPRODCODE}='" & sProdCode & "' "
   
   If optShow(1).Value = True Then
      sSql = sSql & " AND ({PartTable.PALEVEL}=4 OR {PartTable.PALEVEL}=5) "
   Else
      If optShow(2).Value = True Then _
                 sSql = sSql & " AND {PartTable.PALEVEL}=7"
   End If
   cCRViewer.SetReportSelectionFormula (sSql)
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub
   
Prc01:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub optDis_Click()
   If bGoodVendor Then
      PrintReport
   Else
      MsgBox "Requires a Valid Vendor.", vbInformation, Caption
   End If
   
End Sub

Private Sub optExt_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optMoa_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optMoa_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optPrn_Click()
   If bGoodVendor Or Compress(cmbVnd) = "ALL" Then
      PrintReport
   Else
      MsgBox "Requires a Valid Vendor.", vbInformation, Caption
   End If
   
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


Private Function GetProductClass()
   On Error Resume Next
   Dim RdoCode As ADODB.Recordset
   sSql = "SELECT CCREF,CCDESC FROM PclsTable WHERE " _
          & "CCREF='" & Compress(cmbCls) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCode, ES_FORWARD)
   If bSqlRows Then GetProductClass = "" & Trim(RdoCode!CCDESC) _
                                      Else GetProductClass = "* ALL *"
   Set RdoCode = Nothing
   
End Function

Private Function GetProductCode() As String
   On Error Resume Next
   Dim RdoCode As ADODB.Recordset
   sSql = "SELECT PCREF,PCDESC FROM PcodTable WHERE " _
          & "PCREF='" & Compress(cmbCde) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCode, ES_FORWARD)
   If bSqlRows Then GetProductCode = "" & Trim(RdoCode!PCDESC) _
                                     Else GetProductCode = "* ALL *"
   Set RdoCode = Nothing
   
End Function


