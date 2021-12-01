VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SaleSLp06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer List By Zip Code"
   ClientHeight    =   2970
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7110
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2970
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLp06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   13
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "SaleSLp06a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "SaleSLp06a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   4
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.CheckBox optNte 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1800
      Width           =   1555
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   1
      Tag             =   "1"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Tag             =   "1"
      Top             =   1080
      Width           =   1335
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
      FormDesignHeight=   2970
      FormDesignWidth =   7110
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   6
      Left            =   4600
      TabIndex        =   14
      Top             =   1800
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   5
      Left            =   4600
      TabIndex        =   11
      Top             =   1440
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   4
      Left            =   4600
      TabIndex        =   10
      Top             =   1080
      Width           =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Selling Notes"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   2625
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Nickname"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   2625
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Up To But Not Including:"
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   2745
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customers With Zip Codes Matching:"
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   2745
   End
End
Attribute VB_Name = "SaleSLp06a"
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

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sZip0 As String * 10
   Dim sZip1 As String * 10
   sZip0 = txtZip(0)
   sZip1 = txtZip(1)
   
   'Save by Menu Option
   sOptions = RTrim(optNte.Value) _
              & sZip0 & sZip1
   SaveSetting "Esi2000", "EsiSale", "sl07", Trim(sOptions)
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "sl07", sOptions)
   If Len(sOptions) > 0 Then
      optNte.Value = Val(Left(sOptions, 1))
      txtZip(0) = Trim(Mid(sOptions, 2, 10))
      txtZip(1) = Trim(Mid(sOptions, 12, 10))
   End If
   
End Sub

Private Sub cmbCst_GotFocus()
   SelectFormat Me
   
End Sub

Private Sub cmbCst_KeyPress(KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Len(cmbCst) = 0 Then cmbCst = "ALL"
   
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
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      cmbCst.AddItem "ALL"
      FillCustomers
      cmbCst = "ALL"
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   bOnLoad = 1
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
   Set SaleSLp06a = Nothing
   
End Sub

Private Sub PrintReport()
   Dim sCust As String
   Dim sBegZip As String
   Dim sEndZip As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   If cmbCst = "ALL" Then sCust = "" Else sCust = Compress(cmbCst)
   sBegZip = Trim(txtZip(0))
   If sBegZip = "" Then sBegZip = "0"
   sEndZip = Trim(txtZip(1))
   If sEndZip = "" Then sEndZip = "999999"
   On Error GoTo DiaErr1
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowNotes"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Customer " & CStr(cmbCst _
                        & ", Zip Codes " & txtZip(0)) & "...'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optNte.Value
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("sleco07")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{CustTable.CUREF} LIKE '" & sCust & "*' " _
          & "AND {CustTable.CUZIP}>='" & sBegZip & "' " _
          & "AND {CustTable.CUZIP}<'" & sEndZip & "' "
          
   cCRViewer.SetReportSelectionFormula sSql
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
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

Private Sub optDis_Click()
   If Len(Trim(cmbCst)) = 0 Then cmbCst = "ALL"
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub optNte_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optNte_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Sub optPrn_Click()
   If Len(Trim(cmbCst)) = 0 Then cmbCst = "ALL"
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub txtZip_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtZip_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtZip_LostFocus(Index As Integer)
   txtZip(Index) = CheckLen(txtZip(Index), 10)
   
End Sub
