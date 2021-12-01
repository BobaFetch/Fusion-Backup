VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SaleSLp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer List"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6780
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLp02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      Top             =   960
      Width           =   1555
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5640
      TabIndex        =   14
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "SaleSLp02a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   490
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "SaleSLp02a.frx":0938
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
      Left            =   5640
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.TextBox txtDiv 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Tag             =   "3"
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txtTyp 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Tag             =   "3"
      Top             =   1800
      Width           =   375
   End
   Begin VB.CheckBox optAdr 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   2160
      TabIndex        =   3
      Top             =   2520
      Value           =   1  'Checked
      Width           =   735
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   6780
   End
   Begin VB.Label lblCUName 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2160
      TabIndex        =   18
      Top             =   1320
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   8
      Left            =   3960
      TabIndex        =   16
      Top             =   2280
      Width           =   1932
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   7
      Left            =   3960
      TabIndex        =   15
      Top             =   1920
      Width           =   1932
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   6
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   1992
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   5
      Left            =   3960
      TabIndex        =   11
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Addresses"
      Height          =   288
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   1872
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Division(s)"
      Height          =   288
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1992
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Type(s)"
      Height          =   288
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1992
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1830
      TabIndex        =   7
      Top             =   1680
      Width           =   105
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Nickname(s)"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "SaleSLp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'1/11/07 Added GetThisCustomer 7.2.1
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT CUREF,CUNICKNAME FROM CustTable"
   LoadComboBox cmbCst
   cmbCst = "ALL"
   GetThisCustomer
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

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sTyp As String * 2
   Dim sDiv As String * 2
   
   If txtTyp = "ALL" Then txtTyp = ""
   If txtDiv = "ALL" Then txtDiv = ""
   sTyp = txtTyp
   sDiv = txtDiv
   'Save by Menu Option
   sOptions = RTrim(optAdr.Value) _
              & sTyp _
              & sDiv
   SaveSetting "Esi2000", "EsiSale", "sl02", Trim(sOptions)
   
End Sub


Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiSale", "sl02", sOptions)
   If Len(sOptions) > 0 Then
      optAdr.Value = Val(Left(sOptions, 1))
      txtTyp = Mid(sOptions, 2, 2)
      txtDiv = Mid(sOptions, 4, 2)
   End If
   
End Sub

Private Sub cmbCst_Click()
   GetThisCustomer
   
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   If Len(cmbCst) = 0 Then cmbCst = "ALL"
   GetThisCustomer
   
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
   MouseCursor 13
   If bOnLoad Then FillCombo
   bOnLoad = 0
   MdiSect.lblBotPanel = Caption
   
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
   Set SaleSLp02a = Nothing
   
End Sub

Private Sub PrintReport()
    MouseCursor 13
    Dim sDiv As String
    Dim sCust As String
    Dim sType As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   
   If Len(Trim(cmbCst)) = 0 Then cmbCst = "ALL"
   If Len(Trim(txtTyp)) = 0 Then txtTyp = "AL"
   If Len(Trim(txtDiv)) = 0 Then txtDiv = "AL"
   If cmbCst = "ALL" Then
      sCust = ""
   Else
      sCust = Compress(cmbCst)
   End If
   If Left(txtTyp, 2) = "AL" Then sType = "" Else sType = txtTyp
   If Left(txtDiv, 2) = "AL" Then sDiv = "" Else sDiv = txtDiv
   On Error GoTo DiaErr1
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowAddress"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbCst) & "...'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optAdr.Value
   

   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("sleco02")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{CustTable.CUREF} LIKE '" & sCust & "*' " _
          & "AND {CustTable.CUTYPE} LIKE '" & sType & "*' " _
          & "AND {CustTable.CUDIVISION} LIKE '" & sDiv & "*' "
   
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

Private Sub optAdr_KeyDown(KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub optAdr_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   MouseCursor 13
   PrintReport
   
End Sub

Private Sub txtDiv_LostFocus()
   txtDiv = CheckLen(txtDiv, 2)
   If Len(txtDiv) = 0 Then txtDiv = "ALL"
   
End Sub

Private Sub txtTyp_LostFocus()
   txtTyp = CheckLen(txtTyp, 2)
   If Len(txtTyp) = 0 Then txtTyp = "ALL"
   
End Sub
