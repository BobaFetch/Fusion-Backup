VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SaleSLp07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Sales Order List By PO"
   ClientHeight    =   2055
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   7065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLp07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5880
      TabIndex        =   16
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "SaleSLp07a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "SaleSLp07a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox txtCpo 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Customer Purchase Orders"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Retrieves List"
      Top             =   1560
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.ListBox lstSos 
      Height          =   1425
      Left            =   2040
      TabIndex        =   2
      ToolTipText     =   "Double Click Selection To Open Sales Order"
      Top             =   2280
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Tag             =   "3"
      Top             =   840
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5880
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   3840
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2055
      FormDesignWidth =   7065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   6
      Left            =   4560
      TabIndex        =   14
      Top             =   1560
      Width           =   1785
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   960
      TabIndex        =   13
      ToolTipText     =   "Total Sales Orders Found"
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Count"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label txtSon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label txtPre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order                                                "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3120
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Orders"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   1425
   End
   Begin VB.Label txtNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO's Beginning With"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Nickname"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1750
   End
End
Attribute VB_Name = "SaleSLp07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/6/05 Correct the ComboBox (omit "ALL") and the FormUnload
Option Explicit
Dim bOnLoad As Byte
Dim bSoSel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCst_Click()
   If lstSos.ListCount > 0 Then lstSos.Clear
   FindCustomer Me, cmbCst, False
   FillPurchaseOrders
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   FindCustomer Me, cmbCst, False
   FillPurchaseOrders
   
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




Private Sub cmdSel_Click()
   SelectSalesOrders
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCustomers
      If cmbCst.ListCount > 0 Then
         cmbCst = cmbCst.List(0)
         FindCustomer Me, cmbCst, False
      End If
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Public Sub FillPurchaseOrders()
   txtCpo.Clear
   sSql = "SELECT DISTINCT SOPO FROM SohdTable WHERE (SOCUST='" _
          & Compress(cmbCst) & "' AND SOPO<>'')"
   LoadComboBox txtCpo, -1
   If txtCpo.ListCount > 0 Then txtCpo = txtCpo.List(0)
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If bSoSel = 0 Then FormUnload
   Set SaleSLp07a = Nothing
   
End Sub






Private Sub SelectSalesOrders()
   Dim RdoGet As ADODB.Recordset
   Dim sCust As String
   Dim sPoNum As String
   lstSos.Clear
   On Error GoTo DiaErr1
   If cmbCst = "ALL" Then
      sCust = ""
   Else
      sCust = Compress(cmbCst)
   End If
   sPoNum = Trim(txtCpo)
   On Error GoTo DiaErr1
   sSql = "SELECT SONUMBER,SOTYPE,SOCUST,SOPO FROM SohdTable WHERE " _
          & "(SOCUST LIKE '" & sCust & "%' AND SOPO LIKE '" & sPoNum & "%') " _
          & "AND SOCANCELED=0 ORDER BY SOPO,SONUMBER "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         Do Until .EOF
            lstSos.AddItem !SOTYPE & Format(!SoNumber, SO_NUM_FORMAT) & String(10, Chr(160)) & !SOPO
            .MoveNext
         Loop
         ClearResultSet RdoGet
      End With
   End If
   Set RdoGet = Nothing
   lblCount = Format(lstSos.ListCount, "#####0")
   cmdSel.Enabled = False
   Exit Sub
   
DiaErr1:
   sProcName = "selectsal"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub PrintReport()
   Dim sCust As String
   Dim sSopo As String
   
   MouseCursor 13
   If cmbCst = "ALL" Then sCust = "" Else sCust = Compress(cmbCst)
   sSopo = Trim(txtCpo)
   On Error GoTo DiaErr1
   
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'Customer" & CStr(cmbCst _
                        & ", Purchase Orders " & sSopo) & "...'")
    aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
    
   sCustomReport = GetCustomReport("sleco08")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{CustTable.CUREF} LIKE '" & sCust & "*' " _
          & "AND {SohdTable.SOPO} LIKE '" & sSopo & "*' "
   
   cCRViewer.SetReportSelectionFormula sSql
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


Private Sub lstSos_DblClick()
   Dim sSONum As String
   On Error Resume Next
   sSONum = Left(lstSos.List(lstSos.ListIndex), 7)
   If Val(Right(sSONum, 6)) > 0 Then
      txtSon = Right(sSONum, 6)
      txtPre = Left(sSONum, 1)
      bSoSel = 1
      SaleSLe02a.optLst = vbChecked
      SaleSLe02a.Show
   Else
      MsgBox "No Sales Order Selected.", vbInformation, Caption
   End If
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub txtCpo_Click()
   cmdSel.Enabled = True
   
End Sub

Private Sub txtCpo_GotFocus()
   txtCpo_Click
   
End Sub


Private Sub txtCpo_LostFocus()
   txtCpo = CheckLen(txtCpo, 20)
   SelectSalesOrders
   
End Sub


Private Sub txtNme_Change()
   If txtNme = "*** Customer Wasn't Found ***" Then
      txtNme.ForeColor = ES_RED
   Else
      txtNme.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub txtPre_Click()
   'store prefix
   
End Sub

Private Sub txtSon_Click()
   'stores so num
   
End Sub




Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub
