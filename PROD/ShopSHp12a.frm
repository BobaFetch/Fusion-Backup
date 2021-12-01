VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHp12a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Order Sales Order Allocations"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPrt 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1320
      Width           =   3255
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "ShopSHp12a.frx":0000
      Height          =   315
      Left            =   4800
      Picture         =   "ShopSHp12a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   1320
      Width           =   350
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHp12a.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtSon 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Enter Sales Order Number Or Blank For All"
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Contains Part Number With Runs Not CA"
      Top             =   1320
      Width           =   3545
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   6040
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Select Run Number"
      Top             =   1320
      Width           =   1000
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
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "ShopSHp12a.frx":0E32
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "ShopSHp12a.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   7125
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   6
      Left            =   5280
      TabIndex        =   17
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2520
      TabIndex        =   16
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   960
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   13
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblSta 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6600
      TabIndex        =   12
      Top             =   1680
      Width           =   396
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "ShopSHp12a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/9/05 fixed the report and force Part Number
Option Explicit
Dim bOnLoad As Byte
Dim bCanceled As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Click()
   GetCurrentPart cmbPrt, lblDsc
   GetRuns
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bCanceled Then Exit Sub
   If Len(Trim(cmbPrt)) Then
      txtSon = "ALL"
      lblNme = ""
      lblTyp = ""
      GetCurrentPart cmbPrt, lblDsc
      GetRuns
   Else
      lblDsc = ""
      lblSta = ""
   End If
   
End Sub

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   cmbPrt = txtPrt
   cmbPrt = CheckLen(cmbPrt, 30)
   If bCanceled Then Exit Sub
   If Len(Trim(cmbPrt)) Then
      txtSon = "ALL"
      lblNme = ""
      lblTyp = ""
      GetCurrentPart cmbPrt, lblDsc
      GetRuns
   Else
      lblDsc = ""
      lblSta = ""
   End If
End Sub



Private Sub cmbRun_Click()
   GetThisRun
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   If Val(cmbRun) > 0 Then
      txtSon = "ALL"
      lblNme = ""
      lblTyp = ""
      'GetRuns
   Else
      lblDsc = ""
      lblSta = ""
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = 1
   
End Sub


Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPrt = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPrt
   ViewParts.Show
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
   MouseCursor 13
   sSql = "Qry_RunsNotCanceled"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetRuns()
   On Error GoTo DiaErr1
   cmbRun.Clear
   If Trim(cmbPrt) = "" Then Exit Sub
   MouseCursor 13
   sSql = "SELECT RUNNO FROM RunsTable WHERE RUNREF='" & Compress(cmbPrt) & "' "
   LoadNumComboBox cmbRun, "####0"
   If Not bSqlRows Then lblDsc = "*** No Runs Found For That Part ***"
   MouseCursor 0
   If cmbRun.ListCount > 0 Then
      cmbRun = cmbRun.List(0)
      If GetPreferenceValue("AutoSelectLastRun") = "1" Then cmbRun = cmbRun.List(cmbRun.ListCount - 1)
      GetThisRun
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillCombo
      
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
   Set ShopSHp12a = Nothing
   
End Sub

Private Sub PrintReport()
'   MouseCursor 13
'   On Error GoTo Psh01
'   SetMdiReportsize MDISect
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "RequestBy = 'Requested By: " & sInitials & "'"
'   MDISect.Crw.Formulas(2) = "PartNumber='" & Trim(cmbPrt) & "'"
'   MDISect.Crw.Formulas(3) = "RunNumber='" & Trim(cmbRun) & "'"
'   sCustomReport = GetCustomReport("prdsh17")
'   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
'
''   sSql = "{RunsTable.RUNREF}='" & Compress(cmbPrt) & "' " _
''          & "AND {RunsTable.RUNNO}=" & Val(cmbRun) & " "
''   If Val(txtSon) > 0 Then
''      sSql = sSql & "AND {RnalTable.RASO}=" & Val(txtSon) & " "
''   End If
''   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
'   MouseCursor 0
'   Exit Sub
'
'Psh01:
'   sProcName = "printreport"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   Resume Psh02
'Psh02:
'   DoModuleErrors Me
'

PrintReportSalesOrderAllocations Me, Trim(cmbPrt), cmbRun

End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtSon = "ALL"
   
End Sub

Private Sub SaveOptions()
   
End Sub

Private Sub GetOptions()
   
End Sub


Private Sub lblDsc_Change()
   If Trim(lblDsc) = "" Then lblDsc = "*** No Part Number Or Valid Run Selected ***"
   If lblDsc = "" Or Left(lblDsc, 6) = "*** No" Then
      optPrn.Enabled = False
      optDis.Enabled = False
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = Es_TextForeColor
      optPrn.Enabled = True
      optDis.Enabled = True
   End If
   
End Sub

Private Sub optDis_Click()
   If lblDsc.ForeColor = ES_RED Then
      MsgBox "Requires A Valid Part And Run.", _
         vbExclamation, Caption
   Else
      PrintReport
   End If
   
End Sub


Private Sub optPrn_Click()
   If lblDsc.ForeColor = ES_RED Then
      MsgBox "Requires A Valid Part And Run.", _
         vbExclamation, Caption
   Else
      PrintReport
   End If
   
   
End Sub



Private Sub GetThisRun()
   Dim RdoRun As ADODB.Recordset
   If Val(txtSon) > 0 Then Exit Sub
   sSql = "SELECT RUNREF,RUNSTATUS,PARTREF,PARTNUM,PADESC FROM " _
          & "RunsTable, PartTable WHERE RUNREF=PARTREF AND (RUNNO=" _
          & Trim(cmbRun) & " AND RUNREF='" & Compress(cmbPrt) & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         lblDsc = "" & Trim(!PADESC)
         lblSta = "" & Trim(!RUNSTATUS)
         ClearResultSet RdoRun
      End With
   Else
      lblDsc = "*** No Matching Run Found ***"
   End If
   
   Set RdoRun = Nothing
   
End Sub



Private Sub txtSon_LostFocus()
   On Error Resume Next
   txtSon = CheckLen(txtSon, SO_NUM_SIZE)
   If Val(txtSon) > 0 Then
      txtSon = Format(Abs(Val(txtSon)), SO_NUM_FORMAT)
      GetSalesOrder
   Else
      txtSon = "ALL"
   End If
   
End Sub



Private Sub GetSalesOrder()
   Dim RdoSon As ADODB.Recordset
   sSql = "Qry_GetSalesOrderCustomer " & Val(txtSon)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
   If bSqlRows Then
      With RdoSon
         lblTyp = "" & Trim(!SOTYPE)
         lblNme = "" & Trim(!CUNICKNAME) & "-" & Trim(!CUNAME)
         ClearResultSet RdoSon
      End With
   End If
   Set RdoSon = Nothing
   
End Sub

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPrt.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPrt.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function

