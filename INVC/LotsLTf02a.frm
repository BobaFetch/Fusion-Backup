VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form LotsLTf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lot Quantity Reconciliation"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   5551
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   720
      Width           =   3075
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotsLTf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   600
         Picture         =   "LotsLTf02a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "LotsLTf02a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "LotsLTf02a.frx":0AB6
      Height          =   315
      Left            =   6240
      Picture         =   "LotsLTf02a.frx":0F90
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "View Selection"
      Top             =   720
      Width           =   350
   End
   Begin VB.CheckBox optQoh 
      Alignment       =   1  'Right Justify
      Caption         =   "Update Quantity On Hand"
      Height          =   435
      Left            =   3600
      TabIndex        =   4
      ToolTipText     =   "Sets Quantity On Hand Equal To Lot Remaining Quantity"
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5760
      TabIndex        =   17
      ToolTipText     =   "Update Part Number And Apply Changes"
      Top             =   1560
      Width           =   875
   End
   Begin VB.CheckBox optLots 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   16
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5760
      TabIndex        =   3
      ToolTipText     =   "Select Values For The Current Part Number"
      Top             =   1080
      Width           =   875
   End
   Begin VB.TextBox txtLoi 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Sum Of The Lot Transactions For This Part"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtLor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Lot Remaining Quantity (Part Table)"
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox TxtPaq 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Part Quantity On Hand (Part Table)"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox lblDsc 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "LotsLTf02a.frx":139D
      Height          =   315
      Left            =   5520
      Picture         =   "LotsLTf02a.frx":16DF
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   720
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Part Number(Blank For All)"
      Top             =   750
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5760
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3255
      FormDesignWidth =   6705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sum InvaTable (Not visible)"
      Height          =   285
      Index           =   4
      Left            =   3120
      TabIndex        =   24
      ToolTipText     =   "Lot Tracking"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Label lblEdit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Set Lot Remaining Quantity Equal To The Sum Of The Lots (See Help)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   360
      TabIndex        =   23
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Activity 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   5520
      TabIndex        =   22
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Tracked Part Number"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   15
      ToolTipText     =   "Lot Tracking"
      Top             =   2640
      Width           =   1995
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sum Of Lot Quantities"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "Sum Of The Lot Transactions For This Part"
      Top             =   2280
      Width           =   1995
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Remaining Quantity"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "Lot Remaining Quantity (Part Table)"
      Top             =   1920
      Width           =   1995
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity On Hand"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Part Quantity On Hand (Part Table)"
      Top             =   1560
      Width           =   1995
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   21
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   22
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1305
   End
End
Attribute VB_Name = "LotsLTf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'7/18/05 new
Option Explicit
Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bGoodPart As Byte
Dim cStdCost As Currency
Dim sDebitAcct As String
Dim sCreditAcct As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdFnd_Click()
   ZeroBoxes
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   ViewParts.Show
   bCancel = 0
   
End Sub

Private Sub cmdFnd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 5551
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdSel_Click()
   If lblDsc.ForeColor = vbBlack Then
      bGoodPart = GetPartInfo()
   End If
   
End Sub

Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   If optQoh.Value = vbChecked Then
      sMsg = "Do You Want To Set The Lot Remaining Quantity " & vbCr _
             & "(Part Table) Equal To The Sum Of the Lot Quantities " & vbCr _
             & "And Set The Quantity On Hand The Same (Part Table).."
   Else
      sMsg = "Do You Want To Set The Lot Remaining Quantity " & vbCr _
             & "(Part Table) Equal To The Sum Of the Lot Quantities " & vbCr _
             & "And Leave The Quantity On Hand As Is (Part Table).."
   End If
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then AdjustQuantity Else CancelTrans
   
   
End Sub

Private Sub cmdVew_Click()
   If lblDsc.ForeColor = vbBlack Then PrintReport
   
End Sub


Private Sub PrintReport()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
      
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr("Requested By: " & sInitials) & "'")
   sCustomReport = GetCustomReport("invlt06")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
   
   sSql = "{PartTable.PARTREF}='" & Compress(cmbPrt) & "' "
   cCRViewer.SetReportSelectionFormula (sSql)
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   

   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   MouseCursor 0
   Exit Sub
   
   
   
 '  Dim sPartNumber As String
 '
 '  MouseCursor 13
 '  'SetMdiReportsize MdiSect
 '  MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
 '  MdiSect.Crw.Formulas(1) = "RequestBy='Requested By: " & sInitials & "'"
 '  sCustomReport = GetCustomReport("invlt06")
 '  MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
 '  sSql = "{PartTable.PARTREF}='" & Compress(txtPrt) & "' "
 '  MdiSect.Crw.SelectionFormula = sSql
 '  'SetCrystalAction Me
 '  MouseCursor 0
 '  Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then CreateEditTable
   bOnLoad = 0
   MouseCursor 0
   FillCombo
   cmbPrt = ""
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set LotsLTf02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblDsc.BackColor = Me.BackColor
   TxtPaq.BackColor = Es_TextDisabled
   txtLor.BackColor = Es_TextDisabled
   txtLoi.BackColor = Es_TextDisabled
   lblDsc.ForeColor = ES_RED
   ZeroBoxes
   
End Sub

Private Sub ZeroBoxes()
   TxtPaq = "0.000"
   txtLor = "0.000"
   txtLoi = "0.000"
   Activity = "0.000"
   cmdUpd.Enabled = False
   optLots.Value = vbUnchecked
   optQoh.Enabled = True
   
End Sub

Private Sub lblDsc_Change()
   ZeroBoxes
   If Left(lblDsc, 4) = "*** " Then
      lblDsc.ForeColor = ES_RED
      cmdSel.Enabled = False
   Else
      lblDsc.ForeColor = vbBlack
      cmdSel.Enabled = True
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
   ZeroBoxes
   If bCancel = 1 Then Exit Sub
   If Trim(txtPrt) <> "" Then GetCurrentPart txtPrt, lblDsc
   
End Sub

Private Sub cmbPrt_LostFocus()
   ZeroBoxes
   
   If bCancel = 1 Then Exit Sub
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   If Trim(cmbPrt) <> "" Then GetCurrentPart cmbPrt, lblDsc
   
End Sub

Private Sub cmbPrt_Change()
   ZeroBoxes
   If bCancel = 1 Then Exit Sub
   If Trim(cmbPrt) <> "" Then GetCurrentPart cmbPrt, lblDsc
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE " _
          & "PALOTTRACK=1 AND PAINACTIVE = 0 AND PAOBSOLETE = 0 ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetPartInfo() As Byte
   Dim RdoQoh As ADODB.Recordset
   Dim sLotQty As Currency
   On Error GoTo DiaErr1
   sSql = "SELECT PAQOH,PALOTQTYREMAINING,PALOTTRACK,PASTDCOST FROM " _
          & "PartTable WHERE PARTREF='" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoQoh, ES_FORWARD)
   If bSqlRows Then
      With RdoQoh
         TxtPaq = Format(!PAQOH, ES_QuantityDataFormat)
         txtLor = Format(!PALOTQTYREMAINING, ES_QuantityDataFormat)
         optLots.Value = !PALOTTRACK
         cStdCost = !PASTDCOST
         ClearResultSet RdoQoh
      End With
      sLotQty = GetRemainingLotQty(Compress(cmbPrt), True)
      txtLoi = Format(sLotQty, ES_QuantityDataFormat)
      sLotQty = GetActivityQuantity(Compress(cmbPrt))
      Activity = Format(sLotQty, ES_QuantityDataFormat)
      If Val(TxtPaq) < Val(txtLoi) Then
         optQoh.Value = vbChecked
         optQoh.Enabled = False
      End If
      If Val(TxtPaq) < 0 Then
         optQoh.Value = vbChecked
         optQoh.Enabled = False
      End If
      cmdUpd.Enabled = True
   Else
      cStdCost = 0
   End If
   Set RdoQoh = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpartinfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AdjustQuantity()
   Dim bByte As Byte
   Dim lNextRow As Long
   Dim lCOUNTER As Long
   Dim lSysCount As Long
   Dim cAdjQty As Currency
   
   lNextRow = GetNextRow
   On Error Resume Next
   bByte = GetPartAccounts(Compress(cmbPrt), sCreditAcct, sDebitAcct)
   
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   sSql = "UPDATE PartTable SET PALOTQTYREMAINING=" & Val(txtLoi) & " " _
          & "WHERE PARTREF='" & Compress(cmbPrt) & "'"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "INSERT INTO LorcTable (LOTREC_ROW,LOTREC_PARTREF,LOTREC_PAQOH," _
          & "LOTREC_PALOTQTYREMAINING,LOTREC_SUMLOTS,LOTREC_SUMACTIVITY) VALUES(" _
          & lNextRow & ",'" & Compress(cmbPrt) & "'," & Val(TxtPaq) & "," _
          & Val(txtLor) & "," & Val(txtLoi) & "," _
          & Val(Activity) & ")"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      lCOUNTER = GetLastActivity() + 1
      lSysCount = lCOUNTER
      If optQoh.Value = vbChecked Then
         sSql = "UPDATE PartTable SET PAQOH=" & Val(txtLoi) & " " _
                & "WHERE PARTREF='" & Compress(cmbPrt) & "'"
         clsADOCon.ExecuteSQL sSql
         
         If Val(Activity) <> Val(txtLoi) Then
            cAdjQty = Val(txtLoi) - Val(Activity)
            sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
                   & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INUSER) " _
                   & "VALUES(19,'" & Compress(cmbPrt) & "','Manual Adjustment','LOT Reconciliation'," _
                   & "'" & Format(Now, "mm/dd/yy") & "','" & Format(Now, "mm/dd/yy") & "'," & cAdjQty _
                   & "," & cAdjQty & "," & cStdCost & ",'" & sCreditAcct _
                   & "','" & sDebitAcct & "'," & lCOUNTER & ",'" & sInitials & "')"
            clsADOCon.ExecuteSQL sSql
         End If
         sSql = "UPDATE LorcTable SET LOTREC_PARTADJ='Y' WHERE " _
                & "LOTREC_ROW=" & lNextRow & " "
         clsADOCon.ExecuteSQL sSql
      Else
         If Val(TxtPaq) <> Val(Activity) Then
            cAdjQty = Val(TxtPaq) - Val(Activity)
            sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE," _
                   & "INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INNUMBER,INUSER) " _
                   & "VALUES(19,'" & Compress(cmbPrt) & "','Manual Adjustment','LOT Reconciliation'," _
                   & "'" & Format(Now, "mm/dd/yy") & "','" & Format(Now, "mm/dd/yy") & "'," & cAdjQty _
                   & "," & cAdjQty & "," & cStdCost & ",'" & sCreditAcct _
                   & "','" & sDebitAcct & "'," & lCOUNTER & ",'" & sInitials & "')"
            clsADOCon.ExecuteSQL sSql
         End If
      End If
      UpdateWipColumns lSysCount
      SysMsg "Part Number Was Successfully Updated.", True
      bGoodPart = GetPartInfo()
   Else
      MsgBox Err.Description
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MsgBox "The Part Number Could Not Be Updated.", _
         vbInformation, Caption
   End If
   optQoh.Enabled = True
   Exit Sub
   
DiaErr1:
   sProcName = "adjustquantity"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Table doesn't exist? Create it

Private Sub CreateEditTable()
   On Error Resume Next
   sSql = "DROP TABLE LorcTable"
   'clsADOCon.ExecuteSQL sSql
   'RdoCon.Execute sSql, rdExecDirect
   
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT LOTREC_ROW FROM LorcTable WHERE LOTREC_ROW=1"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum <> 0 Then
      Err.Clear
      clsADOCon.ADOErrNum = 0
      sSql = "CREATE TABLE LorcTable (" _
             & "LOTREC_ROW INT NULL DEFAULT(1)," _
             & "LOTREC_PARTREF CHAR(30) NULL DEFAULT('')," _
             & "LOTREC_DATE SMALLDATETIME NULL DEFAULT(GetDate())," _
             & "LOTREC_PAQOH SMALLMONEY NULL DEFAULT(0)," _
             & "LOTREC_PALOTQTYREMAINING SMALLMONEY NULL DEFAULT(0)," _
             & "LOTREC_SUMLOTS SMALLMONEY NULL DEFAULT(0)," _
             & "LOTREC_SUMACTIVITY SMALLMONEY NULL DEFAULT(0)," _
             & "LOTREC_PARTADJ CHAR(1) DEFAULT('N'))"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "CREATE UNIQUE CLUSTERED INDEX LotRecRow ON " _
                & "LorcTable(LOTREC_ROW) WITH FILLFACTOR = 80"
         clsADOCon.ExecuteSQL sSql
         
         sSql = "CREATE INDEX LotPartRef ON " _
                & "LorcTable((LOTREC_PARTREF) WITH FILLFACTOR = 80"
         clsADOCon.ExecuteSQL sSql
      End If
   End If
   
End Sub

Private Function GetNextRow() As Long
   Dim RdoRow As ADODB.Recordset
   sSql = "SELECT MAX(LOTREC_ROW)as LastRow FROM LorcTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRow, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoRow!LastRow) Then GetNextRow = RdoRow!LastRow _
                    Else GetNextRow = 0
   Else
      GetNextRow = 0
   End If
   GetNextRow = GetNextRow + 1
   Set RdoRow = Nothing
   
End Function

