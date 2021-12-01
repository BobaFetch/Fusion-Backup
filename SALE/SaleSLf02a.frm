VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form SaleSLf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy A Sales Order"
   ClientHeight    =   3600
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optTemplate 
      Caption         =   "Remember As Template"
      Height          =   252
      Left            =   3480
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Save This Sales Order As The Template For Future Sessions"
      Top             =   960
      Width           =   2292
   End
   Begin VB.ComboBox cmbSon 
      Height          =   288
      Left            =   2400
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Select or Enter Sales Order Number (Contains 300 Max)"
      Top             =   960
      Width           =   975
   End
   Begin VB.Timer tmr1 
      Interval        =   10000
      Left            =   5640
      Top             =   2160
   End
   Begin VB.CommandButton cmdCpy 
      Caption         =   "&Copy"
      Height          =   315
      Left            =   5880
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Copy Sales Order To The New Sales Order"
      Top             =   840
      Width           =   915
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      ToolTipText     =   "New Sales Order"
      Top             =   1680
      Width           =   735
   End
   Begin VB.CheckBox optCom 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox optRem 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.CheckBox optItm 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdDis 
      Caption         =   "&Display"
      Height          =   315
      Left            =   5880
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Display Sales Order"
      Top             =   480
      Width           =   915
   End
   Begin VB.TextBox txtOld 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Current Sales Order"
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   3480
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3600
      FormDesignWidth =   6870
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   2160
      TabIndex        =   21
      Top             =   3120
      Width           =   4452
      _ExtentX        =   7858
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblCst 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label lblNew 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   15
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To Sales Order"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy:"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Commission Information"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Remarks"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Comments"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy Sales Order Number"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   2025
   End
   Begin VB.Label lblOld 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   960
      Width           =   255
   End
End
Attribute VB_Name = "SaleSLf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'6/18/04 Correct loop to copy All Items
'10/7/05 Corrected Column Not Found and SOTEXT
'10/14/05 Added So Combo and Remember Option
'10/2/06 See CopySalesOrder - Changed to SQL Commands (revamped)
Option Explicit
Dim bNewExists As Byte
Dim bOldExists As Byte
Dim bOnLoad As Byte

Dim sTemplate As String

Private txtKeyPress() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd


Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim iList As Integer
   Dim sYear As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   cmbSon.Clear
   iList = Format(Now, "yyyy")
   iList = iList - 2
   sYear = Trim$(iList) & "-" & Format(Now, "mm-dd")
   sSql = "Qry_FillSalesOrders '" & sYear & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      iList = -1
      With RdoCmb
         lblOld = "" & Trim(!SOTYPE)
         cmbSon = Format(!SoNumber, SO_NUM_FORMAT)
         Do Until .EOF
            iList = iList + 1
            If iList > 999 Then Exit Do
            AddComboStr cmbSon.hWnd, Format$(!SoNumber, SO_NUM_FORMAT)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   Else
      MouseCursor 0
      MsgBox "No Sales Orders Where Found.", vbInformation, Caption
      Exit Sub
   End If
   Set RdoCmb = Nothing
   MouseCursor 0
   If sTemplate <> "" Then cmbSon = sTemplate
   txtOld = cmbSon
   If cmbSon.ListCount > 0 Then bOldExists = GetSalesOrder(txtOld, False)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbSon_Click()
   txtOld = cmbSon
   bOldExists = GetSalesOrder(txtOld, False)
   
End Sub

Private Sub cmbSon_LostFocus()
   cmbSon = CheckLen(cmbSon, SO_NUM_SIZE)
   cmbSon = Format(Abs(Val(cmbSon)), SO_NUM_FORMAT)
   txtOld = cmbSon
   If Val(txtOld) > 0 Then
      bOldExists = GetSalesOrder(txtOld, False)
   Else
      bOldExists = False
      lblCst = ""
      lblNme = ""
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCpy_Click()
   tmr1.Enabled = False
   If bOldExists And Not bNewExists Then
      CopySalesOrder
   Else
      MsgBox "Requires A Valid Current And New Sales Order.", vbInformation, Caption
      On Error Resume Next
      txtOld.SetFocus
   End If
   
End Sub

Private Sub cmdDis_Click()
   If bOldExists Then
      PrintReport
   Else
      MsgBox "Requires A Valid Sales Order.", vbInformation, Caption
   End If
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2152
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   MouseCursor 0
   If bOnLoad Then
      GetLastSalesOrder
      FillCombo
      sJournalID = GetOpenJournal("SJ", Format(ES_SYSDATE, "mm/dd/yy"))
      If Left(sJournalID, 4) = "None" Then
         sJournalID = ""
         b = 1
      Else
         If sJournalID = "" Then b = 0 Else b = 1
      End If
      
      If b = 0 Then
         MsgBox "There Is No Open Sales Journal For This Period.", _
            vbExclamation, Caption
         Sleep 500
         Unload Me
         Exit Sub
      End If
      bOnLoad = 0
   End If
   
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
   Set SaleSLf02a = Nothing
   
End Sub





Private Sub optCom_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optItm_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optRem_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub tmr1_Timer()
   GetLastSalesOrder
   
End Sub

Private Sub txtNew_LostFocus()
   txtNew = CheckLen(txtNew, SO_NUM_SIZE)
   txtNew = Format(Abs(Val(txtNew)), SO_NUM_FORMAT)
   tmr1.Enabled = False
   bNewExists = GetSalesOrder(txtNew, True)
   
End Sub

Private Sub txtOld_Click()
   tmr1.Enabled = True
   
End Sub

Private Sub txtOld_GotFocus()
   txtOld_Click
   
End Sub


Function GetSalesOrder(lSalesOrder As Variant, bNew As Byte) As Byte
   Dim RdoSon As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT SONUMBER,SOTYPE,SOCUST FROM SohdTable " _
          & "WHERE SONUMBER=" & Trim(str(lSalesOrder)) & " "
   If Not bNew Then sSql = sSql & "AND SOCANCELED=0 "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon)
   If bSqlRows Then
      With RdoSon
         lblOld = "" & Trim(!SOTYPE)
         lblNew = "" & Trim(!SOTYPE)
         If bNew Then txtOld = Format(!SoNumber, SO_NUM_FORMAT)
         FindCustomer Me, !SOCUST, False
         ClearResultSet RdoSon
         lblCst.Alignment = 0
         GetSalesOrder = True
      End With
   Else
      If Not bNew Then
         Beep
         lblCst.Alignment = 1
         lblOld = ""
         lblNew = ""
         lblCst = "****No Such "
         lblNme = "Sales Order Or Sales Order Canceled****"
      End If
      GetSalesOrder = False
   End If
   If bNew And GetSalesOrder Then
      MsgBox "Sales Order Number Is In Use.", vbInformation, Caption
      GetLastSalesOrder
      tmr1.Enabled = True
      If Val(txtNew) > 0 Then GetSalesOrder = True Else GetSalesOrder = False
   End If
   On Error Resume Next
   Set RdoSon = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getsaleso"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub PrintReport()
   Dim sSoNumber As String
   MouseCursor 13
   
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   sSoNumber = Trim(str(Val(txtOld)))
   sCustomReport = GetCustomReport("sleco01")
   
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath
    cCRViewer.SetReportTitle = sCustomReport
   
   sSql = "{SohdTable.SONUMBER}=" & Val(txtOld) & " "
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
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetLastSalesOrder()
   Dim RdoSon As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT SONUMBER FROM SohdTable ORDER BY SONUMBER DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon)
   If bSqlRows Then
      With RdoSon
         txtNew = Format$(!SoNumber + 1, SO_NUM_FORMAT)
         ClearResultSet RdoSon
      End With
   Else
      txtNew = SO_NUM_FORMAT
   End If
   Set RdoSon = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getlastso"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtOld_LostFocus()
   txtOld = CheckLen(txtOld, SO_NUM_SIZE)
   txtOld = Format(Abs(Val(txtOld)), SO_NUM_FORMAT)
   If Val(txtOld) > 0 Then
      bOldExists = GetSalesOrder(txtOld, False)
   Else
      bOldExists = False
      lblCst = ""
      lblNme = ""
   End If
   
End Sub


'' pre 8/8/2017
'Private Sub CopySalesOrder()
'   Dim RdoCpy As ADODB.Recordset
'   Dim iRow As Integer
'   Dim bResponse As Integer
'   Dim sComments As String
'   Dim sMsg As String
'
'   sMsg = "Are You Sure That You Wish To Copy " & vbCrLf _
'          & "Sales Order " & lblOld & txtOld & " To " & lblNew & txtNew & "?"
'   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
'   If bResponse = vbNo Then Exit Sub
'   MouseCursor 13
'   iRow = 10
'   prg1.Value = iRow
'   prg1.Visible = True
'   cmdCpy.Enabled = False
'   cmdDis.Enabled = False
'
'   On Error Resume Next
'   'sSql = "DROP TABLE #Soit" 'Just in case
'   sSql = "DropTempTableIfExists '#Soit'"
'   clsADOCon.ExecuteSql sSql 'rdExecDirect
'
'   'sSql = "DROP TABLE #Sohd" 'Just in case
'   sSql = "DropTempTableIfExists '#Sohd'"
'   clsADOCon.ExecuteSql sSql 'rdExecDirect
'
'   On Error GoTo DiaErr1
'   sSql = "SELECT SONUMBER FROM SohdTable WHERE SONUMBER=" & Val(txtOld) & " "
'   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCpy, ES_FORWARD)
'   If bSqlRows Then
'      Err.Clear
'      On Error Resume Next
'      clsADOCon.BeginTrans
'      clsADOCon.ADOErrNum = 0
'
'      With RdoCpy
'         'Sales Order Header
'         'Changed to all SQL (removed static Cursor) commands 10/2/06
'         sSql = "SELECT * INTO #Sohd from SohdTable where SONUMBER=" & Val(txtOld) & " "
'         clsADOCon.ExecuteSql sSql 'rdExecDirect
'
'         'Remarks?
''         If optRem.Value = vbChecked Then
''            sSql = "UPDATE #Sohd SET SONUMBER=" & Val(txtNew) & ",SOTEXT='" _
''                   & txtNew & "',SODATE='" & Format(Now, "mm/dd/yy") & "'"
''            clsADOCon.ExecuteSql sSql 'rdExecDirect
''         Else
''            sSql = "UPDATE #Sohd SET SONUMBER=" & Val(txtNew) & ",SOTEXT='" _
''                   & txtNew & "',SODATE='" & Format(Now, "mm/dd/yy") & "',SOREMARKS=''"
''            clsADOCon.ExecuteSql sSql 'rdExecDirect
''         End If
'
'         If optRem.Value = vbChecked Then
'            sSql = "UPDATE #Sohd SET SONUMBER=" & Val(txtNew) _
'                   & ",SODATE='" & Format(Now, "mm/dd/yy") & "'"
'            clsADOCon.ExecuteSql sSql 'rdExecDirect
'         Else
'            sSql = "UPDATE #Sohd SET SONUMBER=" & Val(txtNew) _
'                  & ",SODATE='" & Format(Now, "mm/dd/yy") & "',SOREMARKS=''"
'            clsADOCon.ExecuteSql sSql 'rdExecDirect
'         End If
'
'         'Commissions?
'         If optCom.Value = vbUnchecked Then
'            sSql = "UPDATE #Sohd SET SOCYN=0,SOCOMMISSION=0"
'            clsADOCon.ExecuteSql sSql 'rdExecDirect
'         End If
'         sSql = "INSERT INTO SohdTable SELECT * FROM #Sohd "
'         clsADOCon.ExecuteSql sSql 'rdExecDirect
'
'         iRow = 30
'         prg1.Value = iRow
'         If clsADOCon.RowsAffected = 0 Then
'            MouseCursor 0
'            MsgBox "Couldn't Finish Sales Order Header." & vbCrLf _
'               & "Operation Terminated.", vbExclamation, Caption
'            On Error Resume Next
'            clsADOCon.RollbackTrans
'            cmdCpy.Enabled = True
'            cmdDis.Enabled = True
'            prg1.Visible = False
'            Exit Sub
'         End If
'      End With
'
'      sSql = "SELECT * INTO #Soit from SoitTable where ITSO=" & Val(txtOld) & " " _
'             & "AND ITCANCELED=0 "
'      clsADOCon.ExecuteSql sSql 'rdExecDirect
'
'      sSql = "UPDATE #Soit SET ITSO=" & Val(txtNew) & ",ITPSNUMBER='',ITPSITEM=0," _
'             & "ITBOOKDATE='" & Format(Now, "mm/dd/yy") & "',ITCREATED='" & Format(Now, "mm/dd/yy") _
'             & "',ITINVOICE=0,ITPSSHIPPED=0,ITACTUAL=NULL"
'      clsADOCon.ExecuteSql sSql 'rdExecDirect
'
'      If optItm.Value = vbUnchecked Then
'         sSql = "UPDATE #Soit SET ITCOMMENTS=''"
'         clsADOCon.ExecuteSql sSql 'rdExecDirect
'      End If
'
'      If optCom.Value = vbUnchecked Then
'         sSql = "UPDATE #Soit SET ITCOMMISSION=0"
'         clsADOCon.ExecuteSql sSql 'rdExecDirect
'      End If
'      sSql = "INSERT INTO SoitTable SELECT * FROM #Soit "
'      clsADOCon.ExecuteSql sSql 'rdExecDirect
'
'      prg1.Value = 100
'      Dim strPre As String
'
'      If clsADOCon.ADOErrNum = 0 Then
'         clsADOCon.CommitTrans
'         strPre = lblNew
'         ResetLastSalesOrderNumber strPre
'         MsgBox "Sales Order Copied.", _
'            vbInformation, Caption
'      Else
'         clsADOCon.RollbackTrans
'         MsgBox "Could Not Copy The Sales Order.", _
'            vbInformation, Caption
'      End If
'   End If
'   MouseCursor 0
'   Set RdoCpy = Nothing
'   cmdCpy.Enabled = True
'   cmdDis.Enabled = True
'   prg1.Visible = False
'   GetLastSalesOrder
'   On Error Resume Next
'   txtOld.SetFocus
'   Exit Sub
'
'DiaErr1:
'   Resume DiaErr2
'DiaErr2:
'   On Error Resume Next
'   clsADOCon.RollbackTrans
''   sSql = "DROP TABLE #Soit"
'   sSql = "DropTempTableIfExists '#Soit'"
'
'   clsADOCon.ExecuteSql sSql 'rdExecDirect
'
'   MouseCursor 0
'   prg1.Visible = False
'   MsgBox "Couldn't Copy The Sales Order.", vbExclamation, Caption
'
'End Sub
'

' post 8/8/2017 with computed SOTEXT column
Private Sub CopySalesOrder()
   Dim RdoCpy As ADODB.Recordset
   Dim iRow As Integer
   Dim bResponse As Integer
   Dim sComments As String
   Dim sMsg As String
   
   sMsg = "Are You Sure That You Wish To Copy " & vbCrLf _
          & "Sales Order " & lblOld & txtOld & " To " & lblNew & txtNew & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then Exit Sub
   MouseCursor 13
   iRow = 10
   prg1.Value = iRow
   prg1.Visible = True
   cmdCpy.Enabled = False
   cmdDis.Enabled = False
   
   On Error Resume Next
   'sSql = "DROP TABLE #Soit" 'Just in case
   sSql = "DropTempTableIfExists '#Soit'"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   'sSql = "DROP TABLE #Sohd" 'Just in case
   sSql = "DropTempTableIfExists '#Sohd'"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   
   'get column list excluding SOTEXT computed column
   Dim soCols As String
   soCols = ""
   sSql = "select column_name from INFORMATION_SCHEMA.COLUMNS" & vbCrLf _
      & "where TABLE_NAME = 'SohdTable' and COLUMN_NAME <> 'SOTEXT'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCpy, ES_FORWARD)
   If bSqlRows Then
         With RdoCpy
            Do Until .EOF
               If soCols <> "" Then soCols = soCols & ","
               soCols = soCols & !COLUMN_NAME
               .MoveNext
            Loop
         End With
   Else
      Exit Sub
   End If
   
   On Error GoTo DiaErr1
   sSql = "SELECT SONUMBER FROM SohdTable WHERE SONUMBER=" & Val(txtOld) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCpy, ES_FORWARD)
   If bSqlRows Then
      Err.Clear
      
      
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      With RdoCpy
         'Sales Order Header
         'Changed to all SQL (removed static Cursor) commands 10/2/06
         sSql = "SELECT " & soCols & " INTO #Sohd from SohdTable where SONUMBER=" & Val(txtOld) & " "
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         'Remarks?
'         If optRem.Value = vbChecked Then
'            sSql = "UPDATE #Sohd SET SONUMBER=" & Val(txtNew) & ",SOTEXT='" _
'                   & txtNew & "',SODATE='" & Format(Now, "mm/dd/yy") & "'"
'            clsADOCon.ExecuteSql sSql 'rdExecDirect
'         Else
'            sSql = "UPDATE #Sohd SET SONUMBER=" & Val(txtNew) & ",SOTEXT='" _
'                   & txtNew & "',SODATE='" & Format(Now, "mm/dd/yy") & "',SOREMARKS=''"
'            clsADOCon.ExecuteSql sSql 'rdExecDirect
'         End If

         If optRem.Value = vbChecked Then
            sSql = "UPDATE #Sohd SET SONUMBER=" & Val(txtNew) _
                   & ",SODATE='" & Format(Now, "mm/dd/yy") & "'"
            clsADOCon.ExecuteSql sSql 'rdExecDirect
         Else
            sSql = "UPDATE #Sohd SET SONUMBER=" & Val(txtNew) _
                  & ",SODATE='" & Format(Now, "mm/dd/yy") & "',SOREMARKS=''"
            clsADOCon.ExecuteSql sSql 'rdExecDirect
         End If

         'Commissions?
         If optCom.Value = vbUnchecked Then
            sSql = "UPDATE #Sohd SET SOCYN=0,SOCOMMISSION=0"
            clsADOCon.ExecuteSql sSql 'rdExecDirect
         End If
         sSql = "INSERT INTO SohdTable (" & soCols & ")" & vbCrLf _
            & "SELECT " & Replace(soCols, "SOREMARKS", "CAST(SOREMARKS as varchar(6000))") & " FROM #Sohd "
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         iRow = 30
         prg1.Value = iRow
         If clsADOCon.RowsAffected = 0 Then
            MouseCursor 0
            MsgBox "Couldn't Finish Sales Order Header." & vbCrLf _
               & "Operation Terminated.", vbExclamation, Caption
            On Error Resume Next
            clsADOCon.RollbackTrans
            cmdCpy.Enabled = True
            cmdDis.Enabled = True
            prg1.Visible = False
            Exit Sub
         End If
      End With
      
      sSql = "SELECT * INTO #Soit from SoitTable where ITSO=" & Val(txtOld) & " " _
             & "AND ITCANCELED=0 "
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      sSql = "UPDATE #Soit SET ITSO=" & Val(txtNew) & ",ITPSNUMBER='',ITPSITEM=0," _
             & "ITBOOKDATE='" & Format(Now, "mm/dd/yy") & "',ITCREATED='" & Format(Now, "mm/dd/yy") _
             & "',ITINVOICE=0,ITPSSHIPPED=0,ITACTUAL=NULL"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      If optItm.Value = vbUnchecked Then
         sSql = "UPDATE #Soit SET ITCOMMENTS=''"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
      End If
      
      If optCom.Value = vbUnchecked Then
         sSql = "UPDATE #Soit SET ITCOMMISSION=0"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
      End If
      sSql = "INSERT INTO SoitTable SELECT * FROM #Soit "
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      prg1.Value = 100
      Dim strPre As String
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         strPre = lblNew
         ResetLastSalesOrderNumber strPre
         MsgBox "Sales Order Copied.", _
            vbInformation, Caption
      Else
         clsADOCon.RollbackTrans
         MsgBox "Could Not Copy The Sales Order.", _
            vbInformation, Caption
      End If
   End If
   MouseCursor 0
   Set RdoCpy = Nothing
   cmdCpy.Enabled = True
   cmdDis.Enabled = True
   prg1.Visible = False
   GetLastSalesOrder
   On Error Resume Next
   txtOld.SetFocus
   Exit Sub
   
DiaErr1:
   Resume DiaErr2
DiaErr2:
   On Error Resume Next
   clsADOCon.RollbackTrans
'   sSql = "DROP TABLE #Soit"
   sSql = "DropTempTableIfExists '#Soit'"
   sSql = "DropTempTableIfExists '#Sohd'"

   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   MouseCursor 0
   prg1.Visible = False
   MsgBox "Couldn't Copy The Sales Order.", vbExclamation, Caption
   
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub GetOptions()
   sTemplate = Trim(GetSetting("Esi2000", "EsiSale", "SaleSLf02a", sTemplate))
   If sTemplate <> "" Then optTemplate.Value = vbChecked
   
End Sub

Private Sub SaveOptions()
   If optTemplate = vbChecked Then
      SaveSetting "Esi2000", "EsiSale", "SaleSLf02a", cmbSon
   Else
      SaveSetting "Esi2000", "EsiSale", "SaleSLf02a", ""
   End If
   
End Sub
