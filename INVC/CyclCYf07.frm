VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form CyclCYf07 
   Caption         =   "ABC Inventory Reconciliation"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12420
   ClipControls    =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   12420
   Begin VB.CheckBox chkDefaultQty 
      Caption         =   "Default to Locked Quantity"
      Height          =   195
      Left            =   1860
      TabIndex        =   3
      Top             =   900
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   5400
      TabIndex        =   16
      Text            =   "Used to edit grid"
      Top             =   540
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CyclCYf07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "Select"
      Height          =   360
      Left            =   11460
      TabIndex        =   4
      ToolTipText     =   "Fill The Form With Qualifying Items"
      Top             =   480
      Width           =   875
   End
   Begin VB.ComboBox cmbCid 
      Height          =   315
      Left            =   1860
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "List Includes Cycle ID's Not Locked Or Completed"
      Top             =   120
      Width           =   2115
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.ComboBox txtPlan 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6060
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "4"
      ToolTipText     =   "Planned Inventory Date"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   11460
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   60
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   60
      Top             =   360
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6915
      FormDesignWidth =   12420
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5595
      Left            =   60
      TabIndex        =   15
      Top             =   1260
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   9869
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lblCountsRequired 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8820
      TabIndex        =   20
      ToolTipText     =   "Total Items Included"
      Top             =   840
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Counts required"
      Height          =   255
      Index           =   1
      Left            =   7620
      TabIndex        =   19
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "No lots"
      Height          =   255
      Index           =   0
      Left            =   7620
      TabIndex        =   18
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblNoLots 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8820
      TabIndex        =   17
      ToolTipText     =   "Total Items Included"
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblCountsEntered 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8820
      TabIndex        =   13
      ToolTipText     =   "Total Items Included"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Counts entered"
      Height          =   255
      Index           =   4
      Left            =   7620
      TabIndex        =   12
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      Height          =   255
      Index           =   10
      Left            =   7620
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblTotalItems 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8820
      TabIndex        =   10
      ToolTipText     =   "Total Items Included"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cycle Count ID"
      Height          =   255
      Index           =   5
      Left            =   540
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   6
      Left            =   540
      TabIndex        =   8
      Top             =   525
      Width           =   1335
   End
   Begin VB.Label lblCabc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4140
      TabIndex        =   7
      ToolTipText     =   "ABC Code Selected"
      Top             =   120
      Width           =   405
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Planned Date"
      Height          =   255
      Index           =   7
      Left            =   4740
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "CyclCYf07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
Option Explicit
Dim bOnLoad As Byte
Private inGridEdit As Boolean

Private TotalItemsFromGrid As Integer
Private CountsEnteredFromGrid As Integer
Private CountsRequiredFromGrid As Integer
Private NoLotsFromGrid As Integer

Private Const LOTREQUIREDMSG As String = "LOT REQUIRED"


'Dim iTotalLots As Integer
'Dim iIndex As Integer
'Dim lCOUNTER As Long

Dim sCreditAcct As String
Dim sDebitAcct As String

Dim vNextDate As Variant
Private editingRow As Integer
Private editingCol As Integer
Private cancelIt As Boolean

'grid columns
Private Const COL_Location = 0
Private Const COL_PartRef = 1
Private Const COL_PartDescription = 2
Private Const COL_PartQty = 3
Private Const COL_UOM = 4
Private Const COL_PartCount = 5
Private Const COL_LotNo = 6
Private Const COL_UserLotNo = 7
Private Const COL_LotQty = 8
Private Const COL_LotCount = 9
Private Const COL_IsLotTracked = 10
Private Const COL_Count = 11      'number of columns

'grid cell colors
Private Const COLOR_NotEntered = &HC0C0FF          'red
Private Const COLOR_Entered = &HC0FFC0             'green
Private Const COLOR_NotEditable = &HE0E0E0         'grey
Private Const COLOR_ReadOnly = &HFFFFFF


'Dim sParts(1000, 5) As String 'Location,PartRef, Number, Description, Lot Tracked
'Dim cValue(1000, 5) As Currency 'Cost, Qoh (At Count), Actual Qty, Reconciled, PartTable.PAQOH
'Dim vCycleLots(100, 4) As Variant 'Lotnumber, Remaining, UserId, AdjustQty

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'1/17/04
'Larry's note
'For Debit Account use "Inventory Over/Short" account.
'For Credit Account use inventory/expense material account.
'loaded in part number,

Private Sub GetAccounts(PartNumber As String)
   Dim rdoAct As ADODB.Recordset
   Dim bType As Byte
   Dim sPcode As String
   
   On Error GoTo DiaErr1
   sDebitAcct = ""
   sCreditAcct = ""
   sSql = "SELECT COADJACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         If Not IsNull(!COADJACCT) Then _
                       sDebitAcct = "" & Trim(!COADJACCT)
         ClearResultSet rdoAct
      End With
   End If
   Set rdoAct = Nothing
   
   'Use current Part
   sSql = "Qry_GetExtPartAccounts '" & PartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         sPcode = "" & Trim(!PAPRODCODE)
         If sDebitAcct = "" Then sDebitAcct = "" & Trim(!PACGSMATACCT)
         sCreditAcct = "" & Trim(!PAINVEXPACCT)
         ClearResultSet rdoAct
         Set rdoAct = Nothing
      End With
   Else
      sCreditAcct = ""
      sDebitAcct = ""
      Exit Sub
   End If
   If sDebitAcct = "" Or sCreditAcct = "" Then
      'None in one or both there, try Product code
      sSql = "Qry_GetPCodeAccounts '" & sPcode & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            If sDebitAcct = "" Then sCreditAcct = "" & Trim(!PCCGSMATACCT)
            If sCreditAcct = "" Then sDebitAcct = "" & Trim(!PCINVMATACCT)
            ClearResultSet rdoAct
            Set rdoAct = Nothing
         End With
      End If
      If sDebitAcct = "" Or sCreditAcct = "" Then
         'Still none, we'll check the common
         sSql = "SELECT COCGSMATACCT" & Trim(str(bType)) & "," _
                & "COINVMATACCT" & Trim(str(bType)) & " " _
                & "FROM ComnTable WHERE COREF=1"
         bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
         If bSqlRows Then
            With rdoAct
               If sCreditAcct = "" Then sDebitAcct = "" & Trim(.Fields(0))
               If sDebitAcct = "" Then sCreditAcct = "" & Trim(.Fields(1))
               ClearResultSet rdoAct
               Set rdoAct = Nothing
            End With
         End If
      End If
   End If
   'After this excercise, we'll give up if none are found
   Set rdoAct = Nothing
   Exit Sub
   
DiaErr1:
   'Just bail for now. May not have anything set
   'CurrError.Number = Err
   'CurrError.Description = Err.Description
   'DoModuleErrors Me
   On Error GoTo 0
   
End Sub

Private Sub FillList()
   Dim RdoAbc As ADODB.Recordset
   Dim bResponse As Byte
   Dim iReconciled As Integer
   Dim sMsg As String
   On Error GoTo DiaErr1
   MouseCursor ccHourglass
   
   Grid1.FixedRows = 0
   Grid1.Rows = 1
   
   'make sure all count sums are correct
   'precautionary during transition 8/1/08.  it can be removed eventually
   sSql = "update CcitTable" & vbCrLf _
      & "set CIACTUALQOH = (select sum(isnull(CLLOTADJUSTQTY,0))" & vbCrLf _
      & "from CcitTable it" & vbCrLf _
      & "join CcltTable lt on lt.CLREF = it.CIREF and lt.CLPARTREF = it.CIPARTREF" & vbCrLf _
      & "where it.CIREF = CcitTable.CIREF" & vbCrLf _
      & "and it.CIPARTREF = CcitTable.CIPARTREF)" & vbCrLf _
      & "from CcitTable it2 join CcitTable on it2.CIREF = CcitTable.CIREF" & vbCrLf _
      & "and it2.CIREF = '" & cmbCid & "'" & vbCrLf
   clsADOCon.ExecuteSql sSql

   lblTotalItems = "0"
   lblCountsEntered = "0"
   lblNoLots = "0"
      
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,PALOCATION,PAQOH,PAABC,PAUNITS," & vbCrLf _
      & "PASTDCOST,CIREF,CIPARTREF,CILOTTRACK,CIPAQOH," & vbCrLf _
      & "ISNULL(CIACTUALQOH,0) AS CIACTUALQOH," & vbCrLf _
      & "CLLOTNUMBER,CLLOTREMAININGQTY,CLLOTADJUSTQTY,LOTUSERLOTID," & vbCrLf _
      & "ISNULL(CLENTERED,0) AS CLENTERED" & vbCrLf _
      & "FROM PartTable" & vbCrLf _
      & "JOIN CcitTable ON PARTREF=CIPARTREF" & vbCrLf _
      & "LEFT JOIN CcltTable ON CLREF=CIREF AND CLPARTREF=CIPARTREF" & vbCrLf _
      & "LEFT JOIN LohdTable ON CLLOTNUMBER = LOTNUMBER" & vbCrLf _
      & "WHERE CIREF='" & cmbCid & "' " & vbCrLf
      
   'sSql = sSql & "ORDER BY PALOCATION,PARTREF,CLLOTNUMBER"
   sSql = sSql & "ORDER BY PARTREF,CLLOTNUMBER"  ' v 20.2
   
   Dim prevPart As String
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAbc, ES_FORWARD)
   If bSqlRows Then
      With RdoAbc
         Do Until .EOF
            
            Dim sItem As String
            
            'if same part, don't show it again
            If prevPart = Trim(!PartNum) Then
               sItem = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9)
            Else
               sItem = Trim(!PALOCATION) _
                  & Chr(9) & " " & Trim(!PartNum) _
                  & Chr(9) & " " & Trim(!PADESC) _
                  & Chr(9) & Format(!CIPAQOH, ES_QuantityDataFormat) _
                  & Chr(9) & !PAUNITS _
                  & Chr(9) & Format(!CIACTUALQOH, ES_QuantityDataFormat)
               
               prevPart = Trim(!PartNum)
            End If
            
            If Trim(!CLLOTNUMBER) = "" Then
               sItem = sItem & Chr(9) & Chr(9) & Chr(9)
            Else
               sItem = sItem & Chr(9) & Trim(!CLLOTNUMBER) _
                  & Chr(9) & Trim(!LOTUSERLOTID) _
                  & Chr(9) & Format(!CLLOTREMAININGQTY, ES_QuantityDataFormat)
            End If
            
            sItem = sItem & Chr(9) & IIf(!CLENTERED = 1, Format(!CLLOTADJUSTQTY, ES_QuantityDataFormat), " ")
            sItem = sItem & Chr(9) & !CILOTTRACK
            Grid1.AddItem sItem
'Debug.Print sItem
            If Grid1.FixedRows <> 1 Then
               Grid1.FixedRows = 1
            End If
            
            Grid1.row = Grid1.Rows - 1
            
            Grid1.Col = COL_LotCount
            If !CLENTERED = 1 Then
               Grid1.CellBackColor = COLOR_Entered
            ElseIf Trim(!CLLOTNUMBER) <> "" Then
               Grid1.CellBackColor = COLOR_NotEntered
            ElseIf !CILOTTRACK = 0 Then
               Grid1.CellBackColor = COLOR_NotEntered
            Else
               Grid1.CellBackColor = COLOR_NotEditable
            End If
            
            'if canceled, stop
            'DoEvents
            If cancelIt Then
               MouseCursor ccDefault
               Exit Sub
            End If
            
            .MoveNext
         Loop
         ClearResultSet RdoAbc
      End With
   End If
   
   'If iReconciled = TotalItemsFromGrid Then
   If AnalyzeCounts() Then
'      sMsg = "All counts have been entered. " & vbCrLf _
'             & "Mark this cycle count as reconciled (complete) now?"
'      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
'      If bResponse = vbYes Then
'         MarkReconciled
'         Exit Sub
'      End If
   End If
   
   If TotalItemsFromGrid > 0 Then
      vNextDate = GetNextDate()
      cmdSel.Enabled = False
   End If
   Set RdoAbc = Nothing
   
   MouseCursor ccDefault
   Exit Sub
   
DiaErr1:
   MouseCursor ccDefault
   sProcName = "FillList"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Private Function GetCycleCount() As Byte
'   Dim RdoCid As rdoResultset
'   On Error GoTo DiaErr1
'   CloseLotBoxes
'   lblCost(1) = ""
'   lblLoc(1) = ""
'   lblPart(1) = ""
'   lblDsc(1) = ""
'   lblLotRem(1) = ""
'   txtAQty = ""
'   cmdAdj.Enabled = False
'   sSql = "Qry_GetCycleCount '" & Trim(cmbCid) & "'"
'   bSqlRows = GetDataSet(RdoCid, ES_FORWARD)
'   If bSqlRows Then
'      With RdoCid
'         lblCabc = "" & Trim(!CCABCCODE)
'         txtDsc = "" & Trim(!CCDESC)
'         txtPlan = Format(!CCPLANDATE, "mm/dd/yy")
'         GetCycleCount = 1
'         cmdSel.Enabled = True
'         ClearResultSet RdoCid
'      End With
'   Else
'      GetCycleCount = 0
'      cmdSel.Enabled = False
'      Select Case MsgBox("That Count ID Wasn't Found, Is Locked, Or Is Not Saved.  Do you wish to cancel?", _
'         vbQuestion + vbYesNo, Caption)
'      Case vbYes
'         Exit Function
'      End Select
'   End If
'   Set RdoCid = Nothing
'   Exit Function
'
'DiaErr1:
'   GetCycleCount = 0
'   sProcName = "getcycleco"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Function

Private Sub GetCycleCount()
   Dim RdoCid As ADODB.Recordset
'   On Error GoTo DiaErr1
'   CloseLotBoxes
'   lblCost(1) = ""
'   lblLoc(1) = ""
'   lblPart(1) = ""
'   lblDsc(1) = ""
'   lblLotRem(1) = ""
'   txtAQty = ""
'   cmdAdj.Enabled = False
   sSql = "SELECT *" & vbCrLf _
      & "FROM CchdTable" & vbCrLf _
      & "WHERE CCREF = '" & cmbCid & "'" & vbCrLf _
      & "AND CCCOUNTLOCKED = 1"
   If clsADOCon.GetDataSet(sSql, RdoCid, ES_FORWARD) Then
      With RdoCid
         lblCabc = "" & Trim(!CCABCCODE)
         txtDsc = "" & Trim(!CCDESC)
         txtPlan = Format(!CCPLANDATE, "mm/dd/yy")
         'GetCycleCount = 1
         cmdSel.Enabled = True
'         ClearResultSet RdoCid
      End With
'   Else
'      GetCycleCount = 0
'      cmdSel.Enabled = False
'      Select Case MsgBox("That Count ID Wasn't Found, Is Locked, Or Is Not Saved.  Do you wish to cancel?", _
'         vbQuestion + vbYesNo, Caption)
'      Case vbYes
'         Exit Function
'      End Select
   End If
   Set RdoCid = Nothing
'   Exit Function
'
'DiaErr1:
'   GetCycleCount = 0
'   sProcName = "getcycleco"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
End Sub


Private Sub cmbCid_Click()
   GetCycleCount
End Sub

Private Sub cmbCid_LostFocus()
   GetCycleCount
End Sub

Private Sub cmdCancel_Click()
   cancelIt = True
   Sleep 1000
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5455"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdReconcile_Click()
   MsgBox "Reconciliation not functional yet"
End Sub

Private Sub cmdSel_Click()
   cmdSel.Enabled = False
   FillList
   cmdSel.Enabled = True
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      vNextDate = GetNextDate()
      GetOptions
   End If
   bOnLoad = 0
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   cancelIt = False
   
   With Grid1
      .RowHeightMin = 255
      '.FixedRows = 1
      .Rows = 1
      .FixedCols = 0
      .Cols = COL_Count
      
      .Col = COL_Location
      .Text = "Loc"
      .ColWidth(.Col) = 600
      
      .Col = COL_PartRef
      .Text = "Part"
      .ColWidth(.Col) = 1500
      
      .Col = COL_PartDescription
      .Text = "Description"
      .ColWidth(.Col) = 3000
      
      .Col = COL_PartQty
      .Text = "Part Qty"
      .ColWidth(.Col) = 900
      
      .Col = COL_UOM
      .Text = "UOM"
      .ColWidth(.Col) = 500
      
      .Col = COL_PartCount
      .Text = "Part Count"
      .ColWidth(.Col) = 900
      
      .Col = COL_LotNo
      .Text = "System Lot"
      .ColWidth(.Col) = 1500
      .ColAlignment(.Col) = flexAlignLeftCenter
      
      .Col = COL_UserLotNo
      .Text = "User Lot"
      .ColWidth(.Col) = 1500
      .ColAlignment(.Col) = flexAlignLeftCenter
      
      .Col = COL_LotQty
      .Text = "Lot Qty"
      .ColWidth(.Col) = 900
      
      .Col = COL_LotCount
      .Text = "Lot Count"
      .ColWidth(.Col) = 900
      
      .Col = COL_IsLotTracked
      .Text = "IsLotTracked"
      .ColWidth(.Col) = 0
   End With
   
End Sub

Private Sub Form_Resize()
   Refresh
   If Me.Width - 270 >= 11115 Then
      Grid1.Width = Me.Width - 270
   Else
      Grid1.Width = 11115
   End If
   Grid1.Height = Me.Height - 1800
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveOptions
   FormUnload
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDsc.BackColor = Es_FormBackColor
   txtPlan.BackColor = Es_FormBackColor
   txtPlan.ToolTipText = "Planned Inventory Date"
   'z1(13).ForeColor = ES_BLUE
   'CloseLotBoxes
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbCid.Clear
   sSql = "SELECT CCREF FROM CchdTable WHERE (CCCOUNTLOCKED=1 AND " _
          & "CCUPDATED=0)"
   LoadComboBox cmbCid, -1
   If cmbCid.ListCount > 0 Then
      If Trim(cmbCid) = "" Then cmbCid = cmbCid.List(0)
      'bGoodCount = GetCycleCount()
   Else
      MsgBox "There Are No Locked And Not Reconciled Counts Recorded.", _
         vbInformation, Caption
      Unload Me
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Grid1_Click()
   GridEdit 32
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then
      KeyCode = 0
   End If
End Sub

Private Sub optShow_Click()
   If cmdSel.Enabled = False Then FillList
   
End Sub


Private Function GetNextDate() As Variant
   Dim RdoDate As ADODB.Recordset
   Dim iFrequency As Integer
   Dim dDate As Date
   
   On Error Resume Next
   dDate = Format(txtPlan, "mm/dd/yy")
   sSql = "SELECT COABCROW,COABCCODE,COABCFREQUENCY " _
          & "FROM CabcTable WHERE COABCCODE='" & lblCabc & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDate, ES_FORWARD)
   If bSqlRows Then iFrequency = RdoDate!COABCFREQUENCY
   GetNextDate = Format(dDate + iFrequency, "mm/dd/yy")
   Set RdoDate = Nothing
   
End Function


'Private Sub MarkReconciled()
'   Dim rdoRec As rdoResultset
'   Dim bByte As Byte
'
'   sSql = "SELECT CIREF,CIPARTREF,CIRECONCILED FROM CcitTable WHERE " _
'          & "(CIREF='" & cmbCid & "' AND CIRECONCILED=0)"
'   bSqlRows = GetDataSet(rdoRec, ES_FORWARD)
'   If bSqlRows Then Exit Sub
'
'   sSql = "UPDATE CchdTable SET CCUPDATEDDATE='" & Format(ES_SYSDATE, "mm/dd/yy") _
'          & "',CCUPDATED=1 WHERE CCREF='" & cmbCid & "'"
'   RdoCon.Execute sSql, rdExecDirect
'   If Err = 0 Then MsgBox Trim(cmbCid) & " Has Been Reconciled.", _
'            vbInformation, Caption
'   FillCombo
'
'End Sub

Private Sub txtPlan_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub Grid1_DblClick()
   GridEdit 32
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
   GridEdit KeyAscii
End Sub

Sub GridEdit(KeyAscii As Integer, Optional GoToRow As Integer)
   'optional GoToRow = 1 to navitate down
   '                 = -1 to navigate up
   '                 = 0 to go to where the mouse is
   
   If inGridEdit Then
      Exit Sub
   End If
   inGridEdit = True
   
   'avoid bad parameters, that could result in an infinite loop
   If Grid1.Rows <= 1 Or GoToRow < -1 Or GoToRow > 1 Then
      inGridEdit = False
      Exit Sub
   End If
   
   If GoToRow <> 0 Then
      Dim found As Boolean
      found = False
'      Do While Grid1.Row > 1 And Grid1.Row < Grid1.Rows - 2
'         Grid1.Row = Grid1.Row + GoToRow
'         Grid1.Col = COL_LotCount
'         If Grid1.CellBackColor = COLOR_NotEntered Or Grid1.CellBackColor = COLOR_Entered Then
'            found = True
'            Exit Do
'         End If
'      Loop
      
      Do While Grid1.row + GoToRow >= 1 And Grid1.row + GoToRow < Grid1.Rows
         Grid1.row = Grid1.row + GoToRow
         Grid1.Col = COL_LotCount
         If Grid1.CellBackColor = COLOR_NotEntered Or Grid1.CellBackColor = COLOR_Entered Then
            found = True
            Exit Do
         End If
      Loop
      
      If Not found Then
         inGridEdit = False
         Exit Sub
      End If
   Else
      Grid1.Col = Grid1.MouseCol
      Grid1.row = Grid1.MouseRow
   End If
   
   If Grid1.row = 0 Then
      inGridEdit = False
      Exit Sub
   End If
   
   'ignore clicks outside of count column
   If Grid1.CellBackColor <> COLOR_Entered And Grid1.CellBackColor <> COLOR_NotEntered Then
      inGridEdit = False
      Exit Sub
   End If
   
   'use correct font
   Text1.FontName = Grid1.FontName
   Text1.FontSize = Grid1.FontSize
   Select Case KeyAscii
      Case 0 To 32
         If Grid1.CellBackColor = COLOR_NotEntered And Me.chkDefaultQty.Value = vbChecked Then
            If Grid1.TextMatrix(Grid1.row, COL_LotNo) <> "" Then
               Text1 = Grid1.TextMatrix(Grid1.row, COL_LotQty)
            Else
               Text1 = GetNonBlankGridCell(Grid1.row, COL_PartQty)
            End If
         Else
            Text1 = Grid1
         End If
         Text1.SelStart = 1000
      Case Else
         Text1 = Chr(KeyAscii)
         Text1.SelStart = 1
   End Select
   
   'position the edit box
'If editingRow > Grid1.Row Then
'MsgBox "backwards " & editingRow & " -> " & Grid1.Row
'End If
   editingRow = Grid1.row
   editingCol = Grid1.Col
   
   Text1.Left = Grid1.CellLeft + Grid1.Left
   Text1.Top = Grid1.CellTop + Grid1.Top
   Text1.Width = Grid1.CellWidth
   Text1.Height = Grid1.CellHeight
   Text1.Visible = True
   Text1.SetFocus
   inGridEdit = False
End Sub

Private Sub Grid1_GotFocus()
   If Text1.Visible Then
      If Grid1.row <> editingRow Or Grid1.Col <> editingCol Then
         Grid1.row = editingRow
         Grid1.Col = editingCol
      End If
      
      If IsNumeric(Text1) Then
         If Format(Text1, ES_QuantityDataFormat) <> Grid1 Then
            Grid1 = Format(Text1, ES_QuantityDataFormat)
            SetReconciledFlag True
         End If
      ElseIf Trim(Text1) = "" And Trim(Grid1) <> "" Then
         Grid1 = Text1
         SetReconciledFlag False
      End If
      Text1.Visible = False
   End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      'this case never happens - form is closed
      Case vbKeyEscape
         Text1.Visible = False
         Grid1.SetFocus
         KeyCode = 0
         DoEvents
      Case vbKeyReturn, vbKeyDown
         Grid1.SetFocus
         DoEvents
         If Grid1.row < Grid1.Rows - 1 Then
            GridEdit 32, 1
         Else
            'Grid1.Col = 0     'force leavecell event
            cmbCid.SetFocus
         End If
         KeyCode = 0
      Case vbKeyUp
         Grid1.SetFocus
         DoEvents
         If Grid1.row > Grid1.FixedRows Then
            GridEdit 32, -1
         Else
            cmbCid.SetFocus
         End If
         KeyCode = 0
   End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   'noise suppression
   If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub SetReconciledFlag(reconciled As Boolean)
   'reconciled = TRUE if item is being reconciled
   '           = FALSE if item is being unreconciled
   'enter this sub positioned on the cell to be updated
   
   'if the current row has a lot number update the corresponding lot
   Dim PartNo As String
   Dim lotNo As String
   Dim partCount As Currency, lotCount As Currency
   Dim dt As String
   
   If reconciled Then
      dt = "'" & Format(GetServerDateTime(), "mm/dd/yy hh:mm") & "'"
      If Grid1.CellBackColor <> COLOR_Entered Then
         Grid1.CellBackColor = COLOR_Entered
         CountsEnteredFromGrid = CountsEnteredFromGrid + 1
         lblCountsEntered = CountsEnteredFromGrid
         CountsRequiredFromGrid = CountsRequiredFromGrid - 1
         lblCountsRequired = CountsRequiredFromGrid
      End If
   Else
      dt = "null"
      If Grid1.CellBackColor <> COLOR_NotEntered Then
         Grid1.CellBackColor = COLOR_NotEntered
         CountsEnteredFromGrid = CountsEnteredFromGrid - 1
         lblCountsEntered = CountsEnteredFromGrid
         CountsRequiredFromGrid = CountsRequiredFromGrid + 1
         lblCountsRequired = CountsRequiredFromGrid
      End If
   End If
   
   'find row with part number
   Dim i As Integer
   For i = Grid1.row To 1 Step -1
      PartNo = Compress(Grid1.TextMatrix(i, COL_PartRef))
      If PartNo <> "" Then
         Exit For
      End If
   Next
   lotNo = Trim(Grid1.TextMatrix(Grid1.row, COL_LotNo))
   lotCount = CCur("0" & Grid1.TextMatrix(Grid1.row, COL_LotCount))
   
   clsADOCon.BeginTrans
   
   sSql = "update CcltTable" & vbCrLf _
      & "set CLENTERED = " & IIf(reconciled, 1, 0) & "," & vbCrLf _
      & "CLENTEREDDATE = " & dt & "," & vbCrLf _
      & "CLLOTADJUSTQTY = " & lotCount & vbCrLf _
      & "where CLREF = '" & cmbCid & "'" & vbCrLf _
      & "and CLPARTREF = '" & PartNo & "'" & vbCrLf _
      & "and CLLOTNUMBER = '" & lotNo & "'"
   clsADOCon.ExecuteSql sSql
   
   'now update the total quantity
   sSql = "update CcitTable" & vbCrLf _
      & "set CIACTUALQOH = (select sum(isnull(CLLOTADJUSTQTY,0))" & vbCrLf _
      & "from CcitTable it" & vbCrLf _
      & "join CcltTable lt on lt.CLREF = it.CIREF and lt.CLPARTREF = it.CIPARTREF" & vbCrLf _
      & "where it.CIREF = '" & cmbCid & "'" & vbCrLf _
      & "and it.CIPARTREF = '" & PartNo & "')" & vbCrLf _
      & "from CcitTable it2 join CcitTable on it2.CIREF = CcitTable.CIREF" & vbCrLf _
      & "and it2.CIPARTREF = CcitTable.CIPARTREF" & vbCrLf _
      & "and it2.CIREF = '" & cmbCid & "'" & vbCrLf _
      & "and it2.CIPARTREF = '" & PartNo & "'"
      
   clsADOCon.ExecuteSql sSql
   
   'now retrieve the new sum(qty) and store it in the appropriate grid cell
   sSql = "select CIACTUALQOH from CcitTable" & vbCrLf _
      & "where CIREF = '" & cmbCid & "'" & vbCrLf _
      & "and CIPARTREF = '" & PartNo & "'"
   Dim rdo As ADODB.Recordset
   If clsADOCon.GetDataSet(sSql, rdo) Then
      Grid1.TextMatrix(i, COL_PartCount) = Format(rdo.Fields(0), ES_QuantityDataFormat)
   End If
   DoEvents
   Set rdo = Nothing
   clsADOCon.CommitTrans
End Sub

Private Function AnalyzeCounts() As Boolean
   'returns True if ready to reconcile
   'gets information from the database, and compares to the grid
   
   TotalItemsFromGrid = Grid1.Rows - 1
   CountsEnteredFromGrid = 0
   NoLotsFromGrid = 0
   CountsRequiredFromGrid = 0
   
   Dim i As Integer

   For i = 1 To Grid1.Rows - 1
      Grid1.row = i
      
      Grid1.Col = COL_LotCount
      If Grid1.CellBackColor = COLOR_Entered Then
         CountsEnteredFromGrid = CountsEnteredFromGrid + 1
      ElseIf Grid1.CellBackColor = COLOR_NotEntered Then
         CountsRequiredFromGrid = CountsRequiredFromGrid + 1
      ElseIf Grid1.CellBackColor = COLOR_NotEditable Then
         NoLotsFromGrid = NoLotsFromGrid + 1
      End If
      
      Grid1.Col = COL_LotNo
      If Grid1.Text = LOTREQUIREDMSG Then
         NoLotsFromGrid = NoLotsFromGrid + 1
      End If

   Next

   lblTotalItems = TotalItemsFromGrid
   lblCountsEntered = CountsEnteredFromGrid
   lblNoLots = NoLotsFromGrid
   Me.lblCountsRequired = CountsRequiredFromGrid
   
   If TotalItemsFromGrid = CountsEnteredFromGrid And NoLotsFromGrid = 0 Then
      AnalyzeCounts = True
   Else
      AnalyzeCounts = False
   End If
End Function

Private Function GetNonBlankGridCell(StartRow As Integer, Col As Integer) As String
   'If the cell specified is nonblank, return it's text
   'Otherwise iterate upwards until a nonblank cell is found
   Dim i As Integer
   Dim contents As String
   For i = StartRow To 1 Step -1
      contents = Trim(Grid1.TextMatrix(i, Col))
      If contents <> "" Then
         Exit For
      End If
   Next
   GetNonBlankGridCell = contents

End Function

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Me.chkDefaultQty.Value & "0000000"
   SaveSetting "Esi2000", "EsiInvc", "CyclCYf07", sOptions
   
End Sub

Private Sub GetOptions()
   Dim iList As Integer
   Dim sOptions As String
   sOptions = GetSetting("Esi2000", "EsiInvc", "CyclCYf07", "00000000")
   Me.chkDefaultQty = CInt(Mid(sOptions, 1, 1))
End Sub




