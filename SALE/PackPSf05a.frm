VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PackPSf05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel A Packing Slip Item (Printed)"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSf05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox lblCst 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame z2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   6615
      Begin VB.ComboBox cmbItem 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Tag             =   "8"
         ToolTipText     =   "Select From List"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdItm 
         Caption         =   "&Apply"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5640
         TabIndex        =   3
         ToolTipText     =   "Cancel This Packing Slip Item"
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lblQty 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5640
         TabIndex        =   23
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblPart 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblSit 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4800
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblSon 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   240
         Width           =   675
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   255
         Index           =   8
         Left            =   4800
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "SO Item"
         Height          =   255
         Index           =   6
         Left            =   3720
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Order"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "PS Item"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbPsl 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Packslips Not Printed"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5760
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
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
      FormDesignHeight=   3240
      FormDesignWidth =   6720
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   3840
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   1104
   End
   Begin VB.Label lblPrinted 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3960
      TabIndex        =   21
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printed"
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   20
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slips Printed"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblDte 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3960
      TabIndex        =   9
      Top             =   720
      Width           =   990
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1320
      TabIndex        =   7
      Top             =   1440
      Width           =   3640
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "PackPSf05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'4/13/05 New
'5/2/05  Changed Lots and altered query (GetPackSlipLots)
'3/3/06  Added lblDate (hidden) to closer define lots to be adjusted
'        Also added the quantity for the test (GetPackSlipLots)
Option Explicit
'Dim rdoQry As rdoQuery
Dim cmdObj As ADODB.Command

Dim bCancel As Byte
Dim bGoodPs As Byte
Dim iLots As Integer
Dim bOnLoad As Byte

Dim sCreditAcct As String
Dim sDebitAcct As String

Dim sItems(300, 6) As String
Dim sLots(50, 2) As String
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetPackslip(Optional bGetItemsToo As Boolean) As Byte
   Dim RdoCpt As ADODB.Recordset
   On Error GoTo DiaErr1
   z2.Enabled = False
   cmbItem.Clear
   cmdItm.Enabled = False
   cmbItem.Enabled = False
   lblSon = ""
   lblSit = ""
   lblPart = ""
   lblQty = ""
   Erase sItems
'   rdoQry.RowsetSize = 1
'   rdoQry(0) = Compress(cmbPsl)
'   bSqlRows = GetQuerySet(RdoCpt, rdoQry, ES_FORWARD)
   
   cmdObj.Parameters(0).Value = Compress(cmbPsl)
   bSqlRows = clsADOCon.GetQuerySet(RdoCpt, cmdObj, ES_FORWARD, True)
      
   If bSqlRows Then
      With RdoCpt
         lblDte = "" & Format(!PSDATE, "mm/dd/yyyy")
         lblPrinted = "" & Format(!PSPRINTED, "mm/dd/yyyy")
         lblDate = "" & Format(!PSPRINTED, "mm/dd/yyyy hh:mm")
         If Trim(!PSCUST) <> "" Then FindCustomer Me, Trim(!PSCUST), True
      End With
      ClearResultSet RdoCpt
      GetPackslip = 1
      If bGetItemsToo Then GetItems
   Else
      lblDte = ""
      lblCst = ""
      lblNme = "*** Invalid Or Not Qualifying Packing Slip ***"
      GetPackslip = 0
   End If
   Set RdoCpt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpacksl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FillPackSlips()
   On Error GoTo DiaErr1
   cmbPsl.Clear
   sSql = "SELECT DISTINCT PSNUMBER,PSTYPE,PIPACKSLIP FROM " _
          & "PshdTable,PsitTable WHERE " _
          & "(PSPRINTED IS NOT NULL AND PSINVOICE=0) " _
          & "AND PSNUMBER=PIPACKSLIP ORDER BY PSNUMBER"
   LoadComboBox cmbPsl, -1
   If cmbPsl.ListCount > 0 Then
      cmdItm.Enabled = True
      cmbPsl = cmbPsl.List(0)
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillpacks"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbItem_Click()
   If cmbItem.ListCount > 0 Then
      lblSon = sItems(cmbItem.ListIndex, 1)
      lblSit = sItems(cmbItem.ListIndex, 2) & sItems(cmbItem.ListIndex, 3)
      lblPart = sItems(cmbItem.ListIndex, 4)
      lblQty = sItems(cmbItem.ListIndex, 5)
   End If
End Sub


Private Sub cmbPsl_Click()
   bGoodPs = GetPackslip()
   
End Sub


Private Sub cmbPsl_LostFocus()
   If bCancel = 0 Then
      cmbPsl = CheckLen(cmbPsl, 8)
      ' Not need to prepend "PS"
      'If Val(cmbPsl) > 0 Then cmbPsl = Format(cmbPsl, "########")
      bGoodPs = GetPackslip(True)
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2253
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdItm_Click()
   Dim b As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   sJournalID = GetOpenJournal("IJ", Format(ES_SYSDATE, "mm/dd/yyyy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   If b = 0 Then
      MsgBox "There Is No Open Inventory Journal For This" & vbCrLf _
         & "Period. Cannot Set The Pack Slip As Not Printed.", _
         vbExclamation, Caption
      Exit Sub
   End If
   If lblNme.ForeColor <> ES_RED Then
      sMsg = "Are You Sure That You Want To Cancel The " & vbCrLf _
             & "Printing And Delete Item " & cmbItem & " From " & cmbPsl & "?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         CancelThisItem
      Else
         CancelTrans
      End If
   Else
      MsgBox "Requires A Valid Packing Slip.", _
         vbInformation, Caption
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillPackSlips
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
 
   sSql = "SELECT PSNUMBER,PSCUST,PSDATE,PSPRINTED,PSSHIPPED,PSINVOICE FROM " _
          & "PshdTable WHERE PSNUMBER= ? AND (PSPRINTED IS NOT NULL AND PSINVOICE=0)"
 '  Set rdoQry = RdoCon.CreateQuery("", sSql)
 '  rdoQry.MaxRows = 1
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql
   
   Dim prmObj As ADODB.Parameter
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adChar
   prmObj.Size = 8
   
   cmdObj.Parameters.Append prmObj
  
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set cmdObj = Nothing
   Set PackPSf05a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblCst.BackColor = Es_FormBackColor
   lblNme.ForeColor = vbBlack
   
End Sub

Private Sub lblNme_Change()
   If Left(lblNme, 5) = "*** I" Then
      lblNme.ForeColor = ES_RED
   Else
      lblNme.ForeColor = Es_TextForeColor
   End If
   
End Sub

Private Sub GetItems()
   Dim RdoItm As ADODB.Recordset
   Dim iList As Integer
   iList = -1
   On Error GoTo DiaErr1
   sSql = "SELECT PIPACKSLIP,PIITNO,PIQTY,PIPART," _
          & "PISONUMBER,PISOITEM,PISOREV,PARTREF,PARTNUM " _
          & "FROM PsitTable,PartTable where (PIPART=PARTREF AND " _
          & "PIPACKSLIP='" & Trim(cmbPsl) & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_FORWARD)
   If bSqlRows Then
      With RdoItm
         Do Until .EOF
            iList = iList + 1
            cmbItem.AddItem str$(!PIITNO)
            sItems(iList, 0) = str$(!PIITNO)
            sItems(iList, 1) = Format$(!PISONUMBER, SO_NUM_FORMAT)
            sItems(iList, 2) = Format$(!PISOITEM, "##0")
            sItems(iList, 3) = "" & Trim(!PISOREV)
            sItems(iList, 4) = "" & Trim(!PARTNUM)
            sItems(iList, 5) = Format$(!PIQTY, ES_QuantityDataFormat)
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
   End If
   If cmbItem.ListCount > 0 Then
      z2.Enabled = True
      cmbItem.Enabled = True
      cmdItm.Enabled = True
      cmbItem = cmbItem.List(0)
      cmbItem.ListIndex = 0
      lblSon = sItems(cmbItem.ListIndex, 1)
      lblSit = sItems(cmbItem.ListIndex, 2) & sItems(cmbItem.ListIndex, 3)
      lblPart = sItems(cmbItem.ListIndex, 4)
      lblQty = sItems(cmbItem.ListIndex, 5)
      cmbItem.SetFocus
   Else
      lblPart = "*** There Are No Items On This Packing Slip ****"
   End If
   Set RdoItm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub CancelThisItem()
   Dim bByte As Byte
   Dim iList As Integer
   Dim lCOUNTER As Long
   Dim lSysCount As Long
   Dim lLOTRECORD As Long
   Dim cPartCost As Currency
   Dim vDate As Variant
   
   'On Error Resume Next
   Err.Clear
   On Error GoTo whoops
   
   MouseCursor 13
   cmdCan.Enabled = False
   cmdItm.Enabled = False
   bByte = GetPartAccounts(Compress(lblPart), sCreditAcct, sDebitAcct)
   vDate = Format(ES_SYSDATE, "mm/dd/yyyy hh:mm")
   
   lCOUNTER = (GetLastActivity) + 1
   lSysCount = lCOUNTER
   iLots = GetPackSlipLots()
   Err.Clear
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   'take it out of the PS and Sales Order
   sSql = "DELETE FROM PsitTable WHERE (PIPACKSLIP='" _
          & Trim(cmbPsl) & "' AND PIITNO=" & Val(cmbItem) & ")"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "UPDATE SoitTable SET ITACTUAL=NULL, ITPSNUMBER=''," _
          & "ITPSITEM=0, ITPSSHIPPED=0 WHERE (ITPSNUMBER='" & Trim(cmbPsl) & "' " _
          & "AND ITPSITEM=" & Val(cmbItem) & ")"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   'Make the adjustments
   cPartCost = GetPartCost(Compress(lblPart))
   If iLots > 0 Then
      For iList = 1 To iLots
         lCOUNTER = lCOUNTER + 1
         lLOTRECORD = GetNextLotRecord(sLots(iList, 0))
         sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                & "LOITYPE,LOIPARTREF,LOIQUANTITY,LOIADATE," _
                & "LOIPSNUMBER,LOIPSITEM,LOICUST,LOIACTIVITY,LOICOMMENT) " _
                & "VALUES('" & sLots(iList, 0) & "'," & lLOTRECORD & "," & IATYPE_CancPackSlip & ",'" _
                & Compress(lblPart) & "'," & Val(sLots(iList, 1)) & ",'" _
                & vDate & "','" & Trim(cmbPsl) & "'," & Val(cmbItem) & ",'" _
                & Compress(lblCst) & "'," & lCOUNTER & ",'Canceled PS Print')"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         
         sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY+" _
                & Val(sLots(iList, 1)) & " WHERE LOTNUMBER='" & sLots(iList, 0) & "'"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         
         'Activity
         sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                & "INPDATE,INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," _
                & "INPSNUMBER,INPSITEM,INNUMBER,INLOTNUMBER,INUSER) " _
                & "VALUES(" & IATYPE_CancPackSlip & ",'" & Compress(lblPart) & "','CANCELED PS PRINT'," _
                & "'" & cmbPsl & "','" & vDate & "','" & vDate & "'," _
                & Val(sLots(iList, 1)) & "," & Val(sLots(iList, 1)) & "," & cPartCost & ",'" & sCreditAcct _
                & "','" & sDebitAcct & "','" & Trim(cmbPsl) & "'," & Val(cmbItem) _
                & "," & lCOUNTER & ",'" & sLots(iList, 0) & "','" & sInitials & "')"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
      Next
   Else
      'Activity (No lots found)
      lCOUNTER = lCOUNTER + 1
      sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
             & "INPDATE,INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," _
             & "INPSNUMBER,INPSITEM,INNUMBER,INUSER) " _
             & "VALUES(" & IATYPE_CancPackSlip & ",'" & Compress(lblPart) & "','CANCELED PS PRINT'," _
             & "'" & cmbPsl & "','" & vDate & "','" & vDate & "'," & Val(lblQty) & "," & Val(lblQty) & "," _
             & cPartCost & ",'" & sCreditAcct & "','" & sDebitAcct & "','" _
             & Trim(cmbPsl) & "'," & Val(cmbItem) & "," & lCOUNTER & ",'" _
             & sInitials & "')"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
   End If
   
   'Update Part
   sSql = "UPDATE PartTable SET PAQOH=PAQOH+" & Val(lblQty) & "," _
          & "PALOTQTYREMAINING=PALOTQTYREMAINING+" & Val(lblQty) & " " _
          & "WHERE PARTREF='" & Compress(lblPart) & "' "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   AverageCost Compress(lblPart)
   
   Dim ia As New ClassInventoryActivity
   ia.UpdatePackingSlipCosts (Trim(cmbPsl))
   
   MouseCursor 0
   cmdCan.Enabled = True
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      MsgBox "The Print Was Successfully Canceled And The Item Deleted.", _
         vbInformation, Caption
      UpdateWipColumns lSysCount
      FillPackSlips
   Else
      clsADOCon.RollbackTrans
      MsgBox "Couldn't Cancel The Selected Item. Transaction Rolled Back.", _
         vbExclamation, Caption
   End If
   Exit Sub
   
whoops:
   MouseCursor ccHourglass
   clsADOCon.RollbackTrans
   sProcName = "CancelThisItem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Added Order By and duplicate protection 3/3/06

Private Function GetPackSlipLots() As Integer
   Dim RdoLots As ADODB.Recordset
   Dim iRows As Integer
   Dim sLoiLot As String
   Erase sLots
   sSql = "SELECT LOINUMBER,LOIQUANTITY,LOIPSNUMBER,LOIPSITEM,LOIRECORD FROM " _
          & "LoitTable WHERE (LOIPSNUMBER='" & Trim(cmbPsl) & "' " _
          & "AND LOIQUANTITY<=" & lblQty & " AND LOIPSITEM=" & Val(cmbItem) _
          & " AND LOITYPE=25 AND LOIADATE>='" & Format(lblDate, "mm/dd/yyyy") & "') ORDER BY " _
          & "LOINUMBER,LOIRECORD DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
   If bSqlRows Then
      With RdoLots
         Do Until .EOF
            If Trim(!LOINUMBER) <> sLoiLot Then
               iRows = iRows + 1
               sLots(iRows, 0) = "" & Trim(!LOINUMBER)
               sLots(iRows, 1) = "" & str$(Abs(!LOIQUANTITY))
            End If
            sLoiLot = Trim(!LOINUMBER)
            .MoveNext
         Loop
         ClearResultSet RdoLots
      End With
      GetPackSlipLots = iRows
   Else
      GetPackSlipLots = 0
   End If
   
End Function
