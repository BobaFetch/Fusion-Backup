VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form PackPSf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel A Packing Slip Printing"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optShipped 
      Caption         =   "  "
      Enabled         =   0   'False
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdItm 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5280
      TabIndex        =   7
      ToolTipText     =   "Cancel Packing Slip Printing"
      Top             =   600
      Width           =   915
   End
   Begin VB.ComboBox cmbPsl 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Contains Printed Packslips Not Printed"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   2040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2145
      FormDesignWidth =   6240
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shipped"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printed"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Width           =   3560
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3960
      TabIndex        =   2
      Top             =   1080
      Width           =   1035
   End
End
Attribute VB_Name = "PackPSf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'9/4/04 Added sPartNumber Array
'10/7/04 Revamped GetLots
Option Explicit
'Dim rdoQry As rdoQuery
Dim cmdObj As ADODB.Command
Dim bOnLoad As Byte
Dim bGoodPs As Byte

Dim iTotalItems As Integer

Dim sCreditAcct As String
Dim sDebitAcct As String

Dim sLots(50, 3) As String 'See GetLots
Dim sPartGroup(800) As String
Dim vItems(800, 7) As Variant
'   0 = PIITNO
'   1 = PITYPE
'   2 = PIQTY
'   3 = PIPART
'   4 = PISONUMBER
'   5 = PISOITEM
'   6 = PISOREV

Dim sTimeStamp(300) As String
'   =PILOTNUMBER (time stamp)

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbPsl_Click()
   bGoodPs = GetPackslip()
   
End Sub


Private Sub cmbPsl_LostFocus()
   cmbPsl = CheckLen(cmbPsl, 8)
   ' Not need to prepend "PS"
   'If Val(cmbPsl) > 0 Then cmbPsl = Format(cmbPsl, "00000000")
   bGoodPs = GetPackslip()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2251
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdItm_Click()
   cmdItm.Enabled = False
   Dim b As Byte
   If bGoodPs Then
      b = GetItems()
      CancelPrint
   Else
      MsgBox "Requires A Valid Packing Slip.", _
         vbExclamation, Caption
   End If
   cmdItm.Enabled = True
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then FillPackSlips
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   
   sSql = "SELECT PSNUMBER,PSCUST,PSPRINTED,PSSHIPPRINT,PSSHIPPED FROM " _
          & "PshdTable WHERE PSNUMBER= ? AND (PSSHIPPRINT=1 AND PSINVOICE=0)"
'  Set rdoQry = RdoCon.CreateQuery("", sSql)
'   rdoQry.MaxRows = 1
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql

   Dim prmObj As ADODB.Parameter
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adChar
   prmObj.Size = 8
   cmdObj.parameters.Append prmObj
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set cmdObj = Nothing
   Set PackPSf02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Function GetPackslip() As Byte
   Dim RdoCap As ADODB.Recordset
   On Error GoTo DiaErr1
'   rdoQry.RowsetSize = 1
'   rdoQry(0) = Compress(cmbPsl)
'   bSqlRows = GetQuerySet(RdoCap, rdoQry, ES_KEYSET)
   cmdObj.parameters(0).Value = Compress(cmbPsl)
   bSqlRows = clsADOCon.GetQuerySet(RdoCap, cmdObj, ES_FORWARD, True)
   
   If bSqlRows Then
      With RdoCap
         lblDte = "" & Format(!PSPRINTED, "mm/dd/yyyy")
         If Trim(!PSCUST) <> "" Then FindCustomer Me, Trim(!PSCUST), True
         optShipped.Value = !PSSHIPPED
      End With
      ClearResultSet RdoCap
      GetPackslip = 1
   Else
      lblDte = ""
      lblCst = ""
      lblNme = "*** Invalid Packing Slip ***"
      GetPackslip = 0
   End If
   Set RdoCap = Nothing
   Exit Function
   
DiaErr1:
   GetPackslip = 0
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
          & "(PSSHIPPRINT=1 AND PSINVOICE=0) " _
          & "AND PSNUMBER=PIPACKSLIP"
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


Private Sub lblNme_Change()
   If Left(lblNme, 5) = "*** I" Then
      lblNme.ForeColor = ES_RED
      cmdItm.Enabled = False
   Else
      lblNme.ForeColor = Es_TextForeColor
      cmdItm.Enabled = True
   End If
   
End Sub


Private Sub CancelPrint()
   Dim bResponse As Byte
   Dim bLotRec As Byte
   Dim b As Byte
   
   Dim iLotList As Integer
   Dim iRow As Integer
   Dim iItem As Integer
   Dim lCOUNTER As Long
   Dim lLOTRECORD As Long
   Dim lSysCount As Long
   
   Dim cItmLot As Currency
   Dim cPartCost As Currency
   Dim cQuantity As Currency
   Dim sMsg As String
   Dim sPackSlip As String
   
   Dim vDate As Variant
   
   vDate = Format(GetServerDateTime(), "mm/dd/yyyy hh:mm")
   On Error GoTo DiaErr1
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
   
   sMsg = "Do You Really Want To Cancel The " & vbCrLf _
          & "Printing Of Packing Slip " & cmbPsl & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      cmdItm.Enabled = False
      lCOUNTER = (GetLastActivity)
      lSysCount = lCOUNTER + 1
      'On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      For iRow = 1 To iTotalItems
         sPackSlip = cmbPsl & "-" & vItems(iRow, 0)
         cPartCost = GetPartCost(sPartGroup(iRow), ES_STANDARDCOST)
         bResponse = GetPartAccounts(sPartGroup(iRow), sCreditAcct, sDebitAcct)

'         'Activity
'         lCOUNTER = lCOUNTER + 1
'         cQuantity = Format(Val(vItems(iRow, 2)), ES_QuantityDataFormat)
'         sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
'                & "INPDATE,INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," _
'                & "INPSNUMBER,INPSITEM,INNUMBER,INUSER) " _
'                & "VALUES(" & IATYPE_CancPackSlip & ",'" & sPartGroup(iRow) & "','CANCELED PS PRINT'," _
'                & "'" & sPackSlip & "','" & vDate & "','" & vDate & "'," & cQuantity & "," & cQuantity & "," _
'                & cPartCost & ",'" & sCreditAcct & "','" & sDebitAcct & "','" _
'                & Trim(cmbPsl) & "'," & Val(vItems(iRow, 0)) & "," & lCOUNTER & ",'" _
'                & sInitials & "')"
'          clsADOCon.ExecuteSQL sSql 'rdExecDirect
         
         'Lots here
         cQuantity = 0
         iItem = Val(vItems(iRow, 0))
         bLotRec = GetLots(iItem)
         cItmLot = 0
         If bLotRec > 0 Then
            For iLotList = 1 To bLotRec
               
               'Activity
               lCOUNTER = lCOUNTER + 1
               cQuantity = cQuantity + Format(Val(sLots(iLotList, 1)), ES_QuantityDataFormat)
               sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                      & "INPDATE,INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," _
                      & "INLOTNUMBER, INPSNUMBER,INPSITEM,INNUMBER,INUSER) " _
                      & "VALUES(" & IATYPE_CancPackSlip & ",'" & sPartGroup(iRow) & "','CANCELED PS PRINT'," _
                      & "'" & sPackSlip & "','" & vDate & "','" & vDate & "'," & Abs(Val(sLots(iLotList, 1))) & "," & Abs(Val(sLots(iLotList, 1))) & "," _
                      & cPartCost & ",'" & sCreditAcct & "','" & sDebitAcct & "','" _
                      & sLots(iLotList, 0) & "','" & Trim(cmbPsl) & "'," & Val(vItems(iRow, 0)) & "," & lCOUNTER & ",'" _
                      & sInitials & "')"
               
               Debug.Print sSql
               clsADOCon.ExecuteSql sSql 'rdExecDirect
               
               lLOTRECORD = GetNextLotRecord(sLots(iLotList, 0))
               cItmLot = cItmLot + Val(sLots(iLotList, 1))
               'insert lot transaction here
               sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                      & "LOITYPE,LOIPARTREF,LOIQUANTITY," _
                      & "LOIPSNUMBER,LOIPSITEM,LOICUST," _
                      & "LOIACTIVITY,LOICOMMENT) " _
                      & "VALUES('" & sLots(iLotList, 0) & "'," _
                      & lLOTRECORD & "," & IATYPE_CancPackSlip & ",'" & sPartGroup(iRow) & "'," _
                      & Abs(Val(sLots(iLotList, 1))) & ",'" & Trim(cmbPsl) & "'," & iItem & ",'" _
                      & sLots(iLotList, 2) & "'," & lCOUNTER & ",'Canceled PS Print')"
                clsADOCon.ExecuteSql sSql 'rdExecDirect
               
               sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY" _
                      & "+" & Abs(sLots(iLotList, 1)) & " WHERE LOTNUMBER='" & sLots(iLotList, 0) & "'"
                clsADOCon.ExecuteSql sSql 'rdExecDirect
            Next
         Else
            
               lCOUNTER = lCOUNTER + 1
               cQuantity = Format(Val(vItems(iRow, 2)), ES_QuantityDataFormat)
               sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                      & "INPDATE,INADATE,INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," _
                      & "INPSNUMBER,INPSITEM,INNUMBER,INUSER) " _
                      & "VALUES(" & IATYPE_CancPackSlip & ",'" & sPartGroup(iRow) & "','CANCELED PS PRINT'," _
                      & "'" & sPackSlip & "','" & vDate & "','" & vDate & "'," & Abs(cQuantity) & "," & Abs(cQuantity) & "," _
                      & cPartCost & ",'" & sCreditAcct & "','" & sDebitAcct & "','" _
                      & Trim(cmbPsl) & "'," & Val(vItems(iRow, 0)) & "," & lCOUNTER & ",'" _
                      & sInitials & "')"
               
               clsADOCon.ExecuteSql sSql 'rdExecDirect
            
         End If
         
         'Update Part Qoh
         sSql = "UPDATE PartTable SET PAQOH=PAQOH+" & Abs(cQuantity) & "," _
                & "PALOTQTYREMAINING=PALOTQTYREMAINING+" & Abs(cQuantity) & " " _
                & "WHERE PARTREF='" & sPartGroup(iRow) & "' "
          clsADOCon.ExecuteSql sSql 'rdExecDirect
         AverageCost sPartGroup(iRow)
      Next
      sSql = "UPDATE PshdTable SET PSPRINTED=NULL,PSSHIPPRINT=0," _
             & "PSSHIPPEDDATE=NULL,PSSHIPPED=0 WHERE PSNUMBER='" _
             & Trim(cmbPsl) & "' "
       clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      'Erase date stamp for PsitTable 11/4/03
      sSql = "UPDATE PsitTable SET PILOTNUMBER='' " _
             & "WHERE PIPACKSLIP='" & Trim(cmbPsl) & "' "
       clsADOCon.ExecuteSql sSql 'rdExecDirect
       
       ' Update the SoitTable shipped flag
      sSql = "UPDATE SoitTable SET ITPSSHIPPED=0 " _
             & "WHERE ITPSNUMBER='" & Trim(cmbPsl) & "'"
       clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      'update ia costs from their associated lots
      Dim ia As New ClassInventoryActivity
      ia.UpdatePackingSlipCosts (Trim(cmbPsl))
         
      MouseCursor 0
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         SysMsg "Print Canceled For " & cmbPsl & ".", True
         UpdateWipColumns lSysCount
         FillPackSlips
         'On Error Resume Next
         cmbPsl.SetFocus
      Else
         clsADOCon.RollbackTrans
         MsgBox "Couldn't Cancel Packing Slip Print.", _
            vbExclamation, Caption
      End If
      cmdItm.Enabled = True
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   MouseCursor ccHourglass
   cmdItm.Enabled = True
   sProcName = "CancelPrint"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetItems() As Byte
   Dim RdoPsl As ADODB.Recordset
   Dim sPackSlip As String
   
   sPackSlip = Compress(cmbPsl)
   Erase vItems
   Erase sPartGroup
   iTotalItems = 0
   On Error GoTo DiaErr1
   sSql = "SELECT PIITNO,PITYPE,PIQTY,PIPART,PISONUMBER," _
          & "PISOITEM,PISOREV,PILOTNUMBER FROM PsitTable WHERE " _
          & "PIPACKSLIP='" & sPackSlip & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPsl, ES_FORWARD)
   If bSqlRows Then
      With RdoPsl
         Do Until .EOF
            iTotalItems = iTotalItems + 1
            vItems(iTotalItems, 0) = !PIITNO
            vItems(iTotalItems, 1) = !PITYPE
            vItems(iTotalItems, 2) = Format(!PIQTY, ES_QuantityDataFormat)
            vItems(iTotalItems, 3) = "" & Trim(!PIPART)
            sPartGroup(iTotalItems) = "" & Trim(!PIPART)
            vItems(iTotalItems, 4) = !PISONUMBER
            vItems(iTotalItems, 5) = !PISOITEM
            vItems(iTotalItems, 6) = "" & Trim(!PISOREV)
            sTimeStamp(iTotalItems) = "" & Trim(!PILOTNUMBER)
            .MoveNext
         Loop
         ClearResultSet RdoPsl
      End With
      GetItems = 1
   Else
      GetItems = False
   End If
   Set RdoPsl = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



'10/7/04

Private Function GetLots(iPsitem As Integer) As Byte
   Dim RdoLots As ADODB.Recordset
   Dim iRow As Integer
   Dim iTotalLots As Integer
   
   Dim sOldLots(50, 2) As String
   Erase sLots
   iTotalLots = 0
   GetLots = 0
   'MM - No need to include LOIACTIVITY in the group by because the
   ' if the PS is canceled we will cancel twice.
'   sSql = "SELECT LOINUMBER, MAX(LOIRECORD) AS LOTRECORD," _
'          & "LOIACTIVITY FROM LoitTable WHERE (LOIPSNUMBER='" _
'          & Trim(cmbPsl) & "' AND LOIPSITEM=" & iPsitem & " " _
'          & "AND LOITYPE=25) GROUP BY LOINUMBER,LOIACTIVITY"
   
   sSql = "SELECT LOINUMBER, MAX(LOIRECORD) AS LOTRECORD " _
          & " FROM LoitTable WHERE (LOIPSNUMBER='" _
          & Trim(cmbPsl) & "' AND LOIPSITEM=" & iPsitem & " " _
          & "AND LOITYPE=25) GROUP BY LOINUMBER"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
   If bSqlRows Then
      With RdoLots
         Do Until .EOF
            iTotalLots = iTotalLots + 1
            sOldLots(iTotalLots, 0) = "" & Trim(!LOINUMBER)
            sOldLots(iTotalLots, 1) = Trim$(str$(!LOTRECORD))
            .MoveNext
         Loop
         ClearResultSet RdoLots
      End With
   End If
   For iRow = 1 To iTotalLots
      sSql = "SELECT LOINUMBER,LOIPARTREF,LOIADATE,LOIQUANTITY,LOIPSNUMBER,LOIPSITEM," _
             & "LOICUST FROM LoitTable WHERE (LOINUMBER='" & sOldLots(iRow, 0) & "' AND " _
             & "LOIRECORD=" & Val(sOldLots(iRow, 1)) & ")"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
      If bSqlRows Then
         With RdoLots
            Do Until .EOF
               GetLots = GetLots + 1
               sLots(GetLots, 0) = "" & Trim(!LOINUMBER)
               sLots(GetLots, 1) = Val(!LOIQUANTITY)
               sLots(GetLots, 2) = "" & Trim(!LOICUST)
               .MoveNext
            Loop
            ClearResultSet RdoLots
         End With
      End If
   Next
   
   Set RdoLots = Nothing
   
   Exit Function
   
DiaErr1:
   GetLots = 0
End Function
