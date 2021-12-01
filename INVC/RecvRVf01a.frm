VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RecvRVf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel A Purchase Order Receipt"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RecvRVf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optDoc 
      Alignment       =   1  'Right Justify
      Caption         =   "On Dock Inspected"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CheckBox optInv 
      Alignment       =   1  'Right Justify
      Caption         =   "Invoiced"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CheckBox optRcd 
      Alignment       =   1  'Right Justify
      Caption         =   "Received"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdItm 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      ToolTipText     =   "Cancel Purchase Order Receipt"
      Top             =   720
      Width           =   915
   End
   Begin VB.TextBox txtRev 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      ToolTipText     =   "Enter Revision (If Any)"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtItm 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "Enter Item Number"
      Top             =   1440
      Width           =   495
   End
   Begin VB.ComboBox cmbPon 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Select Or Enter PO"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5640
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3210
      FormDesignWidth =   6315
   End
   Begin VB.Label lblPdt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5160
      TabIndex        =   24
      ToolTipText     =   "Received Date"
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   11
      Left            =   4440
      TabIndex        =   23
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblADate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblMoRun 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblMoPart 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   3240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   14
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   13
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2640
      TabIndex        =   12
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label lblNik 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   9
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Release"
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label txtRel 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5880
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "RecvRVf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'7/20/04 Added CancelRunAllocated
'12/29/04 fixed GetItem to allow for nulls
'4/27/05 Corrected bug in Button settings
'5/2/05 Added GetLots (accomodate multiple selections
'5/6/05 Delete Lots for Receipts (GetPoLots)
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim bCanceled As Byte
Dim bOnLoad As Byte
Dim bGoodPo As Byte
Dim bGoodPoItem As Byte
Dim iPkRecord As Integer
Dim lRunno As Long
Dim cStdCost As Currency
Dim cQuantity As Currency

Dim sPartNumber As String
Dim sRunPart As String
Dim sCreditAcct As String
Dim sDebitAcct As String
Dim sLotNumber As String
Dim sLotNum(50) As String
Dim sLotQty(50) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim b As Byte
   On Error GoTo DiaErr1
   sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   If b = 0 Then
      MsgBox "There Is No Open Inventory Journal For This Period.", _
         vbExclamation, Caption
      Sleep 500
      Unload Me
      Exit Sub
   Else
      sJournalID = GetOpenJournal("PJ", Format$(ES_SYSDATE, "mm/dd/yy"))
      If Left(sJournalID, 4) = "None" Then
         sJournalID = ""
         b = 1
      Else
         If sJournalID = "" Then b = 0 Else b = 1
      End If
      If b = 0 Then
         MsgBox "There Is No Open Purchases Journal For This Period.", _
            vbExclamation, Caption
         Sleep 500
         Unload Me
         Exit Sub
      End If
   End If
   sProcName = "fillcombo"
   
   cmbPon.Clear
   sSql = "SELECT DISTINCT PONUMBER,PINUMBER,PITYPE " _
          & "FROM PohdTable,PoitTable WHERE (PONUMBER=PINUMBER " _
          & "AND PITYPE=15) ORDER BY PONUMBER DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         cmbPon = Format(!PONUMBER, "000000")
         Do Until .EOF
            AddComboStr cmbPon.hWnd, Format$(!PONUMBER, "000000")
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   Set RdoCmb = Nothing
   bGoodPo = GetPurchaseOrder()
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub cmbPon_Click()
   bGoodPo = GetPurchaseOrder()
   
End Sub

Private Sub cmbPon_GotFocus()
   cmbPon_Click
   
End Sub


Private Sub cmbPon_LostFocus()
   cmbPon = CheckLen(cmbPon, 6)
   cmbPon = Format(Abs(Val(cmbPon)), "000000")
   bGoodPo = GetPurchaseOrder()
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5350"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdItm_Click()
   Dim iLotCount As Integer
   If optRcd.Value = vbUnchecked Then
      MsgBox "This Item Has Not Been Received.", _
         vbInformation, Caption
      Exit Sub
   End If
   If optInv.Value = vbChecked Then
      MsgBox "This Item Has Been Invoiced.", _
         vbInformation, Caption
      Exit Sub
   End If
   If optDoc.Value = vbChecked Then
      MsgBox "This Item Has Been On Dock Inspected.", _
         vbInformation, Caption
      Exit Sub
   End If
   If Val(cmbPon) = 0 Or Val(txtItm) = 0 Then
      MsgBox "Requires a Valid PO and Item.", vbInformation, Caption
   Else
      bGoodPoItem = GetItem()
      If bGoodPoItem Then
         iLotCount = CheckLots()
         If iLotCount > 1 Then
            If iLotCount < 3 And Trim(lblMoPart) <> "" And Val(lblMoRun) > 0 Then
               CancelRunAllocated
            Else
               MsgBox "This Item Has Activity On The Lot " & vbCr _
                  & "And Cannot Be Canceled.", _
                  vbInformation, Caption
            End If
         Else
            CancelThisItem
         End If
      Else
         MsgBox "Item Wasn't Found or Was Canceled or Invoiced.", vbInformation, Caption
      End If
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillCombo
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sSql = "SELECT PONUMBER,PORELEASE,POVENDOR,VEREF,VENICKNAME,VEBNAME " _
          & "FROM PohdTable,VndrTable WHERE (POVENDOR=VEREF AND " _
          & "PONUMBER= ? )"
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   'RdoQry.MaxRows = 1
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adInteger
   AdoQry.Parameters.Append AdoParameter
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set RecvRVf01a = Nothing
   
End Sub


Private Sub txtItm_LostFocus()
   txtItm = CheckLen(txtItm, 3)
   txtItm = Format(Abs(Val(txtItm)), "##0")
   If bCanceled Then Exit Sub
   bGoodPoItem = GetItem()
   If bGoodPo Then cmdItm.Enabled = True Else cmdItm.Enabled = False
   
End Sub

Private Sub txtRev_LostFocus()
   txtRev = CheckLen(txtRev, 2)
   If bCanceled Then Exit Sub
   If bGoodPo Then cmdItm.Enabled = True Else cmdItm.Enabled = False
   bGoodPoItem = GetItem()
   
End Sub



'Lots 3/13/02

Private Sub CancelThisItem()
   Dim bGoodlots As Byte
   Dim bResponse As Byte
   
   Dim lCOUNTER As Long
   Dim lSysCount As Long
   
   Dim cNewQty As Currency
   
   Dim sMsg As String
   Dim sPon As String
   
   On Error GoTo RcarcCi1
   sMsg = "Are You Certain That You Want " & vbCr _
          & "Cancel This Item?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      On Error Resume Next
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      cNewQty = cQuantity - (2 * cQuantity)
      lCOUNTER = GetLastActivity()
      lSysCount = lCOUNTER + 1
      sPon = "PO " & cmbPon & "-" & txtRel & " Item " & txtItm & txtRev
      sSql = "UPDATE PoitTable SET PITYPE=14,PIAQTY=0" & "," _
             & "PILOTNUMBER='', PIADATE=NULL " _
             & "WHERE PINUMBER=" & Val(cmbPon) & " " _
             & "AND PIRELEASE=" & Val(txtRel) & " AND PIITEM=" _
             & Val(txtItm) & " AND PIREV='" & Trim(txtRev) & "' "
      clsADOCon.ExecuteSQL sSql
      
      bGoodlots = GetPoLots()
      Err.Clear
      
      '2/13/01 No activity if allocated
      lCOUNTER = lCOUNTER + 1
      If sRunPart = "" And lRunno = 0 Then
         If bGoodlots = 0 Then
            'Not on used for an MO
            sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                   & "INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INPONUMBER," _
                   & "INPORELEASE,INPOITEM,INPOREV,INNUMBER,INLOTNUMBER,INUSER) " _
                   & "VALUES(16,'" & sPartNumber & "'," _
                   & "'CANCELED PO RECEIPT','" & sPon & "'," _
                   & str(cNewQty) & "," & str(cNewQty) & "," _
                   & Val(cStdCost) & ",'" & sCreditAcct & "','" & sDebitAcct & "'," _
                   & Val(cmbPon) & "," & Val(txtRel) & "," & Val(txtItm) & ",'" _
                   & txtRev & "'," & lCOUNTER & ",'" & sLotNumber & "','" & sInitials & "')"
            clsADOCon.ExecuteSQL sSql
            

               ReturnLotsToInventory False
               

         End If
         sSql = "UPDATE PartTable SET PAQOH=PAQOH-" & str(Abs(cNewQty)) & "," _
                & "PALOTQTYREMAINING=PALOTQTYREMAINING-" & str(Abs(cNewQty)) & " " _
                & "WHERE PARTREF='" & sPartNumber & "'"
         clsADOCon.ExecuteSQL sSql
         AverageCost sPartNumber
      Else
         'Used On an MO
         If bGoodlots = 0 Then
            sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                   & "INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INPONUMBER," _
                   & "INPORELEASE,INPOITEM,INPOREV,INNUMBER,INLOTNUMBER,INUSER) " _
                   & "VALUES(16,'" & sPartNumber & "'," _
                   & "'CANCELED MO ALLOCAT','" & sPon & "'," _
                   & cQuantity & "," & cQuantity & "," _
                   & cStdCost & ",'" & sCreditAcct & "','" & sDebitAcct & "'," _
                   & Val(cmbPon) & "," & Val(txtRel) & "," & Val(txtItm) & ",'" _
                   & txtRev & "'," & lCOUNTER & ",'" & sLotNumber & "','" & sInitials & "')"
            clsADOCon.ExecuteSQL sSql
         End If
         If iPkRecord > 0 Then
            sSql = "DELETE FROM MopkTable where (PKMOPART='" & sRunPart & "' " _
                   & "AND PKMORUN=" & lRunno & " AND PKRECORD=" & iPkRecord & ")"
            clsADOCon.ExecuteSQL sSql
         End If
         If sLotNumber <> "" Then
            If bGoodlots = 0 Then
               'Effectively cancel the lot

            ReturnLotsToInventory True
            
               lCOUNTER = GetLastActivity()
               sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                      & "INPQTY,INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INPONUMBER," _
                      & "INPORELEASE,INPOITEM,INPOREV,INNUMBER,INLOTNUMBER,INUSER) " _
                      & "VALUES(16,'" & sPartNumber & "'," & "'CANCELED PO RECEIPT','" _
                      & sPon & "'," & cNewQty & "," & cNewQty & "," _
                      & cStdCost & ",'" & sCreditAcct & "','" & sDebitAcct & "'," _
                      & Val(cmbPon) & "," & Val(txtRel) & "," & Val(txtItm) & ",'" _
                      & txtRev & "'," & lCOUNTER & ",'" & sLotNumber & "','" & sInitials & "')"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
      MouseCursor 0
      If clsADOCon.ADOErrNum = 0 Then

         sSql = "UPDATE InvaTable SET INPDATE=INADATE WHERE INTYPE=16 AND " _
                & "INPDATE IS NULL"
         clsADOCon.ExecuteSQL sSql
         UpdateWipColumns lSysCount
         clsADOCon.CommitTrans
         txtItm = "0"
         txtRev = ""
         MsgBox "Receipt Was Canceled.", vbInformation, Caption
         On Error Resume Next
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "Couldn't Cancel Receipt.", vbExclamation, Caption
      End If
      cmbPon.SetFocus
   Else
      CancelTrans
   End If
   Exit Sub
   
RcarcCi1:
   sProcName = "cancelthisit"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume RcarcCi2
RcarcCi2:
   On Error Resume Next
   clsADOCon.RollbackTrans
   MouseCursor 0
   sMsg = str(CurrError.Number) & vbCr _
          & CurrError.Description & vbCr _
          & "Could Not Cancel Receipt."
   MsgBox sMsg, vbExclamation, Caption
   DoModuleErrors Me
   
End Sub


Private Function GetItem() As Byte
   Dim RdoRcp As ADODB.Recordset
   Dim bByte As Byte
   'Runs added 2/13/01
   lRunno = 0
   sRunPart = ""
   optRcd.Value = vbUnchecked
   optInv.Value = vbUnchecked
   optDoc.Value = vbUnchecked
   
   On Error GoTo DiaErr1
   sSql = "SELECT PINUMBER,PIRELEASE,PIITEM,PIREV,PITYPE,PIAQTY," _
          & "PIRUNPART,PIRUNNO,PILOTNUMBER,PIADATE,PIPDATE,PIPICKRECORD," _
          & "PARTREF,PARTNUM,PADESC,PASTDCOST FROM PoitTable,PartTable " _
          & "WHERE PIPART=PARTREF AND PINUMBER=" & Val(cmbPon) & " " _
          & "AND PIRELEASE=" & Val(txtRel) & " AND PIITEM=" _
          & Val(txtItm) & " AND PIREV='" & Trim(txtRev) & "' " _
          & "AND PITYPE<>16"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRcp, ES_FORWARD)
   If bSqlRows Then
      With RdoRcp
         lblPdt = Format(!PIPDATE, "mm/dd/yy")
         If !PITYPE = 14 Then
            cmdItm.Enabled = False
            optRcd.Value = vbUnchecked
            optInv.Value = vbUnchecked
            optDoc.Value = vbUnchecked
         ElseIf !PITYPE = 15 Then optRcd.Value = vbChecked
         ElseIf !PITYPE = 17 Then
            optRcd.Value = vbChecked
            optInv.Value = vbChecked
         ElseIf !PITYPE = 18 Then
            optRcd.Value = vbChecked
            optDoc.Value = vbChecked
         End If
         
         cQuantity = !PIAQTY
         cStdCost = !PASTDCOST
         sPartNumber = "" & Trim(!PartRef)
         lblPrt(0) = "" & Trim(!PartNum)
         lblPrt(1) = "" & Trim(!PADESC)
         sRunPart = "" & Trim(!PIRUNPART)
         lRunno = !PIRUNNO
         lblMoPart = sRunPart
         lblMoRun = str$(lRunno)
         On Error Resume Next
         If Not IsNull(!PIADATE) Then
            lblADate = Format(!PIADATE, "mm/dd/yy")
         Else
            lblADate = Format(!PIADATE, "mm/dd/yy")
         End If
         iPkRecord = !PIPICKRECORD
         sLotNumber = "" & Trim(!PILOTNUMBER)
         cmdItm.Enabled = True
         GetItem = True
         ClearResultSet RdoRcp
      End With
      bByte = GetPartAccounts(sPartNumber, sDebitAcct, sCreditAcct)
   Else
      iPkRecord = 0
      cQuantity = 0
      cStdCost = 0
      sPartNumber = ""
      lblPrt(0) = ""
      lblPrt(1) = "Item Wasn't Found Or Was Canceled."
      sLotNumber = ""
      GetItem = False
   End If
   Set RdoRcp = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetPurchaseOrder() As Byte
   Dim RdoRcp As ADODB.Recordset
   On Error GoTo DiaErr1
   lblPrt(0) = ""
   lblPrt(1) = ""
   txtItm = 0
   txtRev = ""
   optRcd.Value = vbUnchecked
   optInv.Value = vbUnchecked
   optDoc.Value = vbUnchecked
   cmdItm.Enabled = False
   AdoQry.Parameters(0).Value = Val(cmbPon)
   bSqlRows = clsADOCon.GetQuerySet(RdoRcp, AdoQry, ES_FORWARD, False, 1)
   If bSqlRows Then
      GetPurchaseOrder = True
      With RdoRcp
         lblNik = "" & !VENICKNAME
         lblNme = "" & !VEBNAME
         ClearResultSet RdoRcp
      End With
   Else
      lblNik = ""
      lblNme = ""
      GetPurchaseOrder = False
   End If
   Set RdoRcp = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpurchaseorder"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


'3/12/02 lots

Private Function CheckLots() As Integer
   Dim RdoLots As ADODB.Recordset
   Dim iCount As Integer
   
   On Error GoTo DiaErr1
   If sLotNumber <> "" Then
      sSql = "SELECT LOINUMBER FROM LoitTable WHERE LOINUMBER='" _
             & sLotNumber & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
      If bSqlRows Then
         With RdoLots
            Do Until .EOF
               iCount = iCount + 1
               .MoveNext
            Loop
            ClearResultSet RdoLots
         End With
      End If
   End If
   CheckLots = iCount
   Set RdoLots = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checklots"
   CheckLots = 0
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

'7/21/04 Resolve problem with allocated items

Private Sub CancelRunAllocated()
   Dim bResponse As Byte
   Dim sMsg As String
   
   sMsg = "Are You Certain That You Want " & vbCr _
          & "Cancel This Item?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      sSql = "DELETE FROM MopkTable WHERE (PKPARTREF='" & Compress(lblPrt(0)) & "' " _
             & "AND PKMOPART='" & lblMoPart & "' AND PKMORUN=1 AND " _
             & "PKADATE='" & lblADate & "')"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM LoitTable WHERE LOINUMBER='" & sLotNumber & "'"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM LohdTable WHERE LOTNUMBER='" & sLotNumber & "'"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM InvaTable WHERE INLOTNUMBER='" & sLotNumber & "'"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE PoitTable SET PITYPE=14,PIAQTY=0" & "," _
             & "PILOTNUMBER='',PIADATE=NULL " _
             & "WHERE (PINUMBER=" & Val(cmbPon) & " " _
             & "AND PIRELEASE=" & Val(txtRel) & " AND PIITEM=" _
             & Val(txtItm) & " AND PIREV='" & Trim(txtRev) & "')"
     Debug.Print sSql
     
      clsADOCon.ExecuteSQL sSql
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         MsgBox "Receipt Was Canceled.", vbInformation, Caption
      Else
         clsADOCon.RollbackTrans
         MsgBox "Couldn't Cancel Receipt.", vbExclamation, Caption
      End If
      lblMoPart = ""
      lblMoRun = ""
      lblADate = ""
      sLotNumber = ""
      FillCombo
   Else
      CancelTrans
   End If
   
End Sub

'5/2/05

Private Function GetPoLots() As Byte
   Dim RdoPol As ADODB.Recordset
   Dim iList As Integer
   Dim iRows As Integer
   Dim lCOUNTER As Long
   Dim cCost As Currency
   Dim cQuantity As Currency
   Dim sPon As String
   
   Erase sLotNum()
   Erase sLotQty()

   On Error Resume Next
   sSql = "SELECT LOTNUMBER,LOTPO,LOTPOITEM,LOTPOITEMREV,LOTADATE," _
          & "LOTORIGINALQTY FROM LohdTable WHERE (LOTPO=" & Val(cmbPon) & " " _
          & "AND LOTPOITEM=" & Val(txtItm) & " AND LOTPOITEMREV='" _
          & Trim(txtRev) & "' AND LOTADATE>='" & lblADate & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPol, ES_FORWARD)
   If bSqlRows Then
      With RdoPol
         Do Until .EOF
            iRows = iRows + 1
            sLotNum(iRows) = "" & Trim(!lotNumber)
            sLotQty(iRows) = "" & Trim(!LOTORIGINALQTY)
            .MoveNext
         Loop
         ClearResultSet RdoPol
      End With
   End If
   
   '5/6/05 Don't mess with then, delete 'em. Saves multiple lots
   For iList = 1 To iRows
      
      sSql = "DELETE FROM LoitTable WHERE LOINUMBER='" & sLotNum(iList) & "'"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM InvaTable WHERE INLOTNUMBER='" & sLotNum(iList) & "'"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "DELETE FROM LohdTable WHERE LOTNUMBER='" & sLotNum(iList) & "'"
      clsADOCon.ExecuteSQL sSql
   Next
   
   If clsADOCon.ADOErrNum <> 0 Then GetPoLots = 0 Else GetPoLots = 1
   Set RdoPol = Nothing
   Exit Function
   
DiaErr1:
   clsADOCon.RollbackTrans
   GetPoLots = 0
   MsgBox "Failed"
End Function


Private Sub ReturnLotsToInventory(UsedOnMO As Boolean)

    Dim iLots As Integer
    Dim iCounter As Long
    Dim iRecord As Integer
    Dim sQty As String
    
    
    If UsedOnMO Then iRecord = 3 Else iRecord = 2
    For iLots = 1 To UBound(sLotNum) - 1
        If sLotNum(iLots) = "" Then Exit For
        If UsedOnMO Then sQty = "0" Else sQty = "-" & Trim(sLotQty(iLots))
    
        sSql = "UPDATE LohdTable SET LOTORIGINALQTY=0,LOTREMAININGQTY=0," _
                      & "LOTAVAILABLE=0 WHERE LOTNUMBER='" & sLotNum(iLots) & "'"
        clsADOCon.ExecuteSQL sSql
    
        iCounter = GetLastActivity()
        sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                      & "LOITYPE,LOIPARTREF,LOIQUANTITY," _
                      & "LOIPONUMBER,LOIPOITEM,LOIPOREV,LOIVENDOR," _
                      & "LOIACTIVITY,LOICOMMENT) " _
                      & "VALUES('" _
                      & sLotNum(iLots) & "'," & Trim(str(iRecord)) & ",16,'" & Compress(lblPrt(0)) _
                      & "'," & sQty & "," & cmbPon & "," & txtItm & ",'" _
                      & txtRev & "','" & Compress(lblNik) & "'," _
                      & iCounter & ",'" _
                      & "Canceled PO Receipt" & "')"
               clsADOCon.ExecuteSQL sSql

    
    Next iLots
    

End Sub


