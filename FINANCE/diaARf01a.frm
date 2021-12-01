VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel An Invoice"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDel 
      Caption         =   "C&ancel"
      Height          =   315
      Left            =   4560
      TabIndex        =   11
      ToolTipText     =   "Cancel Selected Invoice"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbInv 
      Height          =   315
      ItemData        =   "diaARf01a.frx":0000
      Left            =   1800
      List            =   "diaARf01a.frx":0002
      TabIndex        =   2
      Tag             =   "3"
      Top             =   360
      Width           =   1125
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4560
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaARf01a.frx":0004
      PictureDn       =   "diaARf01a.frx":014A
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4080
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2595
      FormDesignWidth =   5520
   End
   Begin VB.Label lbldte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Date"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice/Memo  Number"
      Height          =   525
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   1185
   End
   Begin VB.Label lblPre 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   360
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Type"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblMemo 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label txtNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   1320
   End
End
Attribute VB_Name = "diaARf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions


'*************************************************************************************
' diaARf01a - Cancel an Invoice
'
' Created: (cjs)
' Revisons:
' 09/09/02 (nth) Renamed and updated misc errors fixed
' 12/04/03 (nth) change "Invoice Canceled" to a sysmsg
' 10/20/04 (nth) Check if applied to a cash receipt before canceling.
' 01/19/05 (nth) Correct memos being reversed in current journal rather than posting date journal.
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bGoodInv As Byte

Dim sLots(50, 2) As String

Dim iTotalItems As Integer

Dim sCreditAcct As String
Dim sDebitAcct As String
Dim sJournalID As String

Dim vItems(300, 4) As Variant
'0 = SO #
Private Const CANCEL_SO = 0
'1 = Part #
Private Const CANCEL_PART = 1
'2 = Std Cost
Private Const CANCEL_COST = 2
'3 = Qty
Private Const CANCEL_QTY = 3

Dim sMsg As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub cmbInv_Click()
   bGoodInv = GetInvoice()
End Sub

Private Sub cmbInv_LostFocus()
   cmbInv = CheckLen(cmbInv, 6)
   cmbInv = Format(Abs(Val(cmbInv)), "000000")
   If Val(cmbInv) > 0 Then bGoodInv = GetInvoice()
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdDel_Click()
   Dim sCR As String
   If Not bGoodInv Then
      MsgBox "You Must Select A Valid Invoice.", _
         vbInformation, Caption
   Else
      ' Check for cash receipt
      sCR = AppliedToCR(cmbInv)
      If sCR <> "" Then
         sMsg = "Cash Receipt " & sCR & " Is Applied" & vbCrLf _
                & "To Invoice. Cannot Cancel."
         MsgBox sMsg, vbInformation, Caption
      Else
         Select Case Trim(lblTyp)
            Case "SO"
               CancelSalesOrder
            Case "PS"
               CancelPackslip
            Case Else
               CancelMemo
         End Select
      End If
   End If
   MouseCursor 0
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Cancel An Invoice"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CurrentJournal "SJ", ES_SYSDATE, sJournalID
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaARf01a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim b As Byte
   On Error GoTo DiaErr1
   cmbInv.Clear
   lblTyp = ""
   lblMemo = ""
   lblCst = ""
   txtNme = ""
   
   sProcName = "fillcombo"
   
   'list invoices with open sales journals and no cash receipts:
   'SELECT INVNO FROM CihdTable
   'join JrhdTable on MJTYPE = 'SJ' and MJSTART <= INVDATE and MJEND >= INVDATE  and MJCLOSED is null
   'WHERE INVTYPE<>'TM' AND (INVPIF=0 AND INVCANCELED=0)
   'and INVNO not in (select CAINVNO FROM CashTable)
   'ORDER BY INVNO DESC
   
   
   sSql = "SELECT INVNO FROM CihdTable" & vbCrLf _
          & "join JrhdTable on MJTYPE = 'SJ' and MJSTART <= INVDATE and MJEND >= INVDATE  and MJCLOSED is null" & vbCrLf _
          & "WHERE INVTYPE<>'TM' AND (INVPIF=0 AND INVCANCELED=0) " & vbCrLf _
          & "and INVNO not in (select CAINVNO FROM CashTable)" & vbCrLf _
          & "ORDER BY INVNO DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         cmbInv = Format(!InvNo, "000000")
         Do Until .EOF
            AddComboStr cmbInv.hWnd, Format$(!InvNo, "000000")
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   If cmbInv.ListCount > 0 Then
      cmbInv = cmbInv.List(0)
      bGoodInv = GetInvoice()
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function GetInvoice() As Byte
   
   'returns True if invoice # is valid and cancellable
   '(i.e. open journal for that invdate and no cash receipts)
   'SELECT INVNO FROM CihdTable
   'join JrhdTable on MJTYPE = 'SJ' and MJSTART <= INVDATE and MJEND >= INVDATE  and MJCLOSED is null
   'WHERE INVTYPE<>'TM' AND (INVPIF=0 AND INVCANCELED=0)
   'and INVNO not in (select CAINVNO FROM CashTable)
   'ORDER BY INVNO DESC
   Dim RdoInv As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT INVNO,INVPRE,INVTYPE,INVCUST,INVDATE FROM CihdTable " & vbCrLf _
          & "join JrhdTable on MJTYPE = 'SJ' and MJSTART <= INVDATE and MJEND >= INVDATE  and MJCLOSED is null" & vbCrLf _
          & "WHERE INVNO=" & Val(cmbInv) & " AND (INVPIF=0 AND INVCANCELED=0)" & vbCrLf _
          & "and INVNO not in (select CAINVNO FROM CashTable)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
   If bSqlRows Then
      With RdoInv
         cmbInv = Format(!InvNo, "000000")
         lblPre = "" & Trim(!INVPRE)
         lblTyp = "" & Trim(!INVTYPE)
         lbldte = Format(!INVDATE, "mm/dd/yy")
         FindCustomer Me, "" & Trim(!INVCUST)
         .Cancel
      End With
      GetInvoice = True
   Else
      lblPre = ""
      lblTyp = ""
      lblCst = ""
      txtNme = "*** Not a Cancellable Invoice ***"
      GetInvoice = False
   End If
   Set RdoInv = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   txtNme = "*** No Current Invoice ***"
   DoModuleErrors Me
   
   
End Function

Private Sub lblTyp_Change()
   If lblTyp = "" Then
      lblTyp = "***"
      lblTyp.ForeColor = ES_RED
   Else
      lblTyp.ForeColor = vbBlack
   End If
   lblMemo = GetInvoiceType(lblTyp)
End Sub

Private Sub txtNme_Change()
   If Left(txtNme, 3) = "***" Then
      txtNme.ForeColor = ES_RED
   Else
      txtNme.ForeColor = vbBlack
   End If
End Sub

Private Sub CancelMemo()
   Dim bResponse As Byte
   Dim bRet As Boolean
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   sMsg = "This Function Will Permanently Remove The Memo." & vbCrLf _
          & "Do You Really Want To Cancel This " & lblMemo & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      
      sJournalID = InvJrnClosed
      If sJournalID = "" Then Exit Sub
      
      
      bRet = AdjustInvForPreInvoice(Val(cmbInv))
      
      If bRet = True Then
      
         sSql = "UPDATE CihdTable SET INVFREIGHT=0," _
                & "INVTAX=0,INVTOTAL=0,INVCANCELED=1," _
                & "INVCHECKDATE=NULL," & vbCrLf _
                & "INVCANCDATE='" & Format(Now, "mm/dd/yy") & "'" & vbCrLf _
                & "WHERE INVNO=" & Val(cmbInv) & " "
         
         '                & Format(Now, "mm/dd/yy") & "'," _

         clsADOCon.ExecuteSQL sSql
         If clsADOCon.RowsAffected > 0 Then
            If sJournalID <> "" Then ReverseJournal
            SysMsg "Invoice Canceled.", True, Me
            FillCombo
         Else
            MsgBox "Could Not Cancel The Invoice.", _
               vbExclamation, Caption
         End If
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "cancelmemo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub CancelPackslip()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sJournal As String
   'Dim rdoJrn As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sMsg = "This Function Will Permanently Remove The Invoice." & vbCrLf _
          & "Do You Really Want To Cancel This " & lblMemo & " Invoice?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      
      sJournalID = InvJrnClosed
      If sJournalID = "" Then Exit Sub
      
      On Error Resume Next
      Err = 0
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      sSql = "UPDATE PshdTable SET PSINVOICE=0 " _
             & "WHERE PSINVOICE=" & Val(cmbInv)
      clsADOCon.ExecuteSQL sSql
      
      ' Update lot record
      sSql = "UPDATE LoitTable SET LOICUSTINVNO=0" & vbCrLf _
         & "WHERE LOICUSTINVNO=" & Val(cmbInv)
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE CihdTable SET INVFREIGHT=0," _
             & "INVTAX=0,INVTOTAL=0,INVPACKSLIP=''," _
             & "INVCANCELED=1," _
             & "INVCHECKDATE=NULL, " & vbCrLf _
             & "INVCANCDATE='" & Format(Now, "mm/dd/yy") & "'" & vbCrLf _
             & "WHERE INVNO=" & Val(cmbInv) & " "
      clsADOCon.ExecuteSQL sSql
      
'      sSql = "UPDATE SoitTable SET ITCGSACCT=''," _
'             & "ITACTUAL=NULL,ITINVOICE=0 " _
'             & "WHERE ITINVOICE=" & Val(cmbInv) & " "
'      clsAdoCon.ExecuteSQL sSQL
      
      sSql = "UPDATE SoitTable SET" & vbCrLf _
         & "ITINVOICE=0," & vbCrLf _
         & "ITREVACCT=''," & vbCrLf _
         & "ITCGSACCT=''," & vbCrLf _
         & "ITBOSTATE=''," & vbCrLf _
         & "ITBOCODE=''," & vbCrLf _
         & "ITSLSTXACCT=''," & vbCrLf _
         & "ITTAXCODE=''," & vbCrLf _
         & "ITSTATE=''," & vbCrLf _
         & "ITTAXRATE=0," & vbCrLf _
         & "ITTAXAMT=0" & vbCrLf _
         & "WHERE ITINVOICE=" & Val(cmbInv)
      clsADOCon.ExecuteSQL sSql
      
      ReverseJournal
      
      If clsADOCon.ADOErrNum = 0 Then
         
         clsADOCon.CommitTrans
         SysMsg "Invoice Canceled.", True, Me
         
         FillCombo
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         
         MsgBox "Could Not Cancel Invoice.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "cancelpac"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub CancelSalesOrder()
   Dim bSuccess As Byte
   Dim bResponse As Byte
   Dim i As Integer
   Dim sDate As String
   Dim sMsg As String
   
   sDate = Format(Now, "mm/dd/yy")
   On Error GoTo DiaErr1
   
   sMsg = "This Function Will Permanently Remove The Invoice." & vbCrLf _
          & "Do You Really Want To Cancel This " & lblMemo & " Invoice?"
   
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      
      sJournalID = InvJrnClosed
      If sJournalID = "" Then Exit Sub
      
      GetSalesOrder
      On Error Resume Next
      Err = 0
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      For i = 1 To iTotalItems
         sSql = "UPDATE SoitTable SET ITCGSACCT=''," _
                & "ITACTUAL=NULL,ITINVOICE=0 " _
                & "WHERE ITINVOICE=" & Val(cmbInv) & " "
         clsADOCon.ExecuteSQL sSql
         
         ' Update Part Qoh
         sSql = "UPDATE PartTable SET PAQOH=PAQOH+" & Val(vItems(i, CANCEL_QTY)) & " " _
                & "WHERE PARTREF='" & vItems(i, CANCEL_PART) & "' "
         clsADOCon.ExecuteSQL sSql
         AverageCost LTrim(str(vItems(i, 10)))
         
         'Add to Activity
         GetAccounts i
         sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
                & "INADATE,INAQTY,INAMT,INCREDITACCT,INDEBITACCT) " _
                & "VALUES(5,'" & vItems(i, CANCEL_PART) _
                & "','CANCELED INVOICE','SO " & Format(vItems(i, CANCEL_SO), "00000") & " INV " & cmbInv _
                & "','" & sDate & "'," & Val(vItems(i, CANCEL_QTY)) & "," _
                & Val(vItems(i, CANCEL_COST)) & ",'" & sCreditAcct & "','" _
                & sDebitAcct & "') "
         clsADOCon.ExecuteSQL sSql
      Next
      
      If clsADOCon.ADOErrNum = 0 Then
         
         sSql = "UPDATE CihdTable SET INVFREIGHT=0," _
                & "INVTAX=0,INVTOTAL=0,INVPACKSLIP=0," _
                & "INVCANCELED=1,INVCHECKDATE=NULL, " & vbCrLf _
                & "INVCANCDATE='" & Format(Now, "mm/dd/yy") & "'" & vbCrLf _
                & "WHERE INVNO=" & Val(cmbInv) & " "
         clsADOCon.ExecuteSQL sSql
         
         If clsADOCon.ADOErrNum = 0 Then
            If sJournalID <> "" Then ReverseJournal
            clsADOCon.CommitTrans
            SysMsg "Invoice Canceled.", True, Me
            FillCombo
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            MsgBox "Could Not Cancel The Invoice.", _
               vbExclamation, Caption
         End If
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "cancelsaleso"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetSalesOrder()
   Dim RdoSon As ADODB.Recordset
   Dim i As Integer
   Erase vItems
   iTotalItems = 0
   On Error GoTo DiaErr1
   sSql = "SELECT ITSO,ITPART,ITQTY,ITINVOICE,PARTREF,PASTDCOST " _
          & "FROM SoitTable,PartTable WHERE ITPART=PARTREF AND " _
          & "ITINVOICE=" & Val(cmbInv) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon)
   If bSqlRows Then
      With RdoSon
         Do Until .EOF
            i = i + 1
            vItems(i, CANCEL_SO) = Format(!ITSO)
            vItems(i, CANCEL_PART) = "" & Trim(!PartRef)
            vItems(i, CANCEL_COST) = Format(!PASTDCOST)
            vItems(i, CANCEL_QTY) = Format(!ITQTY)
            .MoveNext
         Loop
         .Cancel
         iTotalItems = i
      End With
   End If
   Set RdoSon = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsalesor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetAccounts(iIndex As Integer)
   Dim rdoAct As ADODB.Recordset
   Dim bType As Byte
   Dim sPcode As String
   
   On Error GoTo DiaErr1
   'Use current Part
   sSql = "Qry_GetExtPartAccounts '" & vItems(iIndex, CANCEL_PART) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         sPcode = "" & Trim(!PAPRODCODE)
         bType = Format(!PALEVEL, "0")
         If bType = 6 Or bType = 7 Then
            sCreditAcct = "" & Trim(!PACGSEXPACCT)
            sDebitAcct = "" & Trim(!PAINVEXPACCT)
         Else
            sCreditAcct = "" & Trim(!PACGSMATACCT)
            sDebitAcct = "" & Trim(!PAINVMATACCT)
         End If
         .Cancel
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
            If bType = 6 Or bType = 7 Then
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(!PCCGSEXPACCT)
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(!PCINVEXPACCT)
            Else
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(!PCCGSMATACCT)
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(!PCINVMATACCT)
            End If
            .Cancel
         End With
      End If
      If sDebitAcct = "" Or sCreditAcct = "" Then
         'Still none, we'll check the common
         If bType = 6 Or bType = 7 Then
            sSql = "SELECT COREF,COCGSEXPACCT" & Trim(str(bType)) & "," _
                   & "COINVEXPACCT" & Trim(str(bType)) & " " _
                   & "FROM ComnTable WHERE COREF=1"
         Else
            sSql = "SELECT COREF,COCGSMATACCT" & Trim(str(bType)) & "," _
                   & "COINVMATACCT" & Trim(str(bType)) & " " _
                   & "FROM ComnTable WHERE COREF=1"
         End If
         bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
         If bSqlRows Then
            With rdoAct
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(.Fields(0))
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(.Fields(1))
               .Cancel
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

'Reverses the Journal Entries

Private Sub ReverseJournal()
   Dim rdoJrn As ADODB.Recordset
   Dim i As Integer
   Dim iItems As Integer
   Dim iTrans As Integer
   Dim iRef As Integer
   Dim lInv As Long
   Dim vTrans(300, 11) As Variant
   '1 = DCREF
   '2 = DCDEBIT (OLD)
   '3 = DCCREDIT (OLD)
   '4 = DCACCTNO
   '5 = DCPARTNO
   '6 = DCSONUMBER
   '7 = DCSOITNUMBER
   '8 = DCSOITREV
   '9 = DCCUST
   '10 = DCINVO
   
   iTrans = GetNextTransaction(sJournalID)
   
   MouseCursor 13
   On Error GoTo DiaErr1
   lInv = Val(cmbInv)
   sSql = "SELECT DCHEAD,DCTRAN,DCREF,DCDEBIT,DCCREDIT,DCACCTNO,DCPARTNO," _
          & "DCSONUMBER,DCSOITNUMBER,DCSOITREV,DCCUST,DCINVNO " _
          & "FROM JritTable WHERE DCINVNO=" & lInv & " "
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      With rdoJrn
         Do Until .EOF
            iItems = iItems + 1
            vTrans(iItems, 1) = !DCREF
            vTrans(iItems, 2) = !DCDEBIT 'old db
            vTrans(iItems, 3) = !DCCREDIT 'old cr
            vTrans(iItems, 4) = "" & Trim(!DCACCTNO)
            vTrans(iItems, 5) = "" & Trim(!DCPARTNO)
            vTrans(iItems, 6) = !DCSONUMBER
            vTrans(iItems, 7) = !DCSOITNUMBER
            vTrans(iItems, 8) = !DCSOITREV
            vTrans(iItems, 9) = Trim(!DCCUST)
            vTrans(iItems, 10) = !DCINVNO
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   
   
   For i = 1 To iItems
      iRef = i
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT," _
             & "DCCREDIT,DCACCTNO,DCPARTNO,DCSONUMBER,DCSOITNUMBER,DCSOITREV," _
             & "DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & vTrans(i, 3) & "," _
             & vTrans(i, 2) & ",'" _
             & vTrans(i, 4) & "','" _
             & vTrans(i, 5) & "'," _
             & vTrans(i, 6) & "," _
             & vTrans(i, 7) & ",'" _
             & vTrans(i, 8) & "','" _
             & vTrans(i, 9) & "','" _
             & lbldte & "'," _
             & vTrans(i, 10) & ")"
      
      'old date logic:
      '                & Format(Now, "mm/dd/yy") & "'," _

      
      clsADOCon.ExecuteSQL sSql
   Next
   Set rdoJrn = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "reversejourn"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function InvJrnClosed() As String
   Dim rdoJrn As ADODB.Recordset
   Dim sMsg As String
   
   ' Figure out what journal it is in.
   ' If the journal is posted then do not allow invoice to be canceled
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT MJCLOSED, MJDESCRIPTION, MJGLJRNL " _
          & "FROM JrhdTable INNER JOIN " _
          & "JritTable ON JrhdTable.MJGLJRNL = JritTable.DCHEAD " _
          & "WHERE (JrhdTable.MJTYPE = 'SJ') AND (JritTable.DCINVNO = " _
          & Val(cmbInv) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn)
   
   If bSqlRows Then
      With rdoJrn
         If ("" & Trim(!MJCLOSED)) <> "" Then
            sMsg = "Invoice " & Val(cmbInv) & " Resides In Closed Journal " & Trim(!MJDESCRIPTION) _
                   & vbCrLf & " Unable To Cancel Invoice."
            MsgBox sMsg, vbInformation, Caption
            InvJrnClosed = ""
         Else
            InvJrnClosed = Trim(!MJGLJRNL)
         End If
      End With
   End If
   Set rdoJrn = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "invjrnclosed"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Function AppliedToCR(lInvoice As Long) As String
   Dim rdoCR As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT DCCHECKNO FROM JritTable Where DCINVNO = " _
          & lInvoice & " AND DCCHECKNO <> ''"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCR)
   If bSqlRows Then
      With rdoCR
         AppliedToCR = "" & Trim(.Fields(0))
         .Cancel
      End With
   End If
   Set rdoCR = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "appliedtocr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function


Private Function GetPreviousInv(ByVal lCurInv As Long, ByRef lPreInv As Long) As Boolean
   Dim RdoInv As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT INVPRIORINV FROM CihdTable WHERE INVNO = " & lCurInv
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv)
   If bSqlRows Then
      With RdoInv
         lPreInv = "" & Trim(.Fields(0))
         .Cancel
      End With
      GetPreviousInv = True
   Else
      GetPreviousInv = False
   End If
   Set RdoInv = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetPreviousInv"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function


Private Function AdjustInvForPreInvoice(lInvoice As Long) As Boolean
   
   Dim b As Byte
   Dim bRet As Boolean
   
   Dim RdoItm As ADODB.Recordset
   Dim RdoQty As ADODB.Recordset
   Dim lCurInv As Long
   Dim lPreInv As Long
   Dim iTotalItems As Integer
   
   Dim strSoNum As String
   Dim strSoit As Integer
   Dim strSORev  As String
   Dim iLotTrack As Integer
   Dim cUnitPrice As Currency
   Dim cLotQty As Currency
   Dim cInvQty As Currency
   Dim cSoQty As Currency
   
   Dim strPackSlip  As String
   Dim strPsItem  As String
   Dim strPartNum  As String
   Dim strCust  As String
   
   Dim vAdate As Variant
   Dim cPartCost As Currency
   Dim bByte As Byte
   
   'calculate quantity remaining
   Dim cRetQty As Currency

   iTotalItems = 0
   On Error GoTo DiaErr1
   'Dim RdoRet As ADODB.Recordset
   
   vAdate = GetServerDateTime
   bRet = GetPreviousInv(lInvoice, lPreInv)
   
   If (lPreInv = 0) Then
      AdjustInvForPreInvoice = True
      Exit Function
   End If
   
   If (bRet = True) Then

      sSql = "SELECT ITNUMBER, ITREV, ITSO, PARTNUM, ITACTUAL, ITDOLLARS," & vbCrLf _
         & "ITQTY,SOTYPE, PALOTTRACK, ITPSNUMBER, ITPSITEM, SOCUST " & vbCrLf _
         & "FROM SoitTable" & vbCrLf _
         & "JOIN PartTable ON ITPART = PARTREF" & vbCrLf _
         & "JOIN SohdTable ON ITSO = SONUMBER" & vbCrLf _
         & "WHERE ITINVOICE = " & lPreInv
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_STATIC)
      If bSqlRows Then
         
         clsADOCon.BeginTrans
         clsADOCon.ADOErrNum = 0
         
         With RdoItm
            Dim iList As Integer
            'determine lots from which items are drawn
            sSql = "delete from TempPsLots where PsNumber = '" & !ITPSNUMBER & "'" & vbCrLf _
               & "or DateDiff( hour, WhenCreated, getdate() ) > 24"
            clsADOCon.ExecuteSQL sSql 'rdExecDirect
            
            While Not .EOF
                           
               strSoNum = !ITSO
               strSoit = !ITNUMBER
               strSORev = !itrev
               iLotTrack = !PALOTTRACK
               cUnitPrice = !ITDOLLARS
               strPackSlip = !ITPSNUMBER
               strPsItem = !ITPSITEM
               strPartNum = !PARTNUM
               strCust = !SOCUST
               'cRetQty = !ITQTY
               cSoQty = !ITQTY
               
               cInvQty = 0
               sSql = "SELECT ISNULL(SUM(INAQTY),0) FROM InvaTable " & vbCrLf _
                      & "WHERE INSONUMBER = " & strSoNum & vbCrLf _
                      & "AND INSOITEM = " & !ITNUMBER & vbCrLf _
                      & "AND INSOREV = '" & Trim(!itrev) & "' AND INTYPE = 4" ' So Item returned.
               bSqlRows = clsADOCon.GetDataSet(sSql, RdoQty, ES_FORWARD)
               If bSqlRows Then
                  cRetQty = RdoQty.Fields(0)
               End If
               RdoQty.Cancel
               Set RdoQty = Nothing
               
               ' Insert Lot and Inv record
               If (cRetQty > 0) Then
                  Dim cRemPQty As Currency
                  Dim iLots As Integer
                  Dim cPckQty As Currency
                  Dim cItmLot As Currency
                  Dim strLot As String
                  Dim bInvType As Byte
                  
                  Dim cInAmt As Currency
                  Dim cUnitOvHd As Currency
                  Dim cUnitLabor As Currency
                  Dim cUnitMatl As Currency
                  Dim cUnitExp As Currency
                  Dim cUnitHrs As Currency
                  
                  Dim cUnitCost As Currency
                  Dim cOvHd As Currency
                  Dim cLabor As Currency
                  Dim cMatl As Currency
                  Dim cExp As Currency
                  Dim cHours As Currency
                  
                  ' store the tmp value
                  cRemPQty = cRetQty
                  iLots = GetPartLots(Compress(Trim(strPartNum)))
                  cItmLot = 0
                  
                  For iList = 1 To iLots
                     
                     cLotQty = Val(sLots(iList, 1))
                     If cLotQty >= cRemPQty Then
                        cPckQty = cRemPQty
                        cLotQty = cLotQty - cRemPQty
                        cRemPQty = 0
                     Else
                        cPckQty = cLotQty
                        cRemPQty = cRemPQty - cLotQty
                        cLotQty = 0
                     End If
                     If cPckQty > 0 Then
                        cItmLot = cItmLot + cPckQty
                        If cItmLot > Val(sLots(iList, 1)) Then cItmLot = Val(sLots(iList, 1))
                        strLot = sLots(iList, 0)
                        
                        GetInventoryCost strLot, cUnitCost, cUnitMatl, cUnitExp, cUnitLabor, cUnitOvHd, cUnitHrs
                        
                        cMatl = cUnitMatl * cRetQty
                        cLabor = cUnitLabor * cRetQty
                        cExp = cUnitExp * cRetQty
                        cOvHd = cUnitOvHd * cRetQty
                        cHours = cUnitHrs * cRetQty
                        
                        sSql = "INSERT INTO TempPsLots ( PsNumber, PsItem, LotID, LotQty , PartRef, LotItemID)" & vbCrLf _
                           & "Values ( '" & strPackSlip & "', " & strPsItem & ", " _
                           & "'" & strLot & "', " & cPckQty _
                           & ", '" & Compress(strPartNum) & "', '" & CStr(iList) & "') "
                        clsADOCon.ExecuteSQL sSql 'rdExecDirect
                     
                     End If
                     
                     
                  Next
                  ' If still we have remaining Qty we need to quit
                  If (cRemPQty > 0) Then
                     MsgBox "Not sufficient quantity for item " & strPsItem _
                        & " part " & strPartNum & " available. " & vbCrLf _
                        & "It is short by (" & cRemPQty & ") quantity." & vbCrLf _
                        & "We can not Cancel the Credit MO for an Invoice."
                     GoTo NoCanDo
                  End If
                  
                  ' set the cost as standard cost
                  If (cUnitCost = 0) Then
                     cPartCost = GetPartCost(strPartNum, ES_STANDARDCOST)
                     cPartCost = Format(cPartCost, ES_QuantityDataFormat)
                  Else
                     cPartCost = Format(cUnitCost, ES_QuantityDataFormat)
                  End If
                  
                  bByte = GetPartAccounts(strPartNum, sCreditAcct, sDebitAcct)
                  bInvType = IATYPE_PackingSlip
                  
                  Dim sSql1 As String
                  Dim sSql2 As String
                  Dim sSql3 As String
                  
                  'create inventory activities for lots for this packing slip item
                  ' Fusion 5/15/2009
                  sSql1 = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2, " & vbCrLf _
                     & "INNUMBER,INPDATE,INADATE,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," & vbCrLf _
                     & "INTOTMATL,INTOTLABOR,INTOTEXP,INTOTOH,INTOTHRS," & vbCrLf _
                     & "INPSNUMBER,INPSITEM,INLOTNUMBER,INSONUMBER,INSOITEM,INSOREV) " & vbCrLf _
                     & "SELECT " & bInvType & ", '" & Compress(strPartNum) & "', 'PACKING SLIP', "
                     
                  sSql2 = "tmp.PsNumber + '-' + " & "cast( tmp.PsItem as varchar(5) )," & vbCrLf _
                        & "(SELECT MAX(INNUMBER) as num FROM INVATABLE) +  tmp.LotItemID," & vbCrLf _
                        & "'" & vAdate & "', '" & vAdate & "',  -tmp.LotQty, " _
                        & cPartCost & ", '" & sDebitAcct & "', '" & sCreditAcct & "', " & vbCrLf _
                        & cMatl & "," & cLabor & "," & cExp & "," & cOvHd & "," & cHours & "," & vbCrLf _
                        & "'" & strPackSlip & "', " & strPsItem & ", " _
                        & "tmp.LotID, " & strSoNum & ", "
                  
                  sSql3 = strSoit & ", '" & strSORev & "'" & vbCrLf _
                        & "FROM TempPsLots tmp" & vbCrLf _
                        & "JOIN PartTable pt on tmp.PARTREF = pt.PartRef" & vbCrLf _
                        & "WHERE tmp.PsNumber = '" & strPackSlip & "' AND tmp.PsItem = " & strPsItem
                     
                  sSql = sSql1 & sSql2 & sSql3
                     
                  Debug.Print sSql
                  
                  clsADOCon.ExecuteSQL sSql 'rdExecDirect
                  
                  'insert lot items for this packing slip item
                  sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                     & "LOITYPE,LOIPARTREF,LOIADATE,LOIQUANTITY," & vbCrLf _
                     & "LOIPSNUMBER,LOIPSITEM,LOICUST,LOIACTIVITY,LOICOMMENT) " & vbCrLf _
                     & "SELECT tmp.LotID, dbo.fnGetNextLotItemNumber( tmp.LotID ), " _
                     & bInvType & ", '" & Compress(strPartNum) & "', '" & vAdate & "', " & vbCrLf _
                     & "-tmp.LotQty, '" & strPackSlip & "', " _
                     & Val(strPsItem) & ", '" & strCust & "'," _
                     & "ia.INNUMBER, 'Shipped Item'" & vbCrLf _
                     & "FROM TempPsLots tmp" & vbCrLf _
                     & "JOIN InvaTable ia ON ia.INPSNUMBER = tmp.PsNumber AND ia.INPSITEM = tmp.PsItem" & vbCrLf _
                     & "and ia.INADATE = '" & vAdate & "' and ia.INLOTNUMBER = tmp.LotID" & vbCrLf _
                     & "WHERE tmp.PsNumber = '" & strPackSlip & "' AND tmp.PsItem = " & Trim(strPsItem)
                  
                  Debug.Print sSql
                  clsADOCon.ExecuteSQL sSql 'rdExecDirect
                     
                     
                  'update quantities for part
                  sSql = "UPDATE PartTable SET PAQOH=PAQOH - " & cRetQty & ", " _
                         & "PALOTQTYREMAINING = PALOTQTYREMAINING - " & cRetQty & vbCrLf _
                         & "WHERE PARTREF='" & Compress(strPartNum) & "' "
                  clsADOCon.ExecuteSQL sSql 'rdExecDirect
               
               End If
               
               .MoveNext
            Wend
         End With
         
         'update remaining quantity in affected lots
         sSql = "UPDATE LohdTable" & vbCrLf _
            & "SET LOTREMAININGQTY = X.TOTAL" & vbCrLf _
            & "FROM LohdTable lt" & vbCrLf _
            & "JOIN (SELECT LOINUMBER, SUM(LOIQUANTITY) AS TOTAL FROM LOITTABLE GROUP BY LOINUMBER) AS X" & vbCrLf _
            & "ON X.LOINUMBER = LOTNUMBER" & vbCrLf _
            & "WHERE LOTNUMBER IN ( SELECT LotID from TempPsLots where PsNumber = '" & strPackSlip & "' )"
         
         Debug.Print sSql
         
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         AdjustInvForPreInvoice = True
      Else
         clsADOCon.RollbackTrans
         AdjustInvForPreInvoice = False
         
         MsgBox "Could Not Cancel Invoice.", _
            vbExclamation, Caption
      End If
         
         
      End If
      Set RdoItm = Nothing
   End If
   
   
   Exit Function
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me

NoCanDo:
   MouseCursor 0
   AdjustInvForPreInvoice = False
   
   Exit Function

End Function

Private Function GetPartLots(sPartWithLot As String) As Integer
   Dim RdoLots As ADODB.Recordset
   Dim iRow As Integer
   Dim bFIFO As Boolean
   
   Erase sLots
   On Error GoTo DiaErr1
   
   bFIFO = GetInventoryMethod()
   
   sSql = "SELECT LOTNUMBER,LOTPARTREF,LOTREMAININGQTY,LOTAVAILABLE " _
          & "FROM LohdTable WHERE (LOTPARTREF='" & sPartWithLot & "' AND " _
          & "LOTREMAININGQTY > 0 AND LOTAVAILABLE=1) " '
   If bFIFO = 1 Then
      sSql = sSql & "ORDER BY LOTNUMBER ASC"
   Else
      sSql = sSql & "ORDER BY LOTNUMBER DESC"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
   If bSqlRows Then
      With RdoLots
         Do Until .EOF
            If (iRow >= 49) Then Exit Do
            iRow = iRow + 1
            sLots(iRow, 0) = "" & Trim(!lotNumber)
            sLots(iRow, 1) = Format$(!LOTREMAININGQTY, ES_QuantityDataFormat)
            .MoveNext
         Loop
         ClearResultSet RdoLots
      End With
      GetPartLots = iRow
   Else
      GetPartLots = 0
   End If
   Set RdoLots = Nothing
   Exit Function
   
DiaErr1:
   GetPartLots = 0
   
End Function

Private Function GetInventoryCost(ByVal strLotNum As String, ByRef cUnitCost As Currency, _
            ByRef cUnitMatl As Currency, ByRef cUnitExp As Currency, ByRef cUnitLabor As Currency, _
            ByRef cUnitOvHd As Currency, ByRef cUnitHrs As Currency)
   
   Dim RdoLotCost As ADODB.Recordset
   
   sSql = "SELECT LOTUNITCOST,cast(LOTTOTMATL / LOTORIGINALQTY as decimal(12,4)) as TotMatl," _
            & "cast ( LOTTOTLABOR / LOTORIGINALQTY as decimal(12,4)) as TotLab," _
            & "cast ( LOTTOTEXP / LOTORIGINALQTY as decimal(12,4)) as TotExp," _
            & "cast ( LOTTOTOH / LOTORIGINALQTY as decimal(12,4)) as TotOH," _
            & "cast ( LOTTOTHRS / LOTORIGINALQTY as decimal(12,4)) as TotHrs " _
      & "From LohdTable WHERE LOTNUMBER = '" & strLotNum & "' AND LOTORIGINALQTY <> 0"
   
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLotCost, ES_FORWARD)
   If (bSqlRows = True) Then
      With RdoLotCost
         cUnitCost = !LotUnitCost
         cUnitOvHd = !TotOH
         cUnitLabor = !TotLab
         cUnitMatl = !TotMatl
         cUnitExp = !TotExp
         cUnitHrs = !TotHrs
      End With
      RdoLotCost.Close
   Else
      cUnitCost = 0
      cUnitOvHd = 0
      cUnitLabor = 0
      cUnitMatl = 0
      cUnitExp = 0
      cUnitHrs = 0
   End If
   
   Set RdoLotCost = Nothing
   
End Function


