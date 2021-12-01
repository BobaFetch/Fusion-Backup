VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SaleSLf15c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Post VOI Invoice"
   ClientHeight    =   10800
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   10860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPSDte 
      Height          =   315
      Left            =   1920
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdPostInvoice 
      Caption         =   "Post Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8520
      TabIndex        =   8
      ToolTipText     =   "Post the invoice"
      Top             =   2640
      Width           =   2280
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Height          =   315
      Left            =   4800
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Cancel This Sales Order"
      Top             =   720
      Width           =   915
   End
   Begin VB.CommandButton CmdSelAll 
      Caption         =   "Selection All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8520
      TabIndex        =   5
      ToolTipText     =   " Select All"
      Top             =   3600
      Width           =   2280
   End
   Begin VB.ComboBox cmbCheck 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   720
      Width           =   2760
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLf15c.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   10800
      FormDesignWidth =   10860
   End
   Begin VB.CommandButton cmdCnc 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6480
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Cancel This Sales Order"
      Top             =   480
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   14160
      Top             =   10320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin MSComDlg.CommonDialog ExpDlg 
      Left            =   8640
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   7455
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   13150
      _Version        =   393216
      Rows            =   3
      Cols            =   8
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   315
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      Caption         =   "Posted Date"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   14
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Invoice Total"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblInvTotal 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      ToolTipText     =   "Last Sales Order Entered"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Check Amount"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblTotChkAmt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      ToolTipText     =   "Last Sales Order Entered"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7080
      Picture         =   "SaleSLf15c.frx":07AE
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7080
      Picture         =   "SaleSLf15c.frx":0B38
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "SaleSLf15c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Added ITINVOICE

Option Explicit

Dim cmdObj1 As ADODB.Command
Dim cmdObj2 As ADODB.Command

Dim bFIFO As Byte
Dim bGoodJrn As Boolean
Private Const PS_PACKSLIPNO = 0
Private Const PS_ITEMNO = 1
Private Const PS_QUANTITY = 2
'Private Const PS_PIPART = 3
Private Const PS_COST = 4
Private Const PS_LOTTRACKED = 5
Private Const PS_PARTNUM = 6

Dim vItems(800, 7) As Variant
Dim sPartGroup(800) As String '9/23/04 Compressed PartTable!PARTREF
Dim sSoItems(300, 3) As String 'Nathan 3/10/04
Dim sLots(50, 2) As String
Dim sCustomer As String
Dim sCreditAcct As String
Dim sDebitAcct As String

Const SOITEM_SO = 0 ' string of PISONUMBER
Const SOITEM_ITEM = 1 ' string of PISOITEM
Const SOITEM_REV = 2 ' string of PISOREV

' Invoice global variables
Dim bGoodAct As Byte
Dim bGoodPs As Byte
Dim iRow As Integer
Dim lSo As Long
Dim lNewInv As Long
Dim lNextInv As Long
Dim lSalesOrder As Long
Dim iTotalItems As Integer
Dim iTotalChk As Integer
Dim sPsCust As String
Dim sPsStadr As String
Dim sPackSlip As String
Dim sAccount As String
Dim sMsg As String
Dim sDocNumber As String


Dim sTaxAccount As String
Dim sTaxState As String
Dim sTaxCode As String
Dim nTaxRate As Currency
Dim cFREIGHT As Currency
Dim cTotInvAmtSel As Currency

Dim cTax As Currency
Dim sType As String * 1
Dim currentCust As String

Dim sCrCashAcct As String
Dim sCrDiscAcct As String
Dim sCrExpAcct As String
Dim sSJARAcct As String
Dim sCrRevAcct As String
Dim sCrCommAcct As String


' Sales journal
Dim sCOSjARAcct As String
Dim sCOSjINVAcct As String
Dim sCOSjNFRTAcct As String
Dim sCOSjTFRTAcct As String
Dim sCOSjTaxAcct As String
Public lCurrInvoice As Long
Dim vpItems(1000, 12) As Variant


Dim bCutOff As Byte
Dim bOnLoad As Byte
Dim bUnload As Boolean
Dim strXML As String
Dim bNewImport As Boolean
Dim ExtName As String

Dim sJournalID As String

Dim Fields(150) As String

Dim sCust As String
Dim cDiscount As Currency

Private txtKeyPress As New EsiKeyBd




Private Sub cmdCan_Click()
   Unload Me

End Sub



Private Sub cmdHlp_Click()
    If cmdHlp Then
        MouseCursor (13)
        OpenHelpContext (2150)
        MouseCursor (0)
        cmdHlp = False
    End If

End Sub

Private Sub FillGrid()

   Dim CheckNo As String
   Dim invAmount As Currency
   Dim crdAmount As Currency
   Dim totAmount As Currency
   
   CheckNo = cmbCheck
   

   Dim RdoCheck As ADODB.Recordset
   
   sSql = "select DISTINCT FusionSOVOI.Payment_doc_no, SUM(voi.Issue_Amount) as TotAmount, (voi.check_amt * -1) check_amt," _
            & "    ISNULL(FusionSOVOI.ITPSNUMBER, '') ITPSNUMBER, ISNULL(FusionSOVOI.INNO, '')INNO, ISNULL(FusionSOVOI.CASHRECEIPT, '') CASHRECEIPT " _
            & " FROM (SELECT DISTINCT Payment_doc_no,Payment_doc_it_no," _
            & " ISNULL(ITPSNUMBER, '') ITPSNUMBER, " _
            & " ISNULL(INNO, '')INNO, ISNULL(CASHRECEIPT, '') CASHRECEIPT " _
            & "    FROM FusionSOVOI) AS FusionSOVOI inner join tmpVOIPmtIss AS voi" _
            & "    ON FusionSOVOI.Payment_doc_no = voi.payment_doc_no " _
            & "    And FusionSOVOI.Payment_doc_it_no = voi.Payment_doc_it_no " _
            & " WHERE voi.CHECK_NO = '" & CheckNo & "' " _
            & " GROUP BY FusionSOVOI.Payment_doc_no, " _
            & "       FusionSOVOI.ITPSNUMBER , INNO, CASHRECEIPT,voi.check_amt "
            
            
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCheck, ES_DYNAMIC)

   Grd.Rows = 1
   Debug.Print sSql
   
   If bSqlRows Then
   With RdoCheck
      lblTotChkAmt = Trim(!check_amt)
      While Not .EOF

         Grd.Rows = Grd.Rows + 1
         Grd.Row = Grd.Rows - 1
         
         Grd.Col = 0
         Set Grd.CellPicture = Chkno.Picture
         
         Grd.Col = 1
         Grd.Text = Trim(!payment_doc_no)
         
         Grd.Col = 2
         Grd.Text = Trim(!check_amt)
         
         crdAmount = 0
         totAmount = 0
         
         
         GetInvoiceCreditAmount CheckNo, Trim(!payment_doc_no), crdAmount
         totAmount = Trim(!totAmount) + crdAmount
         
         Grd.Col = 3
         Grd.Text = totAmount 'Trim(!totAmount)
         
         ' get Invoice amount
         GetInvoiceAmount Trim(!INNO), invAmount
         
         Grd.Col = 4
         Grd.Text = invAmount

         
         Grd.Col = 5
         Grd.Text = Trim(!ITPSNUMBER)
         
         Grd.Col = 6
         Grd.Text = Trim(!INNO)
         
         Grd.Col = 7
         Grd.Text = Trim(!CASHRECEIPT)
         
         .MoveNext
      Wend
      .Close
      ClearResultSet RdoCheck
      End With
   End If
   Set RdoCheck = Nothing

End Sub


Private Function GetPackslip(sPack As String) As Boolean
   Dim RdoPsl As ADODB.Recordset
   nTaxRate = 0
   sTaxCode = ""
   sTaxState = ""
   sTaxAccount = ""
   
   On Error GoTo DiaErr1
   
   Erase vItems
   sSql = "SELECT DISTINCT PSNUMBER,PSCUST,PSTERMS,PSSTNAME,PSSTADR," _
          & "PSFREIGHT,CUREF,CUNICKNAME,CUNAME,PIPACKSLIP " & vbCrLf _
          & "FROM PshdTable" & vbCrLf _
          & "JOIN CustTable ON PshdTable.PSCUST = CustTable.CUREF" & vbCrLf _
          & "JOIN PsitTable ON PsitTable.PIPACKSLIP = PshdTable.PSNUMBER" & vbCrLf _
          & "AND (PSTYPE=1 AND PSSHIPPRINT=1 AND PSINVOICE=0 )" & vbCrLf _
          & "WHERE PSNUMBER = '" & sPack & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPsl, ES_DYNAMIC)
   If bSqlRows Then
      With RdoPsl
         
         sPackSlip = "" & Trim(!PsNumber)
         sPsCust = "" & Trim(!CUNICKNAME)
         sPsCust = "" & Trim(!CUREF)
         sPsStadr = "" & Trim(!PSSTNAME) & vbCrLf _
                    & Trim(!PSSTADR)
         cFREIGHT = Format(!PSFREIGHT, "#####0.00")
         .Cancel
      End With
      GetSalesTaxInfo Compress(sPsCust), nTaxRate, sTaxCode, sTaxState, sTaxAccount
      sPsStadr = CheckComments(sPsStadr)
      GetPackslip = True
   Else
      cFREIGHT = 0
      GetPackslip = False
   End If
   Set RdoPsl = Nothing
   Exit Function
DiaErr1:
   sProcName = "getpackslip"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function
Private Sub cmbPSDte_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub cmdPostInvoice_Click()
    
   Dim strDate As String
   Dim strCheckNo As String
   Dim strCheckAmt As Currency
   Dim strDocNum As String
   Dim lInvoices As Long
   Dim ldocNumber As Long
   Dim docNumber As String
   Dim TotinvAmount As Currency
   Dim issinvAmount As Currency
   Dim dueInvAmount As Currency
   
   Dim ChkTotAmount As Currency
   Dim SelInvoiceAmount As Currency
   
   Dim sCust As String
   sCust = "SPIAER"
         
   ' get posted date
   strDate = cmbPSDte  'Format(Now, "mm/dd/yyyy")
   
   sJournalID = GetOpenJournal("CR", strDate)
   If sJournalID <> "" Then
      'lTrans = GetNextTransaction(sJournalID)
   Else
      sMsg = "No Open Cash Recipts Journal Found For " & strDate & "."
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   
   ' get cash account
   GetCashAccounts
   
   strCheckNo = cmbCheck
   ChkTotAmount = Val(Abs(lblTotChkAmt))
   SelInvoiceAmount = Val(Abs(lblInvTotal))
   ' get teh check number
   'GetCheckNumVOI strDocNum, lInvoices, strCheckNo, strCheckAmt
    
   Dim acct As String
   Dim gl As GLTransaction
   Dim discountAmount As Currency
   Dim expenseAmount  As Currency
   Dim commissionAmount  As Currency
   Dim revenueAmount  As Currency
   Dim totAmount As Currency
   
   'Dim debitCash As Currency
   Dim creditAR As Currency
   'Dim creditOther As Currency
   Dim iList As Integer
   
   Dim bCaType As Byte

   bCaType = 2
   Set gl = New GLTransaction
   gl.JournalID = sJournalID 'automatically sets next transaction
   
   
   creditAR = ChkTotAmount
   acct = sCrRevAcct
   discountAmount = 0
   expenseAmount = 0
   commissionAmount = 0
   revenueAmount = 0
   
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
'   If (Not CheckExists(strCheckNo)) Then
'
'      gl.AddCashReceipt strCheckNo, sCust, bCaType, strDate, SelInvoiceAmount, _
'      creditAR, discountAmount, commissionAmount, expenseAmount, revenueAmount, strDate, sCrCashAcct
'   Else
'      MsgBox "Check number " & strCheckNo & " is alredy posted.", vbInformation, Caption
'      Exit Sub
'   End If
   
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.Row = iList
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
      
         Grd.Col = 1
         docNumber = Trim(Grd.Text)
         
         Grd.Col = 3
         issinvAmount = Val(Trim(Grd.Text))
         
         Grd.Col = 4
         TotinvAmount = Val(Trim(Grd.Text))
         
         Grd.Col = 6
         lInvoices = Val(Trim(Grd.Text))
         
         dueInvAmount = TotinvAmount - issinvAmount
         
         ' get Invoice amount
         'GetInvoiceAmount lInvoices, invAmount
         
         gl.InvoiceDate = CDate(strDate)
         
         gl.InvoiceNumber = lInvoices
               
         'calculate amount applied.  if pif, amount = inv total
         'if pif, also calculate discount
         Dim amountPaid As Currency
         Dim debitCash As Currency
         
         amountPaid = issinvAmount
         creditAR = issinvAmount
         debitCash = issinvAmount
         
         gl.AddCashReceipt strCheckNo, sCust, bCaType, strDate, SelInvoiceAmount, _
         creditAR, discountAmount, commissionAmount, expenseAmount, revenueAmount, strDate, sCrCashAcct
         
         sSql = "UPDATE CihdTable SET " _
                & "INVCHECKNO='" & strCheckNo & "'," _
                & "INVPAY=" & amountPaid & "," _
                & "INVCRDUE=" & dueInvAmount & "," _
                & "INVPIF=1," _
                & "INVADJUST=0," _
                & "INVARDISC=0," _
                & "INVDAYS=0," _
                & "INVCHECKDATE='" & strDate & "'  " _
                & "WHERE INVNO=" & lInvoices & " "
         clsADOCon.ExecuteSQL sSql
         
         Debug.Print sSql

         sSql = "UPDATE FusionSOVOI set CHECK_NO = '" & strCheckNo & "',CHECK_AMT = '" & SelInvoiceAmount & "', CASHRECEIPT = 1 WHERE PAYMENT_DOC_NO = '" & docNumber & "'"
         clsADOCon.ExecuteSQL sSql

         gl.AddDebitCredit 0, creditAR, sSJARAcct, "", 0, 0, "", sCust, strCheckNo
         gl.AddDebitCredit debitCash, 0, sCrCashAcct, "", 0, 0, "", sCust, strCheckNo
        
      End If
   Next

   If (clsADOCon.ADOErrNum = 0) Then
      clsADOCon.CommitTrans
      sMsg = "Added CheckNumber - " + CStr(strCheckNo)
      MsgBox sMsg, vbInformation, Caption
      FillGrid
   Else
      sMsg = "Adding Check Number Failed - " + CStr(strCheckNo)
      MsgBox sMsg, vbInformation, Caption
      clsADOCon.RollbackTrans
   End If
   
   Exit Sub
   
   
   
End Sub

Private Function GetInvoiceAmount(lInvoices As Long, ByRef invAmount As Currency)
   Dim RdoInvAmt As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT INVTOTAL FROM CihdTable WHERE invno = " & lInvoices
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInvAmt)
   If bSqlRows Then
      With RdoInvAmt
         If Not .EOF Then
            invAmount = !INVTOTAL
         End If
         .Close
         ClearResultSet RdoInvAmt
         
      End With
   End If
   Set RdoInvAmt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetInvoiceAmount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function
   

Private Function GetInvoiceCreditAmount(CheckNo As String, docnum As String, ByRef creditAmt As Currency)
   Dim RdoCrAmt As ADODB.Recordset
   
   On Error GoTo DiaErr1
   creditAmt = 0
   sSql = "select SUM(ISSUE_AMOUNT) AS CreditAmt from VOIPmtIss where ISSUE_AMOUNT < 0 and CHECK_NO = '" & CheckNo & "'" _
            & " and Payment_doc_no = '" & docnum & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCrAmt)
   If bSqlRows Then
      With RdoCrAmt
         If Not .EOF And Not IsNull(!creditAmt) Then
            creditAmt = !creditAmt
         End If
         .Close
         ClearResultSet RdoCrAmt
         
      End With
   End If
   Set RdoCrAmt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetInvoiceCreditAmount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function
   
Private Function GetCashAccounts() As Byte
   Dim rdoCsh As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT COGLVERIFY,COCRCASHACCT,COCRDISCACCT,COSJARACCT," _
          & "COCRCOMMACCT,COCRREVACCT,COCREXPACCT,COTRANSFEEACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCsh, ES_FORWARD)
   sProcName = "getcashacct"
   If bSqlRows Then
      With rdoCsh
         For i = 1 To 7
            If "" & Trim(.Fields(i)) = "" Then
               b = 1
               Exit For
            End If
         Next
         sCrCashAcct = "" & Trim(!COCRCASHACCT)
         sCrDiscAcct = "" & Trim(!COCRDISCACCT)
         sSJARAcct = "" & Trim(!COSJARACCT)
         sCrCommAcct = "" & Trim(!COCRCOMMACCT)
         sCrRevAcct = "" & Trim(!COCRREVACCT)
         sCrExpAcct = "" & Trim(!COCREXPACCT)
         .Cancel
         If b = 1 Then GetCashAccounts = 3 Else GetCashAccounts = 2
      End With
   Else
      GetCashAccounts = 0
   End If
   Set rdoCsh = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetCashAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetCheckNumVOI(strDocNum As String, ByRef lInvNo As Long, ByRef strCheckNo As String, ByRef strCheckAmt As Currency)
   Dim RdoChk As ADODB.Recordset
   
   On Error GoTo DiaErr1
   strCheckNo = ""
   strCheckAmt = 0
   
   sSql = "select distinct voipmtIss.check_no as check_no, voipmtIss.check_amt as check_amt, FusionSovoi.inno as inno " & _
            " from voipmtIss join FusionSovoi " & _
            "    on voipmtIss.payment_doc_no = FusionSovoi.payment_doc_no " & _
             "    and voipmtIss.payment_doc_it_no = FusionSovoi.payment_doc_it_no " & _
             " Where voipmtIss.payment_doc_no = '" & strDocNum & "'" & _
             "    and inno is not null"
   
               
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      With RdoChk
         If .EOF Then
            strCheckNo = "" & Trim(!check_no)
            strCheckAmt = "" & Trim(!check_amt)
            lInvNo = !INNO
         End If
         .Close
         ClearResultSet RdoChk
         
      End With
   End If
   Set RdoChk = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetCheckNumVOI"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Function


Private Function CheckExists(sCheck As String) As Byte
   Dim RdoChk As ADODB.Recordset

   On Error GoTo DiaErr1
   CheckExists = False
   sSql = "SELECT COUNT(CACHECKNO) FROM CashTable WHERE CACHECKNO = '" _
          & sCheck & "' AND CACUST = 'SPIAER'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      With RdoChk
         If .Fields(0) > 0 Then
            CheckExists = True
            sMsg = "Check # " & sCheck & " All Ready Exists " & vbCrLf _
                   & "For Customer - SPIAER"
            MsgBox sMsg, vbInformation, Caption
         End If
         .Cancel
      End With
   End If
   Set RdoChk = Nothing
   Exit Function

DiaErr1:
   sProcName = "CheckExists"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me


End Function




Private Sub GetSoShipTo(ByVal strDocNum As String, ByRef sCust As String, ByRef strStnme As String, _
                        ByRef strSTAdr As String, ByRef strVia As String, ByRef strTerms As String)
   
   Dim lSoNum As Long
   Dim strITSO As String
   Dim Rdodis As ADODB.Recordset
   Dim RdoSto As ADODB.Recordset
   On Error GoTo DiaErr1
   
   lSoNum = 0
   ' get any one SO
   sSql = "select Top(1) ITSO from dbo.FusionSOVOI where Payment_doc_no = '" & strDocNum & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, Rdodis, ES_FORWARD)
   If bSqlRows Then
      With Rdodis
         strITSO = "" & Trim(!itso)
         ClearResultSet Rdodis
      End With
   Else
      MsgBox ("Primary SO number is Zero")
      
   End If
   Set Rdodis = Nothing

   If strITSO <> "" Then
      lSoNum = CLng(strITSO)
   End If
   
   ' Get the ship VIA information
   sSql = "SELECT SONUMBER,SOCUST, SOSTNAME,SOSTADR, SOVIA,SOSTERMS FROM SohdTable " _
          & "WHERE SONUMBER=" & lSoNum & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSto, ES_FORWARD)
   If bSqlRows Then
      With RdoSto
         strStnme = "" & Trim(!SOSTNAME)
         strSTAdr = "" & Trim(!SOSTADR)
         strVia = "" & Trim(!SOVIA)
         strTerms = "" & Trim(!SOSTERMS)
         sCust = "" & Trim(!SOCUST)
         ClearResultSet RdoSto
      End With
   Else
      lSoNum = 0
      strStnme = ""
      strSTAdr = ""
      strVia = ""
      strTerms = ""
      If (lSoNum = 0) Then
         MsgBox ("Primary SO number is Zero")
      End If
      
   End If
   Set RdoSto = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsoship "
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Function GetValidITRev(strSoNum As String, ITNum As String, ByRef itrev As String _
                                    , ByRef ITQty As Currency) As String
   On Error GoTo DiaErr1
   
   Dim RdoRpt As ADODB.Recordset
   Dim RdoRptQ As ADODB.Recordset
   
   sSql = "SELECT MAX(ITREV) LstRev FROM soitTable WHERE ITSO ='" & strSoNum _
            & "' AND ITNUMBER = " & ITNum
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      itrev = Trim(RdoRpt!LstRev)
      ClearResultSet RdoRpt
   End If
   Set RdoRpt = Nothing
   
   sSql = "SELECT ITQTY FROM soitTable WHERE ITSO ='" & strSoNum _
            & "' AND ITNUMBER = " & ITNum & " AND ITREV = '" & itrev & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRptQ, ES_FORWARD)
   If bSqlRows Then
      ITQty = Trim(RdoRptQ!ITQty)
      ClearResultSet RdoRptQ
   End If
   Set RdoRptQ = Nothing
   
   
   
   Exit Function

DiaErr1:
   sProcName = "GetSOITPrice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Function GetSOITPrice(ByVal strSoNum As String, ByVal ITNum As Integer, _
                        ByVal rev As String, ByRef cPrice As Currency)
   On Error GoTo DiaErr1
   
   Dim RdoRpt As ADODB.Recordset
   
   sSql = "SELECT ITDOLLARS FROM soitTable WHERE ITSO ='" & strSoNum _
            & "' AND ITNUMBER = " & ITNum & " AND ITREV = '" & rev & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      cPrice = Trim(RdoRpt!ITDOLLARS)
      ClearResultSet RdoRpt
   Else
      cPrice = 0
   End If
   Set RdoRpt = Nothing
   
   Exit Function

DiaErr1:
   sProcName = "GetSOITPrice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function


Private Function GetPSfromDoc(strDocNum As String) As String
   
   Dim RdoDoc As ADODB.Recordset
   Dim strPS As String
   ' get any one SO
   sSql = "select DISTINCT PIPACKSLIP from psitTable where picomments like '%" & strDocNum & "%'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_FORWARD)
   If bSqlRows Then
      With RdoDoc
         strPS = "" & Trim(!PIPACKSLIP)
         GetPSfromDoc = strPS
         ClearResultSet RdoDoc
      End With
   Else
      MsgBox ("Document to Packslip is not Found")
      GetPSfromDoc = ""
   End If
   Set RdoDoc = Nothing

End Function

Private Function GetItems(sPackSlip As String) As Integer
   Dim RdoItm As ADODB.Recordset
   Dim iRow As Integer
   Dim bLotsAct As Byte
   Dim iTotalItems As Integer
   
   Erase vItems
   Erase sSoItems
   Erase sPartGroup
   MouseCursor 13
   
   On Error GoTo DiaErr1
   iTotalItems = 0
   bLotsAct = CheckLotStatus()
   'RdoQry2(0) = sPackSlip
   cmdObj2.parameters(0).Value = sPackSlip
   bSqlRows = clsADOCon.GetQuerySet(RdoItm, cmdObj2, ES_KEYSET, True)
   
   If bSqlRows Then
      'On Error Resume Next
      With RdoItm
         Do Until .EOF
            iRow = iRow + 1
            vItems(iRow, PS_PACKSLIPNO) = "" & Trim(!PIPACKSLIP) & "-"
            vItems(iRow, PS_ITEMNO) = Format(!PIITNO, "##0")
            vItems(iRow, PS_QUANTITY) = Format(!PIQTY, ES_QuantityDataFormat)
            'vItems(iRow, PS_PIPART) = "" & Trim(!PIPART)
            sPartGroup(iRow) = "" & Trim(!PIPART)
            vItems(iRow, PS_COST) = "0.000"
            If bLotsAct = 1 Then
               vItems(iRow, PS_LOTTRACKED) = !PALOTTRACK
            Else
               vItems(iRow, PS_LOTTRACKED) = 0
            End If
            vItems(iRow, PS_PARTNUM) = "" & Trim(!PartNum)
            sSoItems(iRow, SOITEM_SO) = str$(!PISONUMBER)
            sSoItems(iRow, SOITEM_ITEM) = str$(!PISOITEM)
            sSoItems(iRow, SOITEM_REV) = "" & Trim(!PISOREV)
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
   End If
   iTotalItems = iRow
   GetItems = iTotalItems
   Set RdoItm = Nothing
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetSprintLotRemainingQty(LotPart As String) As Currency

   'same as GetRemainingLotQty, except gets remaining qty from sum(LOTREMAININGQTY)
   'rather than SUM(LOIQUANTITY) to reduce a problem at LUMICOR
   
   Dim ADOQty As ADODB.Recordset
   
   GetSprintLotRemainingQty = 0
   sSql = "select isnull(sum(LOTREMAININGQTY),0)" & vbCrLf _
      & "from LohdTable" & vbCrLf _
      & "where LOTPARTREF='" & LotPart & "' AND LOTAVAILABLE=1 AND LOTLOCATION = 'SPRT'"
   If clsADOCon.GetDataSet(sSql, ADOQty, ES_FORWARD) Then
      GetSprintLotRemainingQty = ADOQty.Fields(0)
   End If
   Set ADOQty = Nothing
   
End Function

Private Function GetPartLots(sPartWithLot As String) As Integer
   Dim RdoLots As ADODB.Recordset
   Dim iRow As Integer
   Erase sLots
   On Error GoTo DiaErr1
   sSql = "SELECT LOTNUMBER,LOTPARTREF,LOTREMAININGQTY,LOTAVAILABLE " _
          & "FROM LohdTable WHERE (LOTPARTREF='" & sPartWithLot & "' AND " _
          & "LOTREMAININGQTY > 0 AND LOTAVAILABLE=1 AND LOTLOCATION = 'SPRT') ORDER BY LOTNUMBER ASC"
'   If bFIFO = 1 Then
'      sSql = sSql & "ORDER BY LOTNUMBER ASC"
'   Else
'      sSql = sSql & "ORDER BY LOTNUMBER DESC"
'   End If
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


Private Function GetPartComm(ByVal strGetPart As String, _
            ByRef strPartNum As String, ByRef bComm As Boolean) As Byte
   Dim RdoPrt As ADODB.Recordset
   
   On Error GoTo DiaErr1
   bComm = False
   strGetPart = Compress(strGetPart)
   If Len(strGetPart) > 0 Then
      sSql = "SELECT PARTNUM,PADESC,PAEXTDESC,PAPRICE,PAQOH," _
             & "PACOMMISSION FROM PartTable WHERE PARTREF='" & strGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_STATIC)
      If bSqlRows Then
         With RdoPrt
            strPartNum = "" & Trim(!PartNum)
            If !PACOMMISSION = 1 Then bComm = True _
                               Else bComm = False
            GetPartComm = 1
            ClearResultSet RdoPrt
         End With
      Else
         GetPartComm = 0
      End If
      'On Error Resume Next
      Set RdoPrt = Nothing
   Else
      GetPartComm = 0
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetPartComm"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



Private Sub GetCustomerRef(ByRef strCusFullName As String, ByRef strCusName As String)

   Dim RdoCus As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "SELECT DISTINCT CUREF FROM CustTable WHERE CUNAME = '" & strCusFullName & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCus)
   If bSqlRows Then
      With RdoCus
         strCusName = Trim(!CUREF)
         ClearResultSet RdoCus
      End With
   Else
      strCusName = ""
   End If
   Set RdoCus = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "GetCustomerRef"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   'GetCustomerRef = False
   DoModuleErrors Me
   

End Sub



Private Sub cmdSel_Click()
   FillGrid
   cTotInvAmtSel = 0
   lblInvTotal = Format(cTotInvAmtSel, "#####0.00")
End Sub

Private Sub CmdSelAll_Click()
   Dim iList As Integer
   
   For iList = 1 To Grd.Rows - 1
       Grd.Col = 0
       Grd.Row = iList
       ' Only if the part is checked
       If Grd.CellPicture = Chkno.Picture Then
           Set Grd.CellPicture = Chkyes.Picture
            Grd.Col = 3
            cTotInvAmtSel = cTotInvAmtSel + Val(Trim(Grd.Text))
       End If
      lblInvTotal = Format(cTotInvAmtSel, "#####0.00")
   Next
   

End Sub

Private Sub Form_Activate()
   Dim bSoAdded As Byte
   MdiSect.lblBotPanel = Caption
   
   sSql = "select distinct Check_no from tmpVOIPmtIss where CHECK_NO NOT IN ('C110447', 'C111060') AND CHECK_NO <> ''"
            
   LoadComboBox cmbCheck, -1
   
   cTotInvAmtSel = 0
   lblInvTotal = Format(cTotInvAmtSel, "#####0.00")
   
   ' Only if the import table is full
   'FillGrid
  
   If bOnLoad Then
       bOnLoad = 0
   End If
    
    
   MouseCursor (0)

End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  
  For Each ctl In Me.Controls
    If TypeOf ctl Is MSFlexGrid Then
      If IsOver(ctl.hWnd, Xpos, Ypos) Then FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
  Next ctl
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' make sure that you release the Hook
   Call WheelUnHook(Me.hWnd)
   
End Sub
Private Sub Form_Load()
    FormLoad Me, ES_DONTLIST
   
   Dim iChar As Integer
    
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 1


      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "Apply"
      .Col = 1
      .Text = "Document Number"
      .Col = 2
      .Text = "Check Amount"
      .Col = 3
      .Text = "Issue Amount"
      .Col = 4
      .Text = "Toal InvAmount"
      .Col = 5
      .Text = "Packslip"
      .Col = 6
      .Text = "Invoice"
      .Col = 7
      .Text = "Posted"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 1800
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .ColWidth(6) = 800
      .ColWidth(7) = 1000
      
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
   
   Call WheelHook(Me.hWnd)
   bOnLoad = 1

End Sub

Function ReadAllFields(ByVal iIndex As Integer, ByRef ws As Worksheet)

    Dim iCols As Integer
        
    Fields(0) = ""
    While (iCols < 150)
        Fields(iCols) = ""
        iCols = iCols + 1
    Wend
    
    iCols = 0
    If (iIndex > 0 And Not ws Is Nothing) Then
        
        While (iCols < 150)
            Fields(iCols) = ws.Cells(iIndex, iCols + 1)
            iCols = iCols + 1
        Wend
    End If

End Function

Function RemoveCommas(sNextLine As String) As String
    
    Dim length As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim strStrip As String
    lngStart = 1
    lngStart = InStr(lngStart, sNextLine, """")
    'lngEnd = InStr(lngStart + 1, sNextLine, """")
    
    length = Len(sNextLine)

    'strStrip = Left$(sNextLine, lngStart) & Right$(sNextLine, (length - lngEnd) + 1)
    
    'RemoveCommas = strStrip
    
    While (lngStart > 0)
        lngEnd = InStr(lngStart + 1, sNextLine, """")
        If (lngEnd > 0) Then
            'ReplaceComma sNextLine, lngStart, lngEnd
            sNextLine = Left$(sNextLine, lngStart) & Right$(sNextLine, (length - lngEnd) + 1)
        End If
        lngStart = InStr(1, sNextLine, """")
    Wend
    
    

End Function

Function ReplaceComma(sNextLine As String, lngStart As Long, lngEnd As Long)
    Dim i As Long
    i = lngStart
    While ((i <= lngEnd) And i > 0)
        i = InStr(i, sNextLine, ",")
        If (i > 0 And i <= lngEnd) Then
            sNextLine = Replace(sNextLine, ",", "-", i, 1)
            i = i + 1
        End If
    Wend

End Function

Private Sub Form_Resize()
    Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
    'FormUnload
    Set SaleSLf15c = Nothing
End Sub

'Private Sub grd_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
'      Grd.Col = 0
'      If Grd.Row >= 1 Then
'         If Grd.Row = 0 Then Grd.Row = 1
'         If Grd.CellPicture = Chkyes.Picture Then
'            Set Grd.CellPicture = Chkno.Picture
'         Else
'            Set Grd.CellPicture = Chkyes.Picture
'         End If
'      End If
'    End If
'
'
'End Sub
'
'
'Private Sub cmdClear_Click()
'    Dim iList As Integer
'    For iList = 1 To Grd.Rows - 1
'        Grd.Col = 0
'        Grd.Row = iList
'        ' Only if the part is checked
'        If Grd.CellPicture = Chkyes.Picture Then
'            Set Grd.CellPicture = Chkno.Picture
'        End If
'    Next
'End Sub


'Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Grd.Col = 0
'   If Grd.Row >= 1 Then
'      If Grd.Row = 0 Then Grd.Row = 1
'      If Grd.CellPicture = Chkyes.Picture Then
'         Set Grd.CellPicture = Chkno.Picture
'      Else
'         Set Grd.CellPicture = Chkyes.Picture
'      End If
'   End If
'End Sub


Private Function CheckForCustomerPO(ByVal strCustomer As String, ByVal strPONum As String) As Byte
   On Error GoTo modErr1
   Dim RdoCpo As ADODB.Recordset
   If Trim(strPONum) = "" Then
      CheckForCustomerPO = 0
   Else
      sSql = "Qry_GetCustomerPo '" & Compress(strCustomer) _
             & "','" & Trim(strPONum) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCpo, ES_FORWARD)
      If bSqlRows Then
         With RdoCpo
            CheckForCustomerPO = 1
            ClearResultSet RdoCpo
         End With
      End If
   End If
   Set RdoCpo = Nothing
   Exit Function
   
modErr1:
   sProcName = "CheckForCustomerPO"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   CheckForCustomerPO = 0
   DoModuleErrors MdiSect.ActiveForm
   
End Function




Private Function MakeAddress(strShipName2 As String, strStreet As String, strStreetSup1 As String, _
                  strStreetSup2 As String, strCity As String, strRegionCode As String, _
                  strPostalCode As String, ByRef strNewAddress As String)

   strNewAddress = ""
   
   ' MM not needed
   'If (strShipName2 <> "") Then strNewAddress = strNewAddress & strShipName2 & vbCrLf
   If (strStreet <> "") Then strNewAddress = strNewAddress & strStreet & vbCrLf
   If (strStreetSup1 <> "") Then strNewAddress = strNewAddress & strStreetSup1 & vbCrLf
   If (strStreetSup2 <> "") Then strNewAddress = strNewAddress & strStreetSup2 & vbCrLf
   
   ' moved Region ==> shiped
   'If (strRegionCode <> "") Then strNewAddress = strNewAddress & strRegionCode & vbCrLf
   If (strCity <> "") Then strNewAddress = strNewAddress & strCity
   
   If (strPostalCode <> "") Then
      If (strRegionCode <> "") Then
         strNewAddress = strNewAddress & ", " & IIf((strRegionCode <> ""), strRegionCode, "") & " - " & strPostalCode
      Else
         strNewAddress = strNewAddress & " - " & strPostalCode
      End If
   End If

End Function


Private Sub GetSJAccounts()
   Dim rdoJrn As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   If sJournalID = "" Then
      bGoodAct = True
      Exit Sub
   End If
   
   On Error GoTo DiaErr1
   sSql = "SELECT COREF,COSJARACCT,COSJNFRTACCT," _
          & "COSJTFRTACCT,COSJTAXACCT FROM ComnTable WHERE " _
          & "COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      With rdoJrn
         ' A/R
         sCOSjARAcct = "" & Trim(.Fields(1))
         If sCOSjARAcct = "" Then b = 1
         ' NonTaxable freight
         sCOSjNFRTAcct = "" & Trim(.Fields(2))
         If sCOSjNFRTAcct = "" Then b = 1
         ' Taxable freight
         sCOSjTFRTAcct = "" & Trim(.Fields(3))
         If sCOSjTFRTAcct = "" Then b = 1
         ' Sales tax
         sCOSjTaxAcct = "" & Trim(.Fields(4))
         If sCOSjTaxAcct = "" Then b = 1
         .Cancel
      End With
   End If
   If b = 1 Then
      bGoodAct = False
      '        lblJrn.Visible = True
   Else
      bGoodAct = True
      '        lblJrn.Visible = False
   End If
   Set rdoJrn = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getsjacco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function GetNextTransaction(sJrnlId As String) As Long
   Dim RdoTrn As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT MAX(DCTRAN) FROM JritTable WHERE DCHEAD='" _
          & Trim(sJrnlId) & "'"
bSqlRows = clsADOCon.GetDataSet(sSql, RdoTrn, ES_FORWARD)
   If bSqlRows Then
      With RdoTrn
         If Not IsNull(.Fields(0)) Then
            GetNextTransaction = (.Fields(0)) + 1
         Else
            GetNextTransaction = 1
         End If
         .Cancel
      End With
   Else
      GetNextTransaction = 1
   End If
   Exit Function
modErr1:
   On Error Resume Next
   sProcName = "getnexttrans"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function

Private Sub GetPartBnO(sPart, nRate, sCode, sState, sType)
   ' Get B&O tax codes from part
   ' Retail takes precidence over wholesale
   
   Dim rdoTx1 As ADODB.Recordset
   Dim rdoTx2 As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,PartTable WHERE " _
          & "PABORTAX = TAXREF AND TAXTYPE = 0 AND PARTREF = '" & sPart & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx1)
   If bSqlRows Then
      With rdoTx1
         nRate = !TAXRATE
         sCode = "" & Trim(!taxCode)
         sState = "" & Trim(!taxState)
         sType = "R"
         .Cancel
      End With
   Else
      sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,PartTable WHERE " _
             & "PABOWTAX = TAXREF AND TAXTYPE = 0 AND PARTREF = '" & sPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx2)
      If bSqlRows Then
         With rdoTx2
            nRate = !TAXRATE
            sCode = "" & Trim(!taxCode)
            sState = "" & Trim(!taxState)
            sType = "W"
            .Cancel
         End With
      End If
   End If
   
   Set rdoTx1 = Nothing
   Set rdoTx2 = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpartbno"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Sub

Private Sub GetSalesTaxInfo( _
                           sCust As String, _
                           nRate As Currency, _
                           sCode As String, _
                           sState As String, _
                           sAccount As String)
   
   On Error GoTo DiaErr1
   
   ' Load tax from customer.
   Dim RdoTax As ADODB.Recordset
   sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE,TAXACCT FROM CustTable INNER JOIN " _
          & "TxcdTable ON CustTable.CUTAXCODE = TxcdTable.TAXREF " _
          & "WHERE CUREF = '" & sCust & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTax)
   If bSqlRows Then
      With RdoTax
         nRate = !TAXRATE
         sCode = "" & Trim(!taxCode)
         sState = "" & Trim(!taxState)
         sAccount = "" & Trim(!TAXACCT)
         .Cancel
      End With
   End If
   Set RdoTax = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getsaletaxinfo"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

Private Sub GetCustBnO(sCust, nRate, sCode, sState, sType)
   ' Get B&O tax codes from customer
   ' Retail takes precidence over wholesale
   
   Dim rdoTx1 As ADODB.Recordset
   Dim rdoTx2 As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,CustTable " _
          & "WHERE CUBORTAXCODE = TAXREF AND CUREF = '" & sCust _
          & "' AND TAXTYPE = 0"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx1)
   If bSqlRows Then
      With rdoTx1
         nRate = !TAXRATE
         sCode = "" & Trim(!taxCode)
         sState = "" & Trim(!taxState)
         sType = "R"
         .Cancel
      End With
   Else
      sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,CustTable " _
             & "WHERE CUBORTAXCODE = TAXREF AND CUREF = '" & sCust _
             & "' AND TAXTYPE = 0"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx2)
      If bSqlRows Then
         With rdoTx2
            nRate = !TAXRATE
            sCode = "" & Trim(!taxCode)
            sState = "" & Trim(!taxState)
            sType = "W"
            .Cancel
         End With
      End If
   End If
   
   Set rdoTx1 = Nothing
   Set rdoTx2 = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcustbno"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub


Private Function GetNextInvoice(strFullPSNum As String) As Long
   
   Dim bDup As Boolean
   
   Dim inv As New ClassARInvoice
   lNextInv = inv.GetNextInvoiceNumber

      
   Dim strPSNum As String
   strPSNum = Mid$(CStr(strFullPSNum), 3, Len(strFullPSNum))
   If (strPSNum <> "") Then
      lNextInv = Val(strPSNum)
   End If

   ' Validate the Invoice number
   If (Trim(lNextInv) <> "") Then
      Dim iCanceled As Integer
      
      iCanceled = 0
      bDup = inv.DuplicateInvNumber(CLng(lNextInv), iCanceled)
      
      If ((bDup = True) And (iCanceled = 0)) Then
         ' if the Inv PS is same then...get the next invoice from the invoice pool.
         lNextInv = inv.GetNextInvoiceNumber
         lNextInv = Format(lNextInv, "000000")
         MsgBox "Invoice number exists for PS Number " & strFullPSNum & ".Using the New Invoice number is " & lNextInv & ".", vbInformation, Caption
      End If
   End If
   
   GetNextInvoice = lNextInv

End Function


Private Function GetPartInvoiceAccounts(SPartRef As String, iLevel As Integer, sCode As String, _
                                       Optional sREVAccount As String, _
                                       Optional sDisAccount As String, _
                                       Optional sCGSMaterialAccount As String, _
                                       Optional sCGSLaborAccount As String, _
                                       Optional sCGSExpAccount As String, _
                                       Optional sCGSOhAccount As String, _
                                       Optional sInvMaterialAccount As String, _
                                       Optional sInvLaborAccount As String, _
                                       Optional sInvExpAccount As String, _
                                       Optional sInvOhAccount As String) As Boolean
   
   Dim rdoAct As ADODB.Recordset
   On Error GoTo modErr1
   
   'Part
   GetPartInvoiceAccounts = True
   SPartRef = Compress(SPartRef)
   
   sSql = "SELECT PACGSMATACCT,PACGSLABACCT,PACGSEXPACCT,PACGSOHDACCT," _
          & "PAINVMATACCT,PAINVLABACCT,PAINVEXPACCT,PAINVOHDACCT," _
          & "PAREVACCT,PADISACCT FROM PartTable WHERE PARTREF='" & SPartRef & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   
   If bSqlRows Then
      With rdoAct
         
         sREVAccount = "" & Trim(!PAREVACCT)
         sDisAccount = "" & Trim(!PADISACCT)
         
         sCGSMaterialAccount = "" & Trim(!PACGSMATACCT)
         sCGSLaborAccount = "" & Trim(!PACGSLABACCT)
         sCGSExpAccount = "" & Trim(!PACGSEXPACCT)
         sCGSOhAccount = "" & Trim(!PACGSOHDACCT)
         
         sInvMaterialAccount = "" & Trim(!PAINVMATACCT)
         sInvLaborAccount = "" & Trim(!PAINVLABACCT)
         sInvExpAccount = "" & Trim(!PAINVEXPACCT)
         sInvOhAccount = "" & Trim(!PAINVOHDACCT)
         
         .Cancel
      End With
   End If
   
   ' Now check the accounts, if any are blank then fill then from the
   ' product code
   sCode = Compress(sCode)
   
   sSql = "SELECT PCCGSMATACCT,PCCGSLABACCT,PCCGSEXPACCT,PCCGSOHDACCT," _
          & "PCINVMATACCT,PCINVLABACCT,PCINVEXPACCT,PCINVOHDACCT," _
          & "PCREVACCT,PCDISCACCT FROM PcodTable WHERE PCREF='" & sCode & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   
   If bSqlRows Then
      With rdoAct
         
         If sREVAccount = "" Then sREVAccount = "" & Trim(!PCREVACCT)
         If sDisAccount = "" Then sDisAccount = "" & Trim(!PCDISCACCT)
         
         If sCGSMaterialAccount = "" Then sCGSMaterialAccount = "" & Trim(!PCCGSMATACCT)
         If sCGSLaborAccount = "" Then sCGSLaborAccount = "" & Trim(!PCCGSLABACCT)
         If sCGSExpAccount = "" Then sCGSExpAccount = "" & Trim(!PCCGSEXPACCT)
         If sCGSOhAccount = "" Then sCGSOhAccount = "" & Trim(!PCCGSOHDACCT)
         
         If sInvMaterialAccount = "" Then sInvMaterialAccount = "" & Trim(!PCINVMATACCT)
         If sInvLaborAccount = "" Then sInvLaborAccount = "" & Trim(!PCINVLABACCT)
         If sInvExpAccount = "" Then sInvExpAccount = "" & Trim(!PCINVEXPACCT)
         If sInvOhAccount = "" Then sInvOhAccount = "" & Trim(!PCINVOHDACCT)
         
         .Cancel
      End With
   End If
   
   ' Last check the company setup and fill any accounts that are still empty.
   sSql = "SELECT COREVACCT" & Trim(str(iLevel)) & "," _
          & "COAPDISCACCT," _
          & "COCGSMATACCT" & Trim(str(iLevel)) & "," _
          & "COCGSLABACCT" & Trim(str(iLevel)) & "," _
          & "COCGSEXPACCT" & Trim(str(iLevel)) & "," _
          & "COCGSOHDACCT" & Trim(str(iLevel)) & "," _
          & "COINVMATACCT" & Trim(str(iLevel)) & "," _
          & "COINVLABACCT" & Trim(str(iLevel)) & "," _
          & "COINVEXPACCT" & Trim(str(iLevel)) & "," _
          & "COINVOHDACCT" & Trim(str(iLevel)) & " FROM " _
          & "ComnTable WHERE COREF=1"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   
   If bSqlRows Then
      With rdoAct
         
         If sREVAccount = "" Then sREVAccount = "" & Trim(.Fields(0))
         If sDisAccount = "" Then sDisAccount = "" & Trim(.Fields(1))
         
         If sCGSMaterialAccount = "" Then sCGSMaterialAccount = "" & Trim(.Fields(2))
         If sCGSLaborAccount = "" Then sCGSLaborAccount = "" & Trim(.Fields(3))
         If sCGSExpAccount = "" Then sCGSExpAccount = "" & Trim(.Fields(4))
         If sCGSOhAccount = "" Then sCGSOhAccount = "" & Trim(.Fields(5))
         
         If sInvMaterialAccount = "" Then sInvMaterialAccount = "" & Trim(.Fields(6))
         If sInvLaborAccount = "" Then sInvLaborAccount = "" & Trim(.Fields(7))
         If sInvExpAccount = "" Then sInvExpAccount = "" & Trim(.Fields(8))
         If sInvOhAccount = "" Then sInvOhAccount = "" & Trim(.Fields(9))
         
         .Cancel
      End With
   End If
   
   
   Set rdoAct = Nothing
   Exit Function
   
modErr1:
   sProcName = "GetPartInvoiceAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   MsgBox CurrError.Number & " " & CurrError.Description
End Function

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Grd.Col = 0
   If Grd.Row >= 1 Then
      If Grd.Row = 0 Then Grd.Row = 1
      If Grd.CellPicture = Chkyes.Picture Then
         Set Grd.CellPicture = Chkno.Picture
         Grd.Col = 3
         cTotInvAmtSel = cTotInvAmtSel - Val(Trim(Grd.Text))
         
      Else
         Set Grd.CellPicture = Chkyes.Picture
         
         Grd.Col = 3
         cTotInvAmtSel = cTotInvAmtSel + Val(Trim(Grd.Text))
      End If
      lblInvTotal = Format(cTotInvAmtSel, "#####0.00")
   End If
End Sub
