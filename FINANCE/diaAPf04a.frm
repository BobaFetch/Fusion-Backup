VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form diaAPf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Void AP Checks "
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEnd 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Tag             =   "1"
      Top             =   1920
      Width           =   1000
   End
   Begin VB.TextBox txtBeg 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   14
      Tag             =   "1"
      Top             =   1560
      Width           =   1000
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   6120
      TabIndex        =   1
      Tag             =   "4"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdcnl 
      Caption         =   "&Reselect"
      Height          =   315
      Left            =   6840
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7320
      Top             =   120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6375
      FormDesignWidth =   8820
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Click On Check Number To Select A Check"
      Top             =   3000
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Height          =   315
      Left            =   7800
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   8655
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      ToolTipText     =   "Nicknames"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdVoid 
      Caption         =   "&Void"
      Height          =   315
      Left            =   7800
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   240
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   8
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
      PictureUp       =   "diaAPf04a.frx":0000
      PictureDn       =   "diaAPf04a.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   18
      Top             =   1200
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Check #"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Check #"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Void Date"
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   13
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4320
      TabIndex        =   12
      Top             =   480
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Checks Found"
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.Image imgInc 
      Height          =   180
      Left            =   3960
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgdInc 
      Height          =   180
      Left            =   3600
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "diaAPf04a"
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
' diaAPf04a - Void AP Checks
'
' Notes:
'
' Created: (nth)
' Revisons:
' 04/01/03 (nth) Fixed errors per LML and added graphic check box to grid.
' 05/14/03 (nth) Fixed number of checks not excluding void checks.
' 05/14/03 (nth) Fixed issue with VIPAY not being correctly deducted.
' 10/01/03 (nth) Added void date per (ENTSYS).
' 02/13/04 (jcw) Added fixed columns to conform to Design Standard.
' 03/25/04 (nth) Fixed error with voiding check that has no invoice per (JEVINT).
' 03/31/04 (nth) Changed Voided Check XXXXXX to Void XXXXXX
' 05/12/04 (nth) Correctly reverse credit memos.
' 08/10/04 (nth) Corrected not voiding alpha numeric check numbers.
' 09/26/04 (nth) Added compress to getchecks query.
' 11/24/04 (nth) Corrected voiding checks with discounts.
' 01/19/05 (nth) Corrected credit memo summary reversal.
' 02/09/05 (nth) Corrected voiding check type 3 not inserting journal entries.
'
'*************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte

Dim sXC As String
Dim sCC As String
Dim sMsg As String
Dim iInvoices As Integer
Dim cCheckTotal As Currency

Dim sInvoices(100) As String
Dim vInvoice(100, 5) As Variant

Dim sChecks() As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub FillCombo()
   Dim RdoVnd As ADODB.Recordset
   On Error GoTo DiaErr1
   
   cmbVnd.Clear
   
   sSql = "SELECT DISTINCT VENICKNAME FROM VndrTable INNER JOIN " _
          & "ChksTable ON VEREF = CHKVENDOR " _
          & "WHERE CHKVOIDDATE IS NULL ORDER BY VENICKNAME"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd)
   
   If bSqlRows Then
      With RdoVnd
         While Not .EOF
            AddComboStr cmbVnd.hWnd, "" & Trim(.Fields(0))
            .MoveNext
         Wend
         .Cancel
      End With
   Else
      'MsgBox "No Vendors Found.", vbInformation, Caption
   End If
   
   AddComboStr cmbVnd.hWnd, "ALL"
   
   Set RdoVnd = Nothing
   If cmbVnd.ListCount > 0 Then
      cmbVnd.ListIndex = 0
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetChecks()
   Dim RdoChk As ADODB.Recordset
   Dim sGridItem As String
   Dim sType As String * 1
   Dim i As Integer
   Dim strBeg As String
   Dim strEnd As String
   
   On Error GoTo DiaErr1
   
   Grid1.Clear
   Grid1.Rows = 1
   SetUpGrid
   
'   sSql = "SELECT CHKNUMBER,CHKACCT, CHKPOSTDATE,CHKACTUALDATE," _
'          & "CHKAMOUNT,CHKMEMO,CHKPRINTED " _
'          & "FROM ChksTable WHERE CHKVOID = 0 " _
'          & "AND CHKVENDOR = '" & Compress(cmbVnd) & "' " _
'          & "AND CHKCLEARDATE IS NULL"
   
   If (Compress(cmbVnd) <> "") Then
      sSql = "SELECT CHKNUMBER,CHKACCT, CHKPOSTDATE,CHKACTUALDATE," _
             & "CHKAMOUNT,CHKMEMO,CHKPRINTED " _
             & "FROM ChksTable WHERE CHKVOID = 0 " _
             & "AND CHKVENDOR = '" & Compress(cmbVnd) & "' " _
             & "AND CHKCLEARDATE IS NULL"
   
   Else
      strBeg = txtBeg.Text
      strEnd = txtEnd.Text
      
      If (strBeg = "" Or strEnd = "") Then
         MsgBox "Enter the Begining and Ending Check Number.", vbInformation, Caption
         Exit Sub
      End If
      
      
      sSql = "SELECT CHKNUMBER,CHKACCT, CHKPOSTDATE,CHKACTUALDATE," _
             & "CHKAMOUNT,CHKMEMO,CHKPRINTED " _
             & "FROM ChksTable WHERE CHKVOID = 0 " _
             & "AND CHKNUMBER BETWEEN '" & strBeg & "' AND '" & strEnd & "' " _
             & "AND CHKCLEARDATE IS NULL"
   
   End If
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk, ES_FORWARD)
   
   If bSqlRows Then
      i = 1
      With RdoChk
         While Not .EOF
            If !CHKPRINTED = 0 Then
               sType = "X"
            Else
               sType = "C"
            End If
            
            sGridItem = Chr(9) _
                        & Trim(!CHKNUMBER) & Chr(9) _
                        & Trim(!CHKACCT) & Chr(9) _
                        & sType & Chr(9) _
                        & Format(!CHKPOSTDATE, "mm/dd/yy") & Chr(9) _
                        & Format(!CHKACTUALDATE, "mm/dd/yy") & Chr(9) _
                        & Format(!CHKAMOUNT, "0.00") & Chr(9) _
                        & "" & Trim(!chkMemo)
            Grid1.AddItem (sGridItem)
            Grid1.Row = i
            Grid1.Col = 0
            Grid1.CellPictureAlignment = flexAlignCenterCenter
            Set Grid1.CellPicture = imgdInc
            .MoveNext
            i = i + 1
         Wend
         .Cancel
      End With
   Else
      FillCombo
      Grid1.enabled = False
      cmdVoid.enabled = False
      cmdSel.enabled = True
      cmdcnl.enabled = False
      cmbVnd.enabled = True
      cmbVnd.SetFocus
   End If
   Set RdoChk = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getchecks"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub VoidCheck()
   Dim cChkAmt As Currency
   Dim RdoChk As ADODB.Recordset
  ' Dim RdoInv As ADODB.Recordset
   Dim sJournalID As String
   Dim sChkNum As String
   Dim sChkAcct As String
   Dim sChkVnd As String
   Dim sChkType As String * 1
   Dim iTrans As Integer
   Dim iRef As Integer
   Dim cDebit As Currency
   Dim cCredit As Currency
   Dim iResponse As Integer
   Dim i As Integer
   Dim sPlural As String * 1
   Dim cPaid As Currency
   Dim bVoidedSomething As Byte
   Dim voidedCount As Integer
   Dim postDate As String
   
   
   On Error GoTo DiaErr1
   
   sChkVnd = Compress(cmbVnd)
   
   iResponse = MsgBox("Void Selected Check(s) ?", ES_YESQUESTION, Caption)
   If iResponse = vbNo Then
      Exit Sub
   End If
   
   'On Error Resume Next
   MouseCursor 13
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   postDate = txtDte
   
   For i = 1 To Grid1.Rows - 1
      
      Grid1.Row = i
      Grid1.Col = 0
      
      If Grid1.CellPicture = imgInc Then
         
         With Grid1
            .Col = 1
            sChkNum = Trim(Grid1)
            .Col = 2
            sChkAcct = Trim(Grid1)
            .Col = 3
            sChkType = Trim(Grid1)
            
            '.Col = 3
            'postDate = Trim(Grid1)
            .Col = 6
            cChkAmt = Val(Grid1)
         End With
         
         sJournalID = ""
         sCC = ""
         sXC = ""
         If sChkType = "C" Then
            If sCC = "" Then
               sCC = GetOpenJournal("CC", postDate)
               If sCC = "" Then
                  sMsg = "Cannot Void Check #" & sChkNum & vbCrLf & _
                         "No Open Computer Check Journal for " & postDate & " Found."
               End If
            End If
            sJournalID = sCC
         Else
            If sXC = "" Then
               sXC = GetOpenJournal("XC", postDate)
               If sXC = "" Then
                  sMsg = "Cannot Void Check #" & sChkNum & vbCrLf & _
                         "No Open External Check Journal  for " & postDate & " Found."
               End If
            End If
            sJournalID = sXC
         End If
         
         If sJournalID = "" Then
            MsgBox sMsg, vbInformation, Caption
            Exit For
         End If
         
         If (Compress(cmbVnd) = "") Then
            sChkVnd = GetVndForCheckNumber(sChkNum)
            If (sChkVnd = "") Then
               MsgBox "The Check Number - " & sChkNum & " doesn't have Vendor name.", vbInformation, Caption
               Exit For
            End If
         End If
         
         sSql = "UPDATE ChksTable SET " _
                & "CHKVOID = 1, " _
                & "CHKVOIDDATE = '" & txtDte & "' " _
                & "WHERE CHKNUMBER = '" & sChkNum & "' AND " _
                & "CHKACCT = '" & sChkAcct & "' AND CHKVENDOR = '" & sChkVnd & "'"
         clsADOCon.ExecuteSQL sSql
         
         'sSql = "SELECT DCVENDOR,DCVENDORINV,DCDEBIT,DCCREDIT,DCTRAN,DCREF," _
         '    & "DCACCTNO,CHKACCT FROM JritTable,ChksTable,VihdTable WHERE " _
         '    & "DCVENDOR = VIVENDOR AND DCVENDORINV = INVO AND DC DCCHECKNO = " _
         '    & "'" & sChkNum & "' AND DCVENDOR = '" & sChkVnd & "' ORDER BY DCTRAN,DCREF"
         
         sSql = "SELECT DCVENDOR,DCVENDORINV,DCDEBIT,DCCREDIT,DCACCTNO,DCTRAN," & vbCrLf _
            & "DCREF,CHKACCT,VITYPE,VIDUE,VIDISCOUNT,DCTYPE" & vbCrLf _
            & "FROM ChksTable" & vbCrLf _
            & "INNER JOIN JritTable ON CHKNUMBER = DCCHECKNO AND CHKVENDOR = DCVENDOR" & vbCrLf _
            & "LEFT JOIN VihdTable ON DCVENDOR = VIVENDOR AND DCVENDORINV = VINO" & vbCrLf _
            & "WHERE DCVENDOR = '" & sChkVnd & "' AND DCCHECKNO = '" & sChkNum & "'" & vbCrLf _
            & " AND CHKACCT = '" & sChkAcct & "'" & vbCrLf _
            & "ORDER BY DCTRAN, DCREF"
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk, ES_STATIC)
         
         'calculate total debit, amt paid, and discount taken
         'this logic assumes that is the order of the debits and credits
         'simultaneously create reversing debits and credits
         Dim totalDebits As Currency, amountPaid As Currency, discountTaken As Currency
         If bSqlRows Then
            With RdoChk
               iRef = 0
               iTrans = GetNextTransaction(sJournalID)
               While Not .EOF
                  ' Make adjusting (reversing) entries to open journal
                  iRef = iRef + 1
                  sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT," _
                     & "DCCREDIT,DCACCTNO,DCDATE,DCDESC,DCTYPE) " & vbCrLf _
                     & "VALUES('" _
                     & sJournalID & "'," _
                     & iTrans & "," _
                     & iRef & "," _
                     & !DCCREDIT & "," _
                     & !DCDEBIT & ",'" _
                     & Trim(!DCACCTNO) & "','" _
                     & postDate & "','" _
                     & "Void " & sChkNum & "'," _
                     & !DCTYPE & ")"
                  clsADOCon.ExecuteSQL sSql
                  
'                  'reverse ap debits (regular inv) or credits (cm)
'                  cPaid = !DCCREDIT - !DCDEBIT
'
'                  'adjust balance due on credit
'                  If !VIDUE < 0 And cPaid < 0 Then
'                     sSql = "UPDATE VihdTable SET VIPAY = VIPAY + " & Abs(cPaid)
'
'                  'adjust balance due on debit (dm & regular invoice)
'                  ElseIf !VIDUE > 0 And cPaid > 0 Then
'                     sSql = "UPDATE VihdTable SET VIPAY = VIPAY - " & cPaid
'                  Else
'                     sSql = ""
'                  End If
'
'                  If sSql <> "" Then
'                     sSql = sSql & ",VIPIF=0,VICHECKNO=0"
'
'                     'If cPaid = Abs(!VIDUE - !VIDISCOUNT) Then
'                        sSql = sSql & ",VIDISCOUNT=0"
'                     'End If
'
'                     sSql = sSql & " WHERE VINO = '" & Trim(!DCVENDORINV) _
'                            & "' AND VIVENDOR = '" & Trim(!DCVENDOR) & "'"
'Debug.Print "** acct=" & !DCACCTNO & ": debit=" & !DCDEBIT & " credit=" & !DCCREDIT & vbCrLf & sSql
'
'                     clsAdoCon.ExecuteSQL sSQL
'                  End If
                  sSql = ""
                  Select Case !DCREF
                  Case 2:
                     sSql = "UPDATE VihdTable" & vbCrLf _
                        & "SET VIPIF = 0, VICHECKNO = 0, VICHKACCT = 0, VIPAY = VIPAY + " & !DCDEBIT - !DCCREDIT
                  Case 3:
                     sSql = "UPDATE VihdTable" & vbCrLf _
                        & "SET VIDISCOUNT = VIDISCOUNT + " & !DCDEBIT - !DCCREDIT
                  End Select
                  
                  If sSql <> "" Then
'                     sSql = sSql & vbCrLf & "WHERE VINO = '" & Trim(!DCVENDORINV) & "'" & vbCrLf _
'                        & "AND VIVENDOR = '" & Trim(!DCVENDOR) & "'" & vbCrLf _
'                        & "AND VICHECKNO = '" & sChkNum & "'"
                     sSql = sSql & vbCrLf & "WHERE VINO = '" & Trim(!DCVENDORINV) & "'" & vbCrLf _
                        & "AND VIVENDOR = '" & Trim(!DCVENDOR) & "'"
                     clsADOCon.ExecuteSQL sSql
                  End If
                  .MoveNext
                  bVoidedSomething = True
               Wend
               voidedCount = voidedCount + 1
               .Cancel
            End With
         End If
      End If
   Next
   
   Set RdoChk = Nothing
   MouseCursor 0
   
   If clsADOCon.ADOErrNum = 0 And voidedCount > 0 Then
      clsADOCon.CommitTrans
      If bVoidedSomething Then
         SysMsg voidedCount & " Check(s) Successfully Voided.", True
      End If
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      MsgBox "Could Not Successfully Void Check(s)." _
         , vbInformation, Caption
   End If
   
   
   Exit Sub
   
DiaErr1:
   sProcName = "VoidChecks"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmbVnd_Click()
   FindVendor Me
   lblNum = NumberOfChecks(cmbVnd, True)
End Sub

Private Function GetVndForCheckNumber(strChkNum As String) As String
   Dim rdoChkNum As ADODB.Recordset
   Dim sVendor As String
   On Error GoTo DiaErr1
   sSql = "SELECT CHKVENDOR FROM ChksTable WHERE " _
          & "CHKNUMBER='" & strChkNum & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoChkNum)
   If bSqlRows Then
      GetVndForCheckNumber = rdoChkNum!CHKVENDOR
   Else
      GetVndForCheckNumber = ""
   End If
   On Error Resume Next
   rdoChkNum.Close
   Set rdoChkNum = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checkinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function
Private Sub cmbVnd_LostFocus()
   'FindVendor Me
   'lblNum = NumberOfChecks(cmbVnd, True)
   
   If cmbVnd <> "" Then
      FindVendor Me
      lblNum = NumberOfChecks(cmbVnd, True)
      
      txtBeg.Text = ""
      txtBeg.enabled = False
      txtEnd.Text = ""
      txtEnd.enabled = False
      
   Else
      txtBeg.enabled = True
      txtEnd.enabled = True
      lblNum = ""
      lblNum = 0
      txtBeg.SetFocus
      
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCnl_Click()
   cmdVoid.enabled = False
   cmdcnl.enabled = False
   cmdSel.enabled = True
   cmbVnd.enabled = True
   txtDte.enabled = True
   Grid1.Rows = 1
   Grid1.Clear
   SetUpGrid
   'cmbVnd.SetFocus
   If (cmbVnd <> "") Then
      cmbVnd.SetFocus
   Else
      txtBeg.SetFocus
   End If
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdSel_Click()
   GetChecks
   cmdVoid.enabled = True
   cmdcnl.enabled = True
   cmdSel.enabled = False
   txtDte.enabled = False
   Grid1.enabled = True
   cmbVnd.enabled = False
   Grid1.SetFocus
End Sub

Private Sub cmdVoid_Click()
   VoidCheck
   cmdVoid.enabled = True
   cmdSel.enabled = False
   Grid1.SetFocus
   GetChecks
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      MouseCursor 13
      SetUpGrid
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   imgdInc.Picture = Resources.imgdInc.Picture
   imgInc.Picture = Resources.imgInc.Picture
   txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   cmdVoid.enabled = False
   cmdcnl.enabled = False
   bOnLoad = True
End Sub

Private Sub SetUpGrid()
   With Grid1
      .Cols = 8
      .Rows = 1
      .ColWidth(0) = 500
      .ColWidth(1) = 1000
      .ColWidth(2) = 1300
      .ColWidth(3) = 200
      .ColWidth(4) = 1000
      .ColWidth(5) = 1100
      .ColWidth(6) = 1200
      .ColWidth(7) = 3000
      .Row = 0
      .Col = 0
      .Text = "Void?"
      .Col = 1
      .Text = "Check #"
      .Col = 2
      .Text = "Check Acct"
      .Col = 4
      .Text = "Post Date"
      .Col = 5
      .Text = "Actual Date"
      .Col = 6
      .Text = "Amount"
      .Col = 7
      .Text = "Memo"
   End With
End Sub


Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set diaAPf04a = Nothing
   
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub Grid1_Click()
   With Grid1
      If .Row > 0 Then
         .Col = 0
         .Row = .RowSel
         If .CellPicture = imgdInc Then
            Set .CellPicture = imgInc
         Else
            Set .CellPicture = imgdInc
         End If
      End If
   End With
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDte_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
End Sub
